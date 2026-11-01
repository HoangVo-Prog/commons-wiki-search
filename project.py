#!/usr/bin/env python3
"""Configuration-backed per-agent quotas and project paths; no global image quota."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / 'config.json'

REQUIRED_PATHS = (
    'worker_root', 'worker_template', 'root_instructions', 'task_instructions',
    'worker_instructions', 'agent_identity', 'keywords', 'cursor',
    'state', 'progress', 'events', 'workbook', 'claims', 'collection_state',
)


def settings():
    data = json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
    count = data.get('agent_count')
    if type(count) is not int or not 1 <= count <= 99:
        raise ValueError('config.json agent_count must be an integer between 1 and 99')
    quotas = data.get('agent_quota')
    expected = {f'{i:02d}' for i in range(1, count + 1)}
    if not isinstance(quotas, dict) or set(quotas) != expected:
        raise ValueError('config.json agent_quota must map exactly the configured two-digit agent IDs to individual quotas')
    for agent, value in quotas.items():
        if type(value) is not int or value < 0:
            raise ValueError(f'config.json agent_quota[{agent!r}] must be a nonnegative integer')
    paths = data.get('paths')
    if not isinstance(paths, dict) or any(k not in paths for k in REQUIRED_PATHS):
        raise ValueError(f'config.json paths must include: {", ".join(REQUIRED_PATHS)}')
    for key in REQUIRED_PATHS:
        value = paths[key]
        if not isinstance(value, str) or not value.strip() or Path(value).is_absolute() or '..' in Path(value).parts:
            raise ValueError(f'Invalid relative path in config.json paths.{key}')
    if '{agent_id}' not in paths['worker_template']:
        raise ValueError('paths.worker_template must contain {agent_id}')
    return data


def project_path(key):
    return ROOT / settings()['paths'][key]


def valid_id(agent_id):
    a = str(agent_id).zfill(2)
    if not re.fullmatch(r'\d{2}', a) or not (1 <= int(a) <= settings()['agent_count']):
        raise ValueError(f'Invalid agent ID {agent_id!r}: expected 01..{settings()["agent_count"]:02d}')
    return a


def ids():
    return [f'{i:02d}' for i in range(1, settings()['agent_count'] + 1)]


def worker_dir(agent_id):
    a = valid_id(agent_id)
    p = settings()['paths']
    name = p['worker_template'].replace('{agent_id}', a)
    if Path(name).is_absolute() or '..' in Path(name).parts:
        raise ValueError('Invalid worker_template expansion')
    return ROOT / p['worker_root'] / name


def worker_file(agent_id, key):
    if key not in ('worker_instructions', 'agent_identity', 'keywords', 'cursor', 'state', 'progress', 'events'):
        raise ValueError(f'{key} is not a worker path')
    return worker_dir(agent_id) / settings()['paths'][key]


def quota(agent_id):
    return settings()['agent_quota'][valid_id(agent_id)]


def manifest(agent_id=None):
    cfg = settings()
    doc = {
        'project_root': str(ROOT),
        'agent_count': cfg['agent_count'],
        'root_instructions': str(project_path('root_instructions')),
        'task_instructions': str(project_path('task_instructions')),
        'workbook': str(project_path('workbook')),
        'claims': str(project_path('claims')),
        'collection_state': str(project_path('collection_state')),
        'worker_root': str(project_path('worker_root')),
    }
    if agent_id is not None:
        a = valid_id(agent_id)
        doc.update({'agent_id': a, 'agent_quota': quota(a), 'worker_dir': str(worker_dir(a))})
        for key in ('worker_instructions', 'agent_identity', 'keywords', 'cursor', 'state', 'progress', 'events'):
            doc[key] = str(worker_file(a, key))
    return doc


def main():
    cli = argparse.ArgumentParser(description='Show independent per-agent quotas and configured document paths')
    cli.add_argument('command', nargs='?', choices=('show', 'manifest'), default='show')
    cli.add_argument('--agent-id')
    args = cli.parse_args()
    if args.command == 'manifest' and not args.agent_id:
        cli.error('manifest requires --agent-id')
    print(json.dumps(manifest(args.agent_id), indent=2, ensure_ascii=False))


if __name__ == '__main__':
    main()
