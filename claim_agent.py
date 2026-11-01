#!/usr/bin/env python3
"""Atomic independent agent ID claim/reservation. No session-specific ID hardcoding."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import tempfile

import project
import record


def _load():
    file = project.project_path('claims')
    if not file.exists():
        return {}
    data = json.loads(file.read_text(encoding='utf-8'))
    if not isinstance(data, dict):
        raise RuntimeError('Claim registry must be an object')
    return data


def _save(data):
    path = project.project_path('claims')
    fd, tmp = tempfile.mkstemp(prefix='.claim_', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write('\n')
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def cmd_claim(_args):
    with record.excel_lock():
        claims = _load()
        remaining = [i for i in project.ids() if i not in claims]
        if not remaining:
            raise RuntimeError('All configured agent IDs are claimed. Reuse your existing assigned ID; never claim twice in one session.')
        agent = remaining[0]
        claims[agent] = {'claimed_at': record.timestamp(), 'mode': 'autoclaim'}
        _save(claims)
    print(f'AGENT_ID={agent}')
    print(f'MANIFEST_COMMAND=python project.py manifest --agent-id {agent}')


def cmd_reserve(args):
    agent = project.valid_id(args.agent_id)
    with record.excel_lock():
        claims = _load()
        if agent in claims:
            raise RuntimeError(f'Agent {agent} is already claimed. Resume that existing session; do not start another collector with the same ID.')
        claims[agent] = {'claimed_at': record.timestamp(), 'mode': 'preassigned'}
        _save(claims)
    print(f'AGENT_ID={agent}')


def cmd_status(_args):
    with record.excel_lock():
        claims = _load()
    print(json.dumps({'claimed': claims, 'unclaimed': [x for x in project.ids() if x not in claims]}, indent=2))


def main():
    cli = argparse.ArgumentParser(description='Atomically claim a unique configured Codex agent ID')
    sub = cli.add_subparsers(dest='command', required=True)
    c = sub.add_parser('claim', help='Claim next available ID atomically')
    c.set_defaults(fn=cmd_claim)
    r = sub.add_parser('reserve', help='Reserve an explicitly assigned ID for a launcher')
    r.add_argument('--agent-id', required=True)
    r.set_defaults(fn=cmd_reserve)
    s = sub.add_parser('status', help='Read claimed/unclaimed IDs')
    s.set_defaults(fn=cmd_status)
    args = cli.parse_args()
    try:
        args.fn(args)
    except (RuntimeError, ValueError, OSError) as exc:
        cli.exit(2, f'ERROR | {exc}\n')


if __name__ == '__main__':
    main()
