#!/usr/bin/env python3
"""Shared Excel append-only writer with fresh, per-run agent quotas.

XLSX is an OOXML ZIP archive. This script edits only the first worksheet's four
columns while retaining the rest of the workbook. OS locking and atomic replace
make simultaneous worker processes safe on a local disk. Historical workbook
rows remain preserved and count toward only the initial count of each new run.
"""
from __future__ import annotations

import argparse
from collections import Counter
from hashlib import sha256
from contextlib import contextmanager
from datetime import datetime
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import time
import unicodedata
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import project

ROOT = Path(__file__).resolve().parent
EXCEL = project.project_path('workbook')
LOCKFILE = ROOT / '.images.xlsx.lock'
COLS = ('AGENT_ID', 'original_filename', 'original_image_url', 'hrvs_suitability')
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XML = 'http://www.w3.org/XML/1998/namespace'
ET.register_namespace('', NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')


def timestamp() -> str:
    return datetime.now().astimezone().isoformat(timespec='seconds')


def fatal(s: str) -> None:
    print('ERROR | ' + s, file=sys.stderr)
    raise SystemExit(2)


def agent_from_workdir(explicit: str | None) -> str:
    local = Path.cwd() / project.settings()['paths']['agent_identity']
    implicit = local.read_text(encoding='utf-8').strip() if local.is_file() else None
    if explicit is not None:
        requested = explicit.zfill(2)
        if implicit is not None and requested != implicit:
            fatal(f'Current directory is agent {implicit}, not agent {requested}')
        agent = requested
    elif implicit:
        agent = implicit
    else:
        fatal('No AGENT_ID. Pass --agent-id or run from your configured worker directory')
    try:
        agent = project.valid_id(agent)
    except ValueError as exc:
        fatal(str(exc))
    expected = project.worker_file(agent, 'agent_identity')
    if not expected.is_file() or expected.read_text(encoding='utf-8').strip() != agent:
        fatal(f'Missing or inconsistent worker identity file: {expected}')
    return agent


@contextmanager
def excel_lock(timeout=120.0):
    LOCKFILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOCKFILE, 'a+b') as lock:
        lock.seek(0, os.SEEK_END)
        if lock.tell() == 0:
            lock.write(b'0')
            lock.flush()
        lock.seek(0)
        start = time.monotonic()
        if os.name == 'nt':
            import msvcrt
            while True:
                try:
                    lock.seek(0)
                    msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError:
                    if time.monotonic() - start > timeout:
                        fatal('Timed out waiting for Excel writer lock')
                    time.sleep(0.15)
        else:
            import fcntl
            while True:
                try:
                    fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    break
                except BlockingIOError:
                    if time.monotonic() - start > timeout:
                        fatal('Timed out waiting for Excel writer lock')
                    time.sleep(0.15)
        try:
            yield
        finally:
            if os.name == 'nt':
                lock.seek(0)
                msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def _read_shared(archive: ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in archive.namelist():
        return []
    tree = ET.fromstring(archive.read('xl/sharedStrings.xml'))
    return [''.join((elem.text or '') for elem in si.iter(f'{{{NS}}}t'))
            for si in tree.findall(f'{{{NS}}}si')]


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    typ = cell.get('t', '')
    if typ == 'inlineStr':
        return ''.join(e.text or '' for e in cell.iter(f'{{{NS}}}t'))
    val = cell.find(f'{{{NS}}}v')
    s = val.text if val is not None and val.text is not None else ''
    if typ == 's' and s:
        return shared[int(s)]
    return s


def load_sheet():
    if not EXCEL.exists():
        fatal('images.xlsx missing. Restore the initialized workbook from the ZIP.')
    with ZipFile(EXCEL) as z:
        xml = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        shared = _read_shared(z)
    data = xml.find(f'{{{NS}}}sheetData')
    if data is None:
        fatal('Excel sheetData missing: images.xlsx is invalid')
    rows: list[list[str]] = []
    for row in data.findall(f'{{{NS}}}row'):
        cells = {re.match(r'^[A-Z]+', c.get('r', '')).group(): _cell_value(c, shared)
                 for c in row.findall(f'{{{NS}}}c') if re.match(r'^[A-Z]+', c.get('r', ''))}
        rows.append([cells.get(col, '') for col in 'ABCD'])
    if not rows or tuple(rows[0]) != COLS:
        fatal('Workbook headers are not exactly the four expected columns')
    return xml, data, rows


def canonical_url(url: str) -> str:
    parts = urlsplit(url.strip())
    if parts.scheme.lower() not in ('http', 'https') or parts.hostname is None:
        raise ValueError('original_image_url must be an absolute HTTP(S) URL')
    if parts.hostname.lower() != 'upload.wikimedia.org':
        raise ValueError('Original image URL must be on upload.wikimedia.org')
    path = unicodedata.normalize('NFC', unquote(parts.path))
    if not path.startswith('/wikipedia/commons/') or '/thumb/' in path or '/archive/' in path:
        raise ValueError('URL must target a full-resolution Commons file, not thumbnail/archive')
    if not re.search(r'\.(?:jpg|jpeg|png|webp|tif|tiff)$', path, re.I):
        raise ValueError('URL must point to a supported photograph format')
    return 'upload.wikimedia.org' + path


def canonical_filename(name: str) -> str:
    return unicodedata.normalize('NFC', unquote(name.strip())).replace('_', ' ').casefold()


def safe_text(s: str, max_len: int, field: str) -> str:
    s = ' '.join(s.strip().split())
    if not s or len(s) > max_len:
        raise ValueError(f'{field} must be nonempty, at most {max_len} chars')
    if any(ord(c) < 32 for c in s):
        raise ValueError(f'{field} contains control characters')
    return s


def _save_appended_row(xml: ET.Element, data: ET.Element, existing_count: int,
                       entry: list[str]):
    row_idx = existing_count + 1
    row_el = ET.SubElement(data, f'{{{NS}}}row', {'r': str(row_idx)})
    for col, value in zip('ABCD', entry):
        c = ET.SubElement(row_el, f'{{{NS}}}c', {'r': f'{col}{row_idx}', 't': 'inlineStr'})
        inline = ET.SubElement(c, f'{{{NS}}}is')
        t = ET.SubElement(inline, f'{{{NS}}}t')
        t.text = value
        if value.startswith(' ') or value.endswith(' '):
            t.set(f'{{{XML}}}space', 'preserve')
    dimension = xml.find(f'{{{NS}}}dimension')
    if dimension is not None:
        dimension.set('ref', f'A1:D{row_idx}')
    filt = xml.find(f'{{{NS}}}autoFilter')
    if filt is not None:
        filt.set('ref', f'A1:D{row_idx}')
    new_sheet = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
    fd, tmp_name = tempfile.mkstemp(prefix='.images_', suffix='.xlsx', dir=EXCEL.parent)
    os.close(fd)
    try:
        with ZipFile(EXCEL, 'r') as src, ZipFile(tmp_name, 'w') as dst:
            for member in src.infolist():
                payload = new_sheet if member.filename == 'xl/worksheets/sheet1.xml' else src.read(member.filename)
                dst.writestr(member, payload)
        os.replace(tmp_name, EXCEL)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def _worker_dir(agent: str) -> Path:
    return project.worker_dir(agent)


def _atomic_text(path: Path, value: str):
    tmp = path.with_name(path.name + '.tmp')
    with tmp.open('w', encoding='utf-8', newline='\n') as f:
        f.write(value)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def _historical_digest(rows: list[list[str]]) -> str:
    """Anchor the original workbook prefix to detect out-of-band row edits/reorders."""
    digest = sha256()
    for row in rows:
        digest.update(json.dumps(row, ensure_ascii=False, separators=(',', ':')).encode('utf-8'))
        digest.update(b'\n')
    return digest.hexdigest()


def _count_agent_rows(rows: list[list[str]]) -> Counter:
    """Count workbook records by agent ID, including old numeric IDs like '1'."""
    result = Counter()
    valid = set(project.ids())
    for r in rows:
        agent = str(r[0]).strip()
        if agent.isascii() and agent.isdecimal() and len(agent) <= 2:
            normalized = agent.zfill(2)
            agent = normalized if normalized in valid else agent
        result[agent] += 1
    return result


def _snapshot_run(rows: list[list[str]], run_number: int, initialized_at: str | None = None,
                  baseline_row_count: int | None = None, baseline_sha256: str | None = None) -> dict:
    """Capture a complete immutable baseline and quota/target for EVERY agent."""
    baseline = len(rows) - 1 if baseline_row_count is None else baseline_row_count
    historical = rows[1:baseline + 1]
    counts = _count_agent_rows(historical)
    agents = {}
    for agent in project.ids():
        initial = counts[agent]
        quota = project.quota(agent)
        agents[agent] = {'initial_count': initial, 'agent_quota': quota,
                         'target_count': initial + quota}
    return {
        'run_id': f'run-{run_number:06d}',
        'initialized_at': initialized_at or timestamp(),
        'baseline_row_count': baseline,
        'baseline_sha256': baseline_sha256 or _historical_digest(rows[:baseline + 1]),
        'agents': agents,
    }


def _save_collection_state(state: dict) -> None:
    _atomic_text(project.project_path('collection_state'),
                 json.dumps(state, ensure_ascii=False, indent=2) + '\n')


def _active_run(state: dict) -> dict:
    matches = [r for r in state['runs'] if r['run_id'] == state['active_run_id']]
    if len(matches) != 1:
        raise RuntimeError('Invalid collection_state.json: active_run_id must identify one run')
    return matches[0]


def _validate_run(run: dict, rows: list[list[str]]) -> None:
    baseline = run.get('baseline_row_count')
    if (type(baseline) is not int or baseline < 0 or baseline > len(rows) - 1
            or run.get('baseline_sha256') != _historical_digest(rows[:baseline + 1])):
        raise RuntimeError('Existing workbook rows changed or baseline is invalid. '
                           'Restore the original matching images.xlsx and collection_state.json.')
    agents = run.get('agents')
    if not isinstance(agents, dict) or set(agents) != set(project.ids()):
        raise RuntimeError(f'Invalid run agent snapshot in {run.get("run_id")}.')
    actual_initial = _count_agent_rows(rows[1:baseline + 1])
    for agent, snapshot in agents.items():
        if not isinstance(snapshot, dict):
            raise RuntimeError(f'Invalid snapshot for agent {agent}')
        initial, quota, target = (snapshot.get('initial_count'), snapshot.get('agent_quota'),
                                  snapshot.get('target_count'))
        if (type(initial) is not int or type(quota) is not int or type(target) is not int
                or initial != actual_initial[agent] or quota < 0 or target != initial + quota):
            raise RuntimeError(f'Invalid initial_count/target_count for agent {agent}; '
                               'do not reset or hand-edit the run snapshot.')


def load_collection_state(rows: list[list[str]]) -> dict:
    """Called exclusively under excel_lock; initialize/migrate without losing progress.

    New runs require an explicit `record.py start-run` invocation. A normal status,
    collector startup, or process resume NEVER creates a new run if a state exists.
    Historical v1 state migrates in-place, keeping the original baseline and all
    already-added records in the original run (even if progress.json is stale).
    """
    path = project.project_path('collection_state')
    if not path.exists():
        first = _snapshot_run(rows, 1)
        state = {'version': 2, 'active_run_id': first['run_id'], 'runs': [first]}
        _save_collection_state(state)
    else:
        state = json.loads(path.read_text(encoding='utf-8'))
        if state.get('version') == 1:
            baseline = state.get('baseline_row_count')
            if (type(baseline) is not int or baseline < 0 or baseline > len(rows) - 1
                    or state.get('baseline_sha256') != _historical_digest(rows[:baseline + 1])):
                raise RuntimeError('Existing workbook rows changed or baseline is invalid. '
                                   'Cannot safely migrate old collection_state.json.')
            first = _snapshot_run(rows, 1, state.get('initialized_at'), baseline,
                                  state.get('baseline_sha256'))
            state = {'version': 2, 'active_run_id': first['run_id'], 'runs': [first]}
            _save_collection_state(state)
    if (state.get('version') != 2 or not isinstance(state.get('runs'), list)
            or not state['runs'] or not isinstance(state.get('active_run_id'), str)):
        raise RuntimeError('Invalid collection_state.json; restore its matching backup.')
    if len({r.get('run_id') for r in state['runs']}) != len(state['runs']):
        raise RuntimeError('Duplicate run IDs in collection_state.json')
    for run in state['runs']:
        _validate_run(run, rows)
    _active_run(state)
    return state


def collection_counts(rows: list[list[str]], require_config_match: bool = True) -> tuple[Counter, dict]:
    """Return counts added in active run, plus immutable per-agent run snapshot."""
    state = load_collection_state(rows)
    run = _active_run(state)
    if require_config_match:
        for agent, snapshot in run['agents'].items():
            if snapshot['agent_quota'] != project.quota(agent):
                raise RuntimeError(
                    f'Config quota for agent {agent} changed during {run["run_id"]}: '
                    f'run={snapshot["agent_quota"]}, config={project.quota(agent)}. '
                    'Restore the run quota in config.json to resume, or start a NEW '
                    'run with `python record.py start-run` (use --force if incomplete).'
                )
    start = run['baseline_row_count'] + 1
    return _count_agent_rows(rows[start:]), run


def _run_snapshot(run: dict, agent: str, new_count: int) -> dict:
    snap = run['agents'][agent]
    return {'run_id': run['run_id'], 'initial_count': snap['initial_count'],
            'agent_quota': snap['agent_quota'], 'target_count': snap['target_count'],
            'current_count': snap['initial_count'] + new_count,
            'agent_new_images_collected': new_count,
            'remaining': max(0, snap['agent_quota'] - new_count)}


def _progress_description(run: dict, agent: str, new_count: int) -> str:
    snap = _run_snapshot(run, agent, new_count)
    return (f'run={snap["run_id"]} | progress={snap["current_count"]}/'
            f'{snap["target_count"]} | agent_new_images_collected={new_count}/'
            f'{snap["agent_quota"]}')


def cmd_start_run(args):
    """Explicitly start a fresh run, snapshotting all existing agent image counts."""
    with excel_lock():
        _, _, rows = load_sheet()
        state_file = project.project_path('collection_state')
        # On a pristine project, the first explicit start creates run-000001,
        # not an empty artificial run immediately followed by run-000002.
        if not state_file.exists():
            state = load_collection_state(rows)
            new_run = _active_run(state)
            previous = '(none)'
        else:
            state = load_collection_state(rows)
            current = _active_run(state)
            counts = _count_agent_rows(rows[current['baseline_row_count'] + 1:])
            unfinished = [a for a in project.ids()
                          if counts[a] < current['agents'][a]['agent_quota']]
            if unfinished and not args.force:
                fatal(f'Current run {current["run_id"]} is still incomplete for agents '
                      f'{", ".join(unfinished)}. To intentionally abandon the remaining '
                      'quota and start a fresh run, use `record.py start-run --force`.')
            next_number = max(int(r['run_id'].split('-')[-1]) for r in state['runs']) + 1
            new_run = _snapshot_run(rows, next_number)
            state['runs'].append(new_run)
            state['active_run_id'] = new_run['run_id']
            _save_collection_state(state)
            previous = current['run_id']
        for agent in project.ids():
            write_state(agent, 0, new_run, last_action='NEW_RUN_INITIALIZED')
        print(f'RUN_STARTED | run={new_run["run_id"]} | existing_workbook_rows={len(rows)-1} '
              f'| previous_run={previous}')
        for agent in project.ids():
            print(f'agent_{agent}: {_progress_description(new_run, agent, 0)}')

def _progress(agent: str) -> dict:
    dest = project.worker_file(agent, 'progress')
    if dest.exists():
        data = json.loads(dest.read_text(encoding='utf-8'))
    else:
        data = {}
    # Remove obsolete counters so state files cannot accidentally report legacy
    # workbook rows as new agent progress.
    for legacy in ('target', 'accepted', 'agent_target'):
        data.pop(legacy, None)
    data.setdefault('current_keyword', '')
    data.setdefault('page', '')
    data.setdefault('next_action', 'Start search using assigned instructions')
    data.setdefault('last_action', 'INIT')
    return data


def write_state(agent: str, new_count: int, run: dict, updates: dict | None = None,
                last_action: str | None = None):
    """Call inside excel_lock; cache is derived from the authoritative workbook/run."""
    p = _progress(agent)
    snap = _run_snapshot(run, agent, new_count)
    p.update({'agent_id': agent, **snap})
    if updates:
        p.update({k: v for k, v in updates.items() if v is not None})
    if last_action is not None:
        p['last_action'] = last_action
    p['last_updated'] = timestamp()
    _atomic_text(project.worker_file(agent, 'progress'), json.dumps(p, ensure_ascii=False, indent=2) + '\n')
    md = (f'# Agent {agent} — State\n\n'
          f'- **Run ID:** {snap["run_id"]}\n'
          f'- **Status:** {"COMPLETE" if new_count >= snap["agent_quota"] else "ACTIVE"}\n'
          f'- **Initial count (existing rows):** {snap["initial_count"]}\n'
          f'- **Target count:** {snap["target_count"]}\n'
          f'- **Total progress:** {snap["current_count"]} / {snap["target_count"]}\n'
          f'- **New unique images this run:** {new_count} / {snap["agent_quota"]}\n'
          f'- **Remaining this run:** {snap["remaining"]}\n'
          f'- **Current keyword:** {p.get("current_keyword", "")}\n'
          f'- **Search page:** {p.get("page", "")}\n'
          f'- **Last action:** {p.get("last_action", "")}\n'
          f'- **Next action:** {p.get("next_action", "")}\n'
          f'- **Last updated:** {p["last_updated"]}\n\n'
          '**Recovery:** Reopen the same run with `python record.py status --agent-id '
          f'{agent}`. The paired workbook and collection_state.json are authoritative. '
          'Only `python record.py start-run` creates a new fresh run.\n')
    _atomic_text(project.worker_file(agent, 'state'), md)

def event(agent: str, kind: str, details: str):
    with project.worker_file(agent, 'events').open('a', encoding='utf-8') as f:
        f.write(f'[{timestamp()}] {kind} | {details}\n')


def _existing_url_key(url: str) -> str:
    """Tolerate old workbook URL values without dropping historical dedupe."""
    try:
        return canonical_url(url)
    except (ValueError, AttributeError):
        return unicodedata.normalize('NFC', unquote(url.strip())).casefold()


def cmd_add(args):
    agent = agent_from_workdir(args.agent_id)
    try:
        name = safe_text(args.original_filename, 240, 'original_filename')
        url = safe_text(args.original_image_url, 4000, 'original_image_url')
        reason = safe_text(args.hrvs_suitability, 500, 'hrvs_suitability')
        canon_url = canonical_url(url)
        canon_name = canonical_filename(name)
    except ValueError as ex:
        fatal(str(ex))
    with excel_lock():
        xml, sheet_data, rows = load_sheet()
        counts, run = collection_counts(rows)
        if args.run_id and args.run_id != run['run_id']:
            fatal(f'Collector started in {args.run_id}, but active run is {run["run_id"]}. '
                  'Stop the old collector and initialize a fresh one for the new run.')
        new_count = counts[agent]
        limit = run['agents'][agent]['agent_quota']
        if new_count >= limit:
            print(f'QUOTA_REACHED | agent={agent} | {_progress_description(run, agent, new_count)}')
            write_state(agent, new_count, run, last_action='QUOTA_REACHED')
            event(agent, 'QUOTA_REACHED', f'{_progress_description(run, agent, new_count)}')
            return
        for r in rows[1:]:
            if canon_name == canonical_filename(r[1]) or canon_url == _existing_url_key(r[2]):
                print(f'DUPLICATE | agent={agent} | existing_agent={r[0]} '
                      f'| filename={r[1]} | {_progress_description(run, agent, new_count)}')
                write_state(agent, new_count, run, last_action=f'DUPLICATE {name}')
                event(agent, 'DUPLICATE', f'filename={name} | existing_agent={r[0]}')
                return
        try:
            _save_appended_row(xml, sheet_data, len(rows), [agent, name, url, reason])
        except PermissionError:
            fatal('Cannot write images.xlsx: close it in Excel and disable cloud live-sync')
        # Row-based reconciliation counts this append even if a crash precedes progress update.
        new_count += 1
        write_state(agent, new_count, run, last_action=f'ACCEPTED {name}')
        event(agent, 'ACCEPTED', f'filename={name} | {_progress_description(run, agent, new_count)}')
        print(f'ACCEPTED | agent={agent} | {_progress_description(run, agent, new_count)} | filename={name}')


def cmd_status(args):
    if args.all:
        with excel_lock():
            counts, rows, run = counts_and_rows()
            for a in project.ids():
                write_state(a, counts[a], run)
            lines = [f'agent_{a}: {_progress_description(run, a, counts[a])}' for a in project.ids()]
            lines.append(f'WORKBOOK | run={run["run_id"]} | existing_rows={run["baseline_row_count"]} '
                         f'| newly_collected_rows={len(rows)-1-run["baseline_row_count"]} '
                         f'| workbook_rows={len(rows)-1} | no_global_cap=true')
        print('\n'.join(lines))
    else:
        agent = agent_from_workdir(args.agent_id)
        with excel_lock():
            counts, rows, run = counts_and_rows()
            new_count = counts[agent]
            write_state(agent, new_count, run)
            state = _progress(agent)
        print(f'STATUS | agent={agent} | {_progress_description(run, agent, new_count)} '
              f'| remaining={max(0, run["agents"][agent]["agent_quota"]-new_count)} '
              f'| existing_workbook_rows={run["baseline_row_count"]}')
        print(f'KEYWORD | {state.get("current_keyword", "")} | page={state.get("page", "")}')
        print(f'NEXT | {state.get("next_action", "")}')


def cmd_checkpoint(args):
    agent = agent_from_workdir(args.agent_id)
    with excel_lock():
        counts, run = collection_counts(load_sheet()[2])
        if args.run_id and args.run_id != run['run_id']:
            fatal(f'Stale collector run {args.run_id}; active run is {run["run_id"]}')
        new_count = counts[agent]
        updates = {'current_keyword': args.keyword, 'page': args.page, 'next_action': args.next_action}
        write_state(agent, new_count, run, updates=updates, last_action='CHECKPOINT')
        event(agent, 'CHECKPOINT', f'{_progress_description(run, agent, new_count)} '
              f'| keyword={args.keyword or "(unchanged)"} | page={args.page or "(unchanged)"} '
              f'| next={args.next_action or "(unchanged)"}')
    print(f'CHECKPOINT | agent={agent} | {_progress_description(run, agent, new_count)}')


def cmd_note(args):
    agent = agent_from_workdir(args.agent_id)
    message = safe_text(args.message, 1000, 'message')
    with excel_lock():
        counts, run = collection_counts(load_sheet()[2])
        if args.run_id and args.run_id != run['run_id']:
            fatal(f'Stale collector run {args.run_id}; active run is {run["run_id"]}')
        new_count = counts[agent]
        write_state(agent, new_count, run, last_action=f'NOTE {message}')
        event(agent, 'NOTE', message)
    print(f'NOTED | agent={agent} | {_progress_description(run, agent, new_count)}')


def cmd_audit(_args):
    with excel_lock():
        counts, rows, run = counts_and_rows()
    urls = [_existing_url_key(r[2]) for r in rows[1:]]
    names = [canonical_filename(r[1]) for r in rows[1:]]
    url_dup = [k for k, v in Counter(urls).items() if k and v > 1]
    name_dup = [k for k, v in Counter(names).items() if k and v > 1]
    unknown_new = {k: v for k, v in counts.items() if k not in project.ids()}
    over = {k: v for k, v in counts.items()
            if k in project.ids() and v > run['agents'][k]['agent_quota']}
    print(f'RUN_ID={run["run_id"]} | WORKBOOK_ROWS={len(rows)-1} '
          f'| HISTORICAL_ROWS={run["baseline_row_count"]} '
          f'| NEW_ROWS={len(rows)-1-run["baseline_row_count"]}')
    print(f'UNIQUE_URLS={len(set(urls))} | UNIQUE_FILENAMES={len(set(names))}')
    print(f'DUPLICATE_URL_KEYS={len(url_dup)} | DUPLICATE_FILENAME_KEYS={len(name_dup)}')
    print(f'UNKNOWN_NEW_AGENT_IDS={unknown_new} | OVER_AGENT_QUOTA={over}')
    if url_dup or name_dup or unknown_new or over:
        raise SystemExit(1)


def counts_and_rows() -> tuple[Counter, list[list[str]], dict]:
    _, _, rows = load_sheet()
    counts, run = collection_counts(rows)
    return counts, rows, run

def main():
    parser = argparse.ArgumentParser(description='Atomic shared XLSX writer for HRVS Codex agents')
    sub = parser.add_subparsers(dest='action', required=True)
    start = sub.add_parser('start-run', help='Explicitly snapshot existing records and start a fresh collection run')
    start.add_argument('--force', action='store_true', help='Abandon any incomplete quota from the preceding run')
    start.set_defaults(func=cmd_start_run)
    add = sub.add_parser('add', help='Append a unique Commons original-file URL candidate')
    add.add_argument('--agent-id', help='Configured agent ID (normally inferred from cwd)')
    add.add_argument('--run-id', help='Optional pinned run ID: reject stale collector processes')
    add.add_argument('--original-filename', required=True)
    add.add_argument('--original-image-url', required=True)
    add.add_argument('--hrvs-suitability', required=True)
    add.set_defaults(func=cmd_add)
    status = sub.add_parser('status', help='Read real counts from shared Excel')
    status.add_argument('--agent-id', help='Configured agent ID (normally inferred from cwd)')
    status.add_argument('--all', action='store_true', help='Print all workers')
    status.set_defaults(func=cmd_status)
    cp = sub.add_parser('checkpoint', help='Remember current keyword / page / next step')
    cp.add_argument('--agent-id')
    cp.add_argument('--run-id')
    cp.add_argument('--keyword')
    cp.add_argument('--page')
    cp.add_argument('--next-action')
    cp.set_defaults(func=cmd_checkpoint)
    nt = sub.add_parser('note', help='Write an important blocker / reminder')
    nt.add_argument('--agent-id')
    nt.add_argument('--run-id')
    nt.add_argument('--message', required=True)
    nt.set_defaults(func=cmd_note)
    audit = sub.add_parser('audit', help='Check uniqueness and per-agent quotas')
    audit.set_defaults(func=cmd_audit)
    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
