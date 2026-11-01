#!/usr/bin/env python3
"""Shared Excel append-only writer with independent config-defined agent quotas.

XLSX is an OOXML ZIP archive. This script edits only the first worksheet's four
columns while retaining the rest of the workbook. OS locking and atomic replace
make simultaneous worker processes safe on a local disk. Historical workbook
rows are baseline-only and never consume newly collected per-agent quotas.
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


def load_collection_state(rows: list[list[str]]) -> dict:
    """Called ONLY inside excel_lock. Persist the existing-row baseline first.

    After initialization, all new accepted rows follow the baseline and counts
    are reconstructed directly from workbook rows. This survives a crash after
    workbook replacement but before progress.json is updated. The file MUST NOT
    be removed/reset on resume. Historical rows (including matching agent IDs)
    never consume new-collection quota.
    """
    path = project.project_path('collection_state')
    if not path.exists():
        state = {
            'version': 1,
            'baseline_row_count': len(rows) - 1,
            'baseline_sha256': _historical_digest(rows),
            'initialized_at': timestamp(),
        }
        _atomic_text(path, json.dumps(state, ensure_ascii=False, indent=2) + '\n')
    else:
        state = json.loads(path.read_text(encoding='utf-8'))
    baseline = state.get('baseline_row_count')
    if (state.get('version') != 1 or type(baseline) is not int or baseline < 0
            or baseline > len(rows) - 1
            or state.get('baseline_sha256') != _historical_digest(rows[:baseline + 1])):
        raise RuntimeError('Existing workbook rows changed or baseline is invalid. '
                           'Restore the original workbook/baseline together; never reset collection_state.json to bypass progress.')
    return state


def collection_counts(rows: list[list[str]]) -> tuple[Counter, dict]:
    """Reconcile durable per-agent counts from post-baseline workbook rows."""
    state = load_collection_state(rows)
    start = state['baseline_row_count'] + 1
    return Counter(r[0] for r in rows[start:]), state


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


def write_state(agent: str, new_count: int, updates: dict | None = None, last_action: str | None = None):
    """Call while holding excel_lock to serialize state/cache updates too."""
    p = _progress(agent)
    limit = project.quota(agent)
    p.update({'agent_id': agent, 'agent_quota': limit,
              'agent_new_images_collected': new_count})
    if updates:
        p.update({k: v for k, v in updates.items() if v is not None})
    if last_action is not None:
        p['last_action'] = last_action
    p['last_updated'] = timestamp()
    _atomic_text(project.worker_file(agent, 'progress'), json.dumps(p, ensure_ascii=False, indent=2) + '\n')
    md = (f'# Agent {agent} — State\n\n'
          f'- **Status:** {"COMPLETE" if new_count >= limit else "ACTIVE"}\n'
          f'- **Agent new images collected:** {new_count} / {limit}\n'
          f'- **Remaining:** {max(0, limit - new_count)}\n'
          f'- **Current keyword:** {p.get("current_keyword", "")}\n'
          f'- **Search page:** {p.get("page", "")}\n'
          f'- **Last action:** {p.get("last_action", "")}\n'
          f'- **Next action:** {p.get("next_action", "")}\n'
          f'- **Last updated:** {p["last_updated"]}\n\n'
          '**Recovery:** Run `python project.py manifest --agent-id '
          f'{agent}` and `python record.py status --agent-id {agent}`. '
          'The workbook plus immutable collection_state.json baseline is authoritative. '
          'Never delete or reset the baseline on resume.\n')
    _atomic_text(project.worker_file(agent, 'state'), md)


def event(agent: str, kind: str, details: str):
    with project.worker_file(agent, 'events').open('a', encoding='utf-8') as f:
        f.write(f'[{timestamp()}] {kind} | {details}\n')


def counts_and_rows() -> tuple[Counter, list[list[str]], dict]:
    _, _, rows = load_sheet()
    counts, baseline = collection_counts(rows)
    return counts, rows, baseline


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
    limit = project.quota(agent)
    with excel_lock():
        xml, sheet_data, rows = load_sheet()
        counts, baseline = collection_counts(rows)  # initializes historical baseline before appending
        new_count = counts[agent]
        if new_count >= limit:
            print(f'QUOTA_REACHED | agent={agent} | agent_new_images_collected={new_count}/{limit}')
            write_state(agent, new_count, last_action='QUOTA_REACHED')
            event(agent, 'QUOTA_REACHED', f'agent_new_images_collected={new_count}')
            return
        for r in rows[1:]:
            if canon_name == canonical_filename(r[1]) or canon_url == _existing_url_key(r[2]):
                print(f'DUPLICATE | agent={agent} | existing_agent={r[0]} | filename={r[1]} | agent_new_images_collected={new_count}/{limit}')
                write_state(agent, new_count, last_action=f'DUPLICATE {name}')
                event(agent, 'DUPLICATE', f'filename={name} | existing_agent={r[0]}')
                return
        try:
            _save_appended_row(xml, sheet_data, len(rows), [agent, name, url, reason])
        except PermissionError:
            fatal('Cannot write images.xlsx: close it in Excel and disable cloud live-sync')
        # If interrupted here, the appended XLSX row is still counted on resume.
        new_count += 1
        write_state(agent, new_count, last_action=f'ACCEPTED {name}')
        event(agent, 'ACCEPTED', f'filename={name} | agent_new_images_collected={new_count}/{limit}')
        print(f'ACCEPTED | agent={agent} | agent_new_images_collected={new_count}/{limit} | filename={name}')


def cmd_status(args):
    if args.all:
        with excel_lock():
            counts, rows, baseline = counts_and_rows()
            for a in project.ids():
                write_state(a, counts[a])
            lines = [f'agent_{a}: agent_new_images_collected={counts[a]}/{project.quota(a)}' for a in project.ids()]
            lines.append(f'WORKBOOK | existing_rows={baseline["baseline_row_count"]} | newly_collected_rows={len(rows)-1-baseline["baseline_row_count"]} | workbook_rows={len(rows)-1} | no_global_cap=true')
        print('\n'.join(lines))
    else:
        agent = agent_from_workdir(args.agent_id)
        with excel_lock():
            counts, rows, baseline = counts_and_rows()
            new_count = counts[agent]
            limit = project.quota(agent)
            write_state(agent, new_count)
            state = _progress(agent)
        print(f'STATUS | agent={agent} | agent_new_images_collected={new_count}/{limit} | remaining={max(0, limit-new_count)} | existing_workbook_rows={baseline["baseline_row_count"]}')
        print(f'KEYWORD | {state.get("current_keyword", "")} | page={state.get("page", "")}')
        print(f'NEXT | {state.get("next_action", "")}')


def cmd_checkpoint(args):
    agent = agent_from_workdir(args.agent_id)
    with excel_lock():
        counts, _, _ = counts_and_rows()
        new_count = counts[agent]
        limit = project.quota(agent)
        updates = {'current_keyword': args.keyword, 'page': args.page, 'next_action': args.next_action}
        write_state(agent, new_count, updates=updates, last_action='CHECKPOINT')
        event(agent, 'CHECKPOINT', f'agent_new_images_collected={new_count}/{limit} | keyword={args.keyword or "(unchanged)"} | page={args.page or "(unchanged)"} | next={args.next_action or "(unchanged)"}')
    print(f'CHECKPOINT | agent={agent} | agent_new_images_collected={new_count}/{limit}')


def cmd_note(args):
    agent = agent_from_workdir(args.agent_id)
    message = safe_text(args.message, 1000, 'message')
    with excel_lock():
        counts, _, _ = counts_and_rows()
        new_count = counts[agent]
        limit = project.quota(agent)
        write_state(agent, new_count, last_action=f'NOTE {message}')
        event(agent, 'NOTE', message)
    print(f'NOTED | agent={agent} | agent_new_images_collected={new_count}/{limit}')


def cmd_audit(_args):
    with excel_lock():
        counts, rows, baseline = counts_and_rows()
    urls = [_existing_url_key(r[2]) for r in rows[1:]]
    names = [canonical_filename(r[1]) for r in rows[1:]]
    url_dup = [k for k, v in Counter(urls).items() if k and v > 1]
    name_dup = [k for k, v in Counter(names).items() if k and v > 1]
    unknown_new = {k: v for k, v in counts.items() if k not in project.ids()}
    # If an operator decreases a quota after collecting, existing records are
    # preserved and collecting stops, but audit reports the over-quota situation.
    over = {k: v for k, v in counts.items() if k in project.ids() and v > project.quota(k)}
    print(f'WORKBOOK_ROWS={len(rows)-1} | HISTORICAL_ROWS={baseline["baseline_row_count"]} | NEW_ROWS={len(rows)-1-baseline["baseline_row_count"]}')
    print(f'UNIQUE_URLS={len(set(urls))} | UNIQUE_FILENAMES={len(set(names))}')
    print(f'DUPLICATE_URL_KEYS={len(url_dup)} | DUPLICATE_FILENAME_KEYS={len(name_dup)}')
    print(f'UNKNOWN_NEW_AGENT_IDS={unknown_new} | OVER_AGENT_QUOTA={over}')
    if url_dup or name_dup or unknown_new or over:
        raise SystemExit(1)

def main():
    parser = argparse.ArgumentParser(description='Atomic shared XLSX writer for HRVS Codex agents')
    sub = parser.add_subparsers(dest='action', required=True)
    add = sub.add_parser('add', help='Append a unique verified Commons original-file URL')
    add.add_argument('--agent-id', help='Configured agent ID (normally inferred from cwd)')
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
    cp.add_argument('--keyword')
    cp.add_argument('--page')
    cp.add_argument('--next-action')
    cp.set_defaults(func=cmd_checkpoint)
    nt = sub.add_parser('note', help='Write an important blocker / reminder')
    nt.add_argument('--agent-id')
    nt.add_argument('--message', required=True)
    nt.set_defaults(func=cmd_note)
    audit = sub.add_parser('audit', help='Check uniqueness and per-agent quotas')
    audit.set_defaults(func=cmd_audit)
    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
