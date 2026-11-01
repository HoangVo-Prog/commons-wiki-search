#!/usr/bin/env python3
"""Wikimedia Commons metadata-only search for Codex HRVS agents (stdlib).

No browser, no image bytes. The `collect` command finds *metadata candidates*,
not visually verified HRVS images. The official imageinfo.url is the original URL.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlsplit, unquote
from urllib.request import Request, urlopen

import record
import project

ROOT = Path(__file__).resolve().parent
ENDPOINT = 'https://commons.wikimedia.org/w/api.php'
GATE = ROOT / '.commons_api.lock'
API_CLOCK = ROOT / '.commons_api_last_request.txt'
CONFIG = project.CONFIG_PATH
DEFAULT_PAGE = 50
MIN_MP = 8_000_000
VALID_TYPES = {'image/jpeg', 'image/png', 'image/webp', 'image/tiff'}
GOOD = re.compile(r'crowd|shopping|market|bazaar|street|pedestrian|crossing|festival|parade|sidewalk|vendors?|mall|commercial|bustl|city\s*centre|city\s*center|downtown|public\s*square|promenade|souq|souk|laneway|plaza|night\s*market|road\s*traffic', re.I)
BAD = re.compile(r'portrait|selfie|photoshoot|cosplay|model\s*shoot|aerial|drone|satellite|painting|watercolou?r|illustration|cartoon|render|map\s*of|panoram|collage|empty\s*street|sunset|skyline|skyscraper|close[ -]?up|studio|logo|poster|advertisement|sculpture|drawing', re.I)


def stamp():
    return datetime.now().astimezone().isoformat(timespec='seconds')


def read_config():
    if not CONFIG.exists():
        raise RuntimeError('config.json not found. Restore it from project package.')
    config = json.loads(CONFIG.read_text(encoding='utf-8'))
    contact = os.environ.get('HRVS_WIKIMEDIA_CONTACT', config.get('wikimedia_contact', '')).strip()
    if not contact or 'YOUR_' in contact.upper() or 'EXAMPLE.' in contact.upper() or contact.lower() == 'unset':
        raise RuntimeError('Set config.json wikimedia_contact to YOUR real email, website or Wikimedia user page (User-Agent policy). No requests sent.')
    interval = float(config.get('min_seconds_between_api_requests', 1.5))
    if interval < 1:
        raise RuntimeError('API global throttle must be at least 1 second')
    return contact, interval


@contextmanager
def gate_lock(timeout=180):
    with open(GATE, 'a+b') as f:
        f.seek(0, os.SEEK_END)
        if f.tell() == 0:
            f.write(b'0')
            f.flush()
        start = time.monotonic()
        while True:
            try:
                f.seek(0)
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except (BlockingIOError, OSError):
                if time.monotonic() - start > timeout:
                    raise RuntimeError('Timed out waiting for shared API lock')
                time.sleep(0.15)
        try:
            yield
        finally:
            if os.name == 'nt':
                f.seek(0)
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)


def api_request(params):
    """One globally serialized API request (including retries/backoff)."""
    contact, interval = read_config()
    headers = {
        'User-Agent': f'HRVSCommonsMetadataBot/1.0 ({contact}) Python-urllib/3',
        'Accept': 'application/json',
    }
    url = ENDPOINT + '?' + urlencode({**params, 'action': 'query', 'format': 'json', 'formatversion': 2, 'maxlag': 5})
    for attempt in range(5):
        with gate_lock():
            now = time.time()
            if API_CLOCK.exists():
                try:
                    last = float(API_CLOCK.read_text(encoding='utf-8').strip())
                    time.sleep(max(0.0, last + interval - now))
                except ValueError:
                    pass
            API_CLOCK.write_text(str(time.time()), encoding='utf-8')
            try:
                with urlopen(Request(url, headers=headers), timeout=35) as response:
                    data = json.loads(response.read(5_000_000).decode('utf-8'))
            except HTTPError as exc:
                if exc.code not in (429, 500, 502, 503, 504):
                    raise RuntimeError(f'HTTP {exc.code} from Wikimedia API; verify User-Agent and access') from exc
                retry = exc.headers.get('Retry-After')
                try:
                    delay = min(120.0, max(float(retry), 2.0)) if retry else min(60.0, 2**(attempt + 1))
                except ValueError:
                    delay = min(60.0, 2**(attempt + 1))
                API_CLOCK.write_text(str(time.time() + delay), encoding='utf-8')
                print(f'API_RETRY | http={exc.code} | delay={delay}s', file=sys.stderr)
                continue
            except (URLError, TimeoutError) as exc:
                print(f'API_RETRY | network={exc} | attempt={attempt + 1}', file=sys.stderr)
                API_CLOCK.write_text(str(time.time() + min(30, 2**(attempt + 1))), encoding='utf-8')
                continue
            if data.get('error'):
                code = data['error'].get('code', 'unknown')
                if code in ('maxlag', 'ratelimited'):
                    delay = min(60.0, max(5.0, float(data['error'].get('lag', 5)) + 1.0))
                    API_CLOCK.write_text(str(time.time() + delay), encoding='utf-8')
                    print(f'API_RETRY | code={code} | delay={delay}s', file=sys.stderr)
                    continue
                raise RuntimeError(f'Wikimedia API error: {data["error"]}')
            return data
    raise RuntimeError('Wikimedia API failed after bounded retries; pause and try later')


def get_candidates(keyword, offset=0, limit=DEFAULT_PAGE):
    payload = api_request({
        'generator': 'search', 'gsrsearch': keyword, 'gsrnamespace': 6,
        'gsrlimit': limit, 'gsroffset': offset,
        'prop': 'imageinfo', 'iiprop': 'url|size|mime|mediatype',
        'iilimit': 1,
    })
    pages = payload.get('query', {}).get('pages', [])
    if isinstance(pages, dict):
        pages = list(pages.values())
    items = []
    for page in pages:
        infos = page.get('imageinfo') or []
        if not infos or not isinstance(infos, list):
            continue
        info = infos[0]
        title = page.get('title', '')
        if not title.startswith('File:'):
            continue
        items.append({
            'original_filename': title[5:],
            'original_image_url': info.get('url', ''),
            'width': int(info.get('width', 0)),
            'height': int(info.get('height', 0)),
            'mime': info.get('mime', ''),
            'mediatype': info.get('mediatype', ''),
            'keyword': keyword,
        })
    has_more = bool(payload.get('continue', {}).get('gsroffset') is not None or payload.get('continue', {}).get('gsrcontinue') is not None)
    return items, has_more


def filter_candidate(item, strict_title=True):
    name = item['original_filename']
    url = item['original_image_url']
    if item['width'] * item['height'] < MIN_MP or min(item['width'], item['height']) < 1500:
        return False
    if item['mime'] not in VALID_TYPES or item['mediatype'].upper() != 'BITMAP':
        return False
    if BAD.search(name):
        return False
    if strict_title and not GOOD.search(name.replace('_', ' ').replace('-', ' ')):
        return False
    try:
        record.canonical_url(url)
    except ValueError:
        return False
    try:
        path = unquote(urlsplit(url).path).rsplit('/', 1)[-1]
        if record.canonical_filename(path) != record.canonical_filename(name):
            return False
    except Exception:
        return False
    return True


def keyword_pool(agent):
    file = project.worker_file(agent, 'keywords')
    return [l.strip() for l in file.read_text(encoding='utf-8').splitlines()
            if l.strip() and not l.strip().startswith('#')]


def cursor_path(agent):
    return project.worker_file(agent, 'cursor')


def load_cursor(agent):
    path = cursor_path(agent)
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'keyword_index': 0, 'offset': 0}


def save_cursor(agent, data):
    path = cursor_path(agent)
    tmp = path.with_suffix('.json.tmp')
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    os.replace(tmp, path)


def run_record(*args):
    p = subprocess.run([sys.executable, str(ROOT / 'record.py'), *args], cwd=ROOT,
                       text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if p.returncode != 0:
        raise RuntimeError(f'record.py error: {p.stderr.strip() or p.stdout.strip()}')
    return p.stdout.strip()


def get_agent(args):
    return record.agent_from_workdir(args.agent_id)


def get_agent_new_images_collected(agent):
    with record.excel_lock():
        counts, _, _ = record.counts_and_rows()
        return counts[agent]


def cmd_test(_args):
    contact, interval = read_config()
    print(f'CONFIG_OK | contact={contact} | min_interval={interval}s | endpoint={ENDPOINT}')
    data = api_request({'meta': 'siteinfo', 'siprop': 'general'})
    print(f'API_OK | wiki={data.get("query", {}).get("general", {}).get("sitename", "?")}')


def cmd_search(args):
    items, more = get_candidates(args.keyword, args.offset, args.limit)
    qualified = [x for x in items if filter_candidate(x, strict_title=not args.relaxed_title)]
    print(f'SEARCH | keyword={args.keyword} | offset={args.offset} | returned={len(items)} | metadata_qualified={len(qualified)} | more={more}')
    for x in qualified:
        print(json.dumps({'original_filename': x['original_filename'], 'original_image_url': x['original_image_url'],
                          'dimensions': f'{x["width"]}x{x["height"]}', 'keyword': x['keyword'],
                          'visual_inspection': 'NOT_PERFORMED'}, ensure_ascii=False))


def cmd_collect(args):
    agent = get_agent(args)
    keywords = keyword_pool(agent)
    if not keywords:
        raise RuntimeError(f'Agent {agent} has no keywords.txt')
    accepted = get_agent_new_images_collected(agent)
    limit = project.quota(agent)
    if accepted >= limit:
        print(f'QUOTA_REACHED | agent={agent} | agent_new_images_collected={accepted}/{limit}')
        return
    cursor = load_cursor(agent)
    num_pages = 0
    new = 0
    seen = 0
    while accepted < limit and num_pages < args.max_pages and new < args.max_accept:
        ki = cursor['keyword_index']
        offset = cursor['offset']
        if ki >= len(keywords):
            run_record('note', '--agent-id', agent, '--message', 'All configured API keyword searches exhausted; append fresh focused keywords to your own keywords file to continue')
            print(f'KEYWORDS_EXHAUSTED | agent={agent} | agent_new_images_collected={accepted}/{limit}')
            return
        keyword = keywords[ki]
        items, more = get_candidates(keyword, offset, args.page_size)
        num_pages += 1
        print(f'PAGE | agent={agent} | keyword={keyword!r} | offset={offset} | results={len(items)}')
        for item in items:
            seen += 1
            if not filter_candidate(item, strict_title=not args.relaxed_title):
                continue
            filename = item['original_filename']
            # EXPLICITLY metadata-only. Do not assert that bags, people, etc. were seen.
            reason = f'Metadata-only candidate; keyword: {keyword}; relevant filename; {item["width"]}x{item["height"]} px; visual complexity unverified'
            result = run_record('add', '--agent-id', agent, '--original-filename', filename,
                                '--original-image-url', item['original_image_url'], '--hrvs-suitability', reason)
            if result.startswith('ACCEPTED'):
                new += 1
                accepted += 1
                print(f'{result}')
            elif result.startswith('QUOTA_REACHED'):
                accepted = limit
                break
            elif not result.startswith('DUPLICATE'):
                raise RuntimeError(f'Unexpected record.py output: {result}')
            # Do not interrupt midway through page: preserve cursor only after page
            # so restart safely resubmits uncommitted candidates (writer dedupes).
        if more and items:
            cursor['offset'] = offset + args.page_size
        else:
            cursor = {'keyword_index': ki + 1, 'offset': 0}
        cursor['last_updated'] = stamp()
        save_cursor(agent, cursor)
        next_kw = keywords[cursor['keyword_index']] if cursor['keyword_index'] < len(keywords) else '(no keywords left)'
        run_record('checkpoint', '--agent-id', agent, '--keyword', keyword,
                   '--page', str(offset // args.page_size + 1),
                   '--next-action', f'API resume: keyword={next_kw}; offset={cursor["offset"]}; cursor=api_cursor.json')
    print(f'BATCH_DONE | agent={agent} | agent_new_images_collected={accepted}/{limit} | new={new} | pages={num_pages} | examined={seen} | resume with same collect command')


def main():
    parser = argparse.ArgumentParser(description='API-only Wikimedia Commons metadata collector. DOES NOT DOWNLOAD IMAGES.')
    sub = parser.add_subparsers(dest='command', required=True)
    t = sub.add_parser('test', help='Verify contact and Wikimedia Action API')
    t.set_defaults(fn=cmd_test)
    s = sub.add_parser('search', help='Read-only search, no Excel write')
    s.add_argument('--keyword', required=True)
    s.add_argument('--offset', type=int, default=0)
    s.add_argument('--limit', type=int, default=50, choices=range(1, 51), metavar='1..50')
    s.add_argument('--relaxed-title', action='store_true', help='Include relevant search matches even when filename lacks scene keywords; lower precision')
    s.set_defaults(fn=cmd_search)
    c = sub.add_parser('collect', help='Batch metadata candidates into shared Excel via record.py')
    c.add_argument('--agent-id', help='Worker identity from config.json')
    c.add_argument('--max-pages', type=int, default=8)
    c.add_argument('--max-accept', type=int, default=25)
    c.add_argument('--page-size', type=int, default=50, choices=range(1, 51), metavar='1..50')
    c.add_argument('--relaxed-title', action='store_true', help='Lower precision; use only if explicitly acceptable')
    c.set_defaults(fn=cmd_collect)
    args = parser.parse_args()
    try:
        args.fn(args)
    except (RuntimeError, ValueError, OSError) as exc:
        print(f'ERROR | {exc}', file=sys.stderr)
        raise SystemExit(2)


if __name__ == '__main__':
    main()
