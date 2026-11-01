"""Offline integration tests: historical preservation, resume, quotas, lock safety."""
from __future__ import annotations

import concurrent.futures
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from zipfile import ZipFile

SOURCE = Path(__file__).resolve().parents[1]


class IndependentQuotaTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / 'project'
        self.root.mkdir()
        for script in ('project.py', 'record.py', 'commons_api.py'):
            shutil.copy2(SOURCE / script, self.root / script)
        shutil.copy2(SOURCE / 'images.xlsx', self.root / 'images.xlsx')
        cfg = json.loads((SOURCE / 'config.json').read_text(encoding='utf-8'))
        cfg['agent_count'] = 2
        cfg['agent_quota'] = {'01': 2, '02': 3}
        (self.root / 'config.json').write_text(json.dumps(cfg), encoding='utf-8')
        for a in ('01', '02'):
            dest = self.root / 'workers' / f'agent_{a}'
            dest.mkdir(parents=True)
            (dest / 'agent_id.txt').write_text(a)
            (dest / 'events.log').write_text('')
            (dest / 'keywords.txt').write_text('crowded street\n')
            (dest / 'api_cursor.json').write_text('{"keyword_index":0,"offset":0}')

    def cli(self, *args, check=True):
        p = subprocess.run([sys.executable, str(self.root / args[0]), *args[1:]],
                           cwd=self.root, capture_output=True, text=True)
        if check and p.returncode:
            self.fail(f'Command failed: {args}\nstdout={p.stdout}\nstderr={p.stderr}')
        return p

    def uri(self, key):
        return f'https://upload.wikimedia.org/wikipedia/commons/a/ab/Street_{key}.jpg'

    def add(self, agent, key):
        return self.cli('record.py', 'add', '--agent-id', agent,
                        '--original-filename', f'Street_{key}.jpg',
                        '--original-image-url', self.uri(key),
                        '--hrvs-suitability', 'Metadata-only candidate test')

    def insert_legacy(self, agent, key):
        # Deliberately seed EXISTING images before any baseline exists.
        code = ('import record\n'
                'xml, data, rows = record.load_sheet()\n'
                f'record._save_appended_row(xml, data, len(rows), [{agent!r}, {"Street_"+key+".jpg"!r}, {self.uri(key)!r}, "Preexisting image"])')
        result = subprocess.run([sys.executable, '-c', code], cwd=self.root, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    def read_rows(self):
        code = 'import record,json; print(json.dumps(record.load_sheet()[2]))'
        result = subprocess.run([sys.executable, '-c', code], cwd=self.root, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def progress(self, agent):
        return json.loads((self.root / 'workers' / f'agent_{agent}' / 'progress.json').read_text())

    def test_historical_urls_never_consume_quota_and_are_deduplicated(self):
        self.insert_legacy('01', 'old1')
        self.insert_legacy('LEGACY', 'old2')
        original_rows = self.read_rows()
        self.assertFalse((self.root / 'collection_state.json').exists())
        status = self.cli('record.py', 'status', '--all').stdout
        self.assertIn('existing_rows=2', status)
        self.assertIn('agent_01: agent_new_images_collected=0/2', status)
        self.assertEqual(json.loads((self.root / 'collection_state.json').read_text())['baseline_row_count'], 2)
        self.assertIn('DUPLICATE', self.add('01', 'old2').stdout)
        self.assertEqual(self.progress('01')['agent_new_images_collected'], 0)
        self.assertIn('ACCEPTED', self.add('01', 'new1').stdout)
        self.assertIn('ACCEPTED', self.add('01', 'new2').stdout)
        self.assertIn('QUOTA_REACHED', self.add('01', 'new3').stdout)
        self.assertIn('ACCEPTED', self.add('02', 'agent2').stdout)
        self.assertEqual(self.read_rows()[:len(original_rows)], original_rows)
        self.assertEqual(self.progress('01')['agent_new_images_collected'], 2)
        self.assertEqual(self.progress('02')['agent_new_images_collected'], 1)
        self.assertEqual(len(self.read_rows()) - 1, 5)
        self.assertEqual(self.cli('record.py', 'audit').returncode, 0)

    def test_crash_after_workbook_append_and_stale_cache_reconciles(self):
        self.insert_legacy('01', 'old')
        self.cli('record.py', 'status', '--all')
        self.insert_legacy('01', 'crash_new')  # simulates successful XLSX append then crash before write_state
        self.assertIn('agent_new_images_collected=1/2', self.cli('record.py', 'status', '--agent-id', '01').stdout)
        p = self.progress('01')
        p['agent_new_images_collected'] = 0  # stale/corrupted cache
        (self.root / 'workers' / 'agent_01' / 'progress.json').write_text(json.dumps(p))
        self.assertIn('agent_new_images_collected=1/2', self.cli('record.py', 'status', '--agent-id', '01').stdout)
        self.assertIn('DUPLICATE', self.add('01', 'crash_new').stdout)
        self.assertIn('ACCEPTED', self.add('01', 'another').stdout)
        self.assertIn('QUOTA_REACHED', self.add('01', 'extra').stdout)
        self.assertEqual(self.progress('01')['agent_new_images_collected'], 2)

    def test_config_change_is_independent_and_no_global_cap(self):
        self.cli('record.py', 'status', '--all')
        cfg_path = self.root / 'config.json'
        cfg = json.loads(cfg_path.read_text())
        cfg['agent_quota'] = {'01': 1, '02': 4}
        cfg_path.write_text(json.dumps(cfg))
        self.assertIn('ACCEPTED', self.add('01', 'one').stdout)
        self.assertIn('QUOTA_REACHED', self.add('01', 'two').stdout)
        for n in range(4):
            self.assertIn('ACCEPTED', self.add('02', f'b{n}').stdout)
        self.assertIn('QUOTA_REACHED', self.add('02', 'fifth').stdout)
        self.assertIn('newly_collected_rows=5', self.cli('record.py', 'status', '--all').stdout)
        self.assertEqual(self.cli('record.py', 'audit').returncode, 0)

    def test_simultaneous_writes_enforce_independent_quotas_and_global_dedupe(self):
        self.insert_legacy('01', 'old')
        self.cli('record.py', 'status', '--all')
        requests = [('01', f'a{i}') for i in range(5)] + [('02', f'b{i}') for i in range(6)]
        requests += [('01', 'shared'), ('02', 'shared'), ('02', 'old')]
        with concurrent.futures.ThreadPoolExecutor(max_workers=9) as pool:
            responses = list(pool.map(lambda x: self.add(*x), requests))
        alloutput = '\n'.join(r.stdout for r in responses)
        self.assertIn('QUOTA_REACHED', alloutput)
        self.assertEqual(self.progress('01')['agent_new_images_collected'], 2)
        self.assertEqual(self.progress('02')['agent_new_images_collected'], 3)
        self.assertEqual(len(self.read_rows()) - 1, 6)
        self.assertEqual(self.cli('record.py', 'audit').returncode, 0)
        with ZipFile(self.root / 'images.xlsx') as x:
            self.assertIsNone(x.testzip())

    def test_historical_modification_detected(self):
        import os
        import xml.etree.ElementTree as ET
        self.insert_legacy('01', 'old')
        self.cli('record.py', 'status', '--all')
        workbook = self.root / 'images.xlsx'
        sheet_name = 'xl/worksheets/sheet1.xml'
        namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        changed = self.uri('modified')
        tmp = self.root / 'modified.xlsx'
        with ZipFile(workbook) as zin, ZipFile(tmp, 'w') as zout:
            for info in zin.infolist():
                content = zin.read(info.filename)
                if info.filename == sheet_name:
                    sheet = ET.fromstring(content)
                    cell = next(c for c in sheet.iter(f'{{{namespace}}}c') if c.get('r') == 'C2')
                    node = next(cell.iter(f'{{{namespace}}}t'))
                    node.text = changed
                    content = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
                zout.writestr(info, content)
        os.replace(tmp, workbook)
        result = self.cli('record.py', 'status', '--agent-id', '01', check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('baseline is invalid', result.stderr)


if __name__ == '__main__':
    unittest.main()
