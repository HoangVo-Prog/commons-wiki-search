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

    def active_run(self):
        data = json.loads((self.root / 'collection_state.json').read_text())
        return next(x for x in data['runs'] if x['run_id'] == data['active_run_id'])

    def test_historical_urls_never_consume_quota_and_are_deduplicated(self):
        self.insert_legacy('01', 'old1')
        self.insert_legacy('LEGACY', 'old2')
        original_rows = self.read_rows()
        self.assertFalse((self.root / 'collection_state.json').exists())
        status = self.cli('record.py', 'status', '--all').stdout
        self.assertIn('existing_rows=2', status)
        self.assertIn('progress=1/3 | agent_new_images_collected=0/2', status)
        run = self.active_run()
        self.assertEqual(run['baseline_row_count'], 2)
        self.assertEqual(run['agents']['01'], {'initial_count': 1, 'agent_quota': 2, 'target_count': 3})
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

    def test_config_changes_apply_only_to_new_runs(self):
        self.cli('record.py', 'status', '--all')
        cfg_path = self.root / 'config.json'
        cfg = json.loads(cfg_path.read_text())
        cfg['agent_quota'] = {'01': 1, '02': 4}
        cfg_path.write_text(json.dumps(cfg))
        # A config edit must not silently mutate the persisted target on resume.
        error = self.cli('record.py', 'status', '--agent-id', '01', check=False)
        self.assertNotEqual(error.returncode, 0)
        self.assertIn('changed during run', error.stderr)
        self.assertIn('still incomplete', self.cli('record.py', 'start-run', check=False).stderr)
        self.assertIn('run-000002', self.cli('record.py', 'start-run', '--force').stdout)
        self.assertEqual(self.active_run()['agents']['01']['target_count'], 1)
        self.assertEqual(self.active_run()['agents']['02']['target_count'], 4)
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

    def test_fresh_second_run_snapshots_all_previous_records(self):
        self.insert_legacy('01', 'old')
        self.insert_legacy('02', 'other_old')
        old_bytes = (self.root / 'images.xlsx').read_bytes()
        initial = self.cli('record.py', 'status', '--all').stdout
        self.assertIn('agent_01: run=run-000001 | progress=1/3', initial)
        self.assertEqual((self.root / 'images.xlsx').read_bytes(), old_bytes)
        self.add('01', 'first_run_1')
        self.add('01', 'first_run_2')
        for n in range(3):
            self.add('02', f'first_run_b{n}')
        self.assertIn('run-000002', self.cli('record.py', 'start-run').stdout)
        second = self.active_run()
        self.assertEqual(second['agents']['01'], {'initial_count': 3, 'agent_quota': 2, 'target_count': 5})
        self.assertEqual(second['agents']['02'], {'initial_count': 4, 'agent_quota': 3, 'target_count': 7})
        progress = self.cli('record.py', 'status', '--agent-id', '01').stdout
        self.assertIn('progress=3/5', progress)
        self.assertIn('agent_new_images_collected=0/2', progress)
        self.assertIn('DUPLICATE', self.add('01', 'first_run_1').stdout)
        self.assertIn('ACCEPTED', self.add('01', 'second_run_1').stdout)
        self.assertIn('progress=4/5', self.cli('record.py', 'status', '--agent-id', '01').stdout)
        self.assertIn('run-000002', self.progress('01')['run_id'])
        self.assertIn('ACCEPTED', self.add('01', 'second_run_2').stdout)
        self.assertIn('QUOTA_REACHED', self.add('01', 'second_run_3').stdout)
        self.assertEqual(len(json.loads((self.root / 'collection_state.json').read_text())['runs']), 2)
        self.assertEqual(self.cli('record.py', 'audit').returncode, 0)

    def test_legacy_v1_baseline_migrates_without_reclassifying_rows(self):
        self.insert_legacy('01', 'before_migration')
        self.cli('record.py', 'status', '--all')
        old = self.active_run()
        prior = {'version': 1, 'initialized_at': old['initialized_at'],
                 'baseline_row_count': old['baseline_row_count'],
                 'baseline_sha256': old['baseline_sha256']}
        (self.root / 'collection_state.json').write_text(json.dumps(prior))
        self.insert_legacy('01', 'after_old_baseline')
        output = self.cli('record.py', 'status', '--agent-id', '01').stdout
        self.assertIn('progress=2/3', output)
        self.assertIn('agent_new_images_collected=1/2', output)
        self.assertEqual(self.active_run()['agents']['01']['initial_count'], 1)
        self.assertEqual(self.progress('01')['agent_new_images_collected'], 1)
        self.assertIn('ACCEPTED', self.add('01', 'final_new').stdout)
        self.assertIn('QUOTA_REACHED', self.add('01', 'excess').stdout)

    def test_example_1000_existing_plus_500_fresh_quota(self):
        cfg_path = self.root / 'config.json'
        cfg = json.loads(cfg_path.read_text())
        cfg['agent_quota']['01'] = 500
        cfg_path.write_text(json.dumps(cfg))
        code = ('import record,json; '
                'rows=[list(record.COLS)]+[["01","old","old","old"] for _ in range(1000)]; '
                'print(json.dumps(record._snapshot_run(rows,1)["agents"]["01"]))')
        p = subprocess.run([sys.executable, '-c', code], cwd=self.root,
                           capture_output=True, text=True)
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout),
                         {'initial_count': 1000, 'agent_quota': 500, 'target_count': 1500})

    def test_stale_run_cannot_write_to_fresh_run(self):
        self.cli('record.py', 'start-run')
        self.assertEqual(self.active_run()['run_id'], 'run-000001')
        self.cli('record.py', 'start-run', '--force')
        p = self.cli('record.py', 'add', '--agent-id', '01', '--run-id', 'run-000001',
                     '--original-filename', 'Street_stale.jpg',
                     '--original-image-url', self.uri('stale'),
                     '--hrvs-suitability', 'Test candidate', check=False)
        self.assertNotEqual(p.returncode, 0)
        self.assertIn('Collector started in', p.stderr)
        self.assertEqual(len(self.read_rows()), 1)

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
