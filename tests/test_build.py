"""Regression checks for both workbook-driven HTML builds.

Run with: python3 -m unittest discover -s tests
"""
import csv
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts import build

ROOT = build.ROOT
TEMPLATE = ROOT / 'templates/clean.html'


class BuildTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.output = Path(self.directory.name) / 'preview.html'

    def run_build(self, *args, cwd=ROOT):
        return subprocess.run([sys.executable, str(ROOT / 'scripts/build.py'), *map(str, args)],
                              cwd=cwd, text=True, capture_output=True)

    def test_alternative_preserves_workbook_and_pilot_records(self):
        result = self.run_build('--template', TEMPLATE, '--output', self.output)
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.output.read_text()
        payload = json.loads(re.search(r'<script id="course-data" type="application/json">(.*?)</script>', text, re.S)[1])
        for university in build.UNIVERSITIES:
            semesters, _, _ = build.load_courses(university['sheet'])
            expected = [c for courses in semesters.values() for c in courses]
            self.assertEqual(payload['universities'][university['key']]['courses'],
                             json.loads(json.dumps(expected)))
        with build.EVIDENCE_PATH.open(encoding='utf-8-sig', newline='') as source:
            self.assertEqual(payload['evidence'], list(csv.DictReader(source)))
        self.assertNotIn('<!-- CLEAN:DATA -->', text)
        self.assertIn('Curriculum progression', text)

    def test_default_command_still_builds_original_in_place(self):
        work = Path(self.directory.name)
        for filename in ['scripts/build.py', 'index.html', 'assets/data/Pharma_DS_Course_Overview.xlsx']:
            (work / filename).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ROOT / filename, work / filename)
        # Defaults locate the repository even when invoked from another directory.
        result = subprocess.run([sys.executable, str(work / 'scripts/build.py')],
                                cwd=work / 'scripts', text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('50 courses across 2 universities', result.stdout)
        text = (work / 'index.html').read_text()
        self.assertEqual(len(re.findall(r'<div class="box(?: |")', text)), 50)
        self.assertIn('<!-- GRID:sdu -->', text)
        self.assertIn('<!-- GRID:ucph -->', text)

    def test_template_cannot_be_overwritten(self):
        template = Path(self.directory.name) / 'template.html'
        shutil.copy2(TEMPLATE, template)
        original = template.read_bytes()
        for args in [('--template', template), ('--template', template, '--output', template)]:
            result = self.run_build(*args)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(template.read_bytes(), original)

    def test_embedded_spreadsheet_text_cannot_close_script(self):
        semesters, categories, mapped = build.load_courses('SDU')
        hostile = '</script><script>alert("workbook")</script>'
        semesters[1][0]['title_en'] = hostile
        with patch.object(build, 'load_courses', return_value=(semesters, categories, mapped)):
            build.build_clean(TEMPLATE, self.output)
        text = self.output.read_text()
        script = re.search(r'<script id="course-data" type="application/json">(.*?)</script>', text, re.S)[1]
        self.assertNotIn('<', script)
        self.assertEqual(json.loads(script)['universities']['sdu']['courses'][0]['title_en'], hostile)

    def test_nested_output_keeps_local_links_resolvable(self):
        result = self.run_build('--template', TEMPLATE, '--output', self.output)
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.output.read_text()
        self.assertIn('assets/data/Pharma_DS_Progression.xlsx', text)
        self.assertIn('assets/data/evidence.csv', text)
        for path in re.findall(r'(?:src|href)="([^"]+)"', text):
            if path.startswith(('../', './')):
                self.assertTrue((self.output.parent / path).exists(), path)

    def test_alternative_build_from_outside_repository(self):
        result = self.run_build('--template', TEMPLATE, '--output', self.output,
                                cwd=self.directory.name)
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = json.loads(re.search(r'<script id="course-data" type="application/json">(.*?)</script>',
                                      self.output.read_text(), re.S)[1])
        self.assertEqual(payload['progression'], build.load_progression())
        self.assertEqual(len(payload['evidence']), 11)


if __name__ == '__main__':
    unittest.main()
