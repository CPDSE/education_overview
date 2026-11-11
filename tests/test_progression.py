"""Check the editable progression source and its proposed/observed boundary."""
import json
from pathlib import Path
import re
import sys
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts import build

ROOT = build.ROOT


class ProgressionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / 'progression.xlsx'
        self.workbook = openpyxl.load_workbook(build.PROGRESSION_PATH)
        self.addCleanup(self.workbook.close)

    def save_and_load(self):
        self.workbook.save(self.path)
        return build.load_progression(self.path)

    def test_complete_proposal_has_different_paces_and_no_claimed_alignment(self):
        model = self.save_and_load()
        self.assertEqual({(t['semester'], t['category']) for t in model['targets']},
                         {(s, c) for s in range(1, 7) for c in build.DS_CATEGORIES})
        self.assertEqual(len(model['targets']), 48)
        self.assertEqual(len(model['review']), 48)
        self.assertEqual({r['match'] for r in model['review']}, {'Not reviewed'})
        self.assertEqual({r['participation'] for r in model['review']}, {'To confirm'})
        ml = [t for t in model['targets'] if t['category'] == 'Machine learning and AI']
        programming = [t for t in model['targets'] if t['category'] == 'Computing and programming']
        self.assertEqual(ml[0]['emphasis'], 'No new target')
        self.assertEqual(ml[3]['emphasis'], 'Introduce')
        self.assertEqual(programming[0]['emphasis'], 'Introduce')
        for sheet in ['Targets', 'Review', 'Pilot links']:
            self.assertTrue(self.workbook[sheet].tables)
            self.assertTrue(self.workbook[sheet].data_validations.dataValidation)
            self.assertTrue(self.workbook[sheet].freeze_panes)
        for row in self.workbook['Semester map'].iter_rows(min_row=5, max_row=10, min_col=2, max_col=9):
            for cell in row:
                self.assertTrue(cell.value.startswith('=HYPERLINK('))
                self.assertIn('MATCH(', cell.value)
        # Preview values are provided even before Excel first recalculates.
        cached = openpyxl.load_workbook(build.PROGRESSION_PATH, data_only=True)
        self.addCleanup(cached.close)
        for row in cached['Semester map'].iter_rows(min_row=5, max_row=10, min_col=2, max_col=9):
            self.assertTrue(all(isinstance(cell.value, str) and '\n\n' in cell.value for cell in row))

    def test_edited_workbook_and_teacher_review_reach_html(self):
        title = 'Check a new dataset </script><script>unsafe()</script>'
        self.workbook['Targets']['E5'] = title
        ws = self.workbook['Review']
        for cell, value in {
            'B5': 'Example course', 'C5': 'Partly aligned', 'D5': 'Required',
            'E5': 'Students explain variation but not uncertainty.',
            'F5': 'An example and a short reference.', 'G5': 'Teacher-reviewed exercise, section 2',
            'H5': 'Example reviewer', 'I5': '2026-11-06', 'J5': 'Add one interpretation prompt.',
        }.items():
            ws[cell] = value
        self.workbook.save(self.path)
        output = Path(self.directory.name) / 'clean.html'
        with patch.object(build, 'PROGRESSION_PATH', self.path):
            build.build_clean(ROOT / 'templates/clean.html', output)
        text = output.read_text()
        script = re.search(r'<script id="course-data" type="application/json">(.*?)</script>', text, re.S)[1]
        self.assertNotIn('<', script)
        model = json.loads(script)['progression']
        self.assertEqual(model['targets'][0]['title'], title)
        self.assertEqual(model['review'][0]['match'], 'Partly aligned')
        self.assertEqual(model['review'][0]['next_change'], 'Add one interpretation prompt.')

    def test_unreviewed_cannot_be_promoted_without_evidence(self):
        self.workbook['Review']['C5'] = 'Gap confirmed'
        with self.assertRaisesRegex(ValueError, 'observed task, evidence, teacher and date'):
            self.save_and_load()

    def test_invalid_references_and_source_formulas_fail_the_build(self):
        cases = [
            ('Targets', 'A6', 'MS-01', 'duplicate or empty'),
            ('Targets', 'I5', 'UNKNOWN', 'unknown or future'),
            ('Targets', 'I5', 'MS-06', 'unknown or future'),
            ('Targets', 'I5', 'MS-01', 'circular'),
            ('Targets', 'C5', 'Unrecognised category', 'unknown category'),
            ('Targets', 'B5', '1.5', 'integer'),
            ('Targets', 'E5', '=1+1', 'plain values'),
            ('Review', 'A5', 'UNKNOWN', 'unknown or duplicate'),
            ('Review', 'C5', 'Done', 'unknown review'),
            ('Pilot links', 'C5', 'UNKNOWN', 'unknown handover'),
            ('Pilot links', 'G5', 'Implemented', 'owner, date and outcome'),
        ]
        for sheet, cell, value, message in cases:
            with self.subTest(sheet=sheet, cell=cell, value=value):
                original = self.workbook[sheet][cell].value
                self.workbook[sheet][cell] = value
                with self.assertRaisesRegex(ValueError, message):
                    self.save_and_load()
                self.workbook[sheet][cell] = original

    def test_incomplete_semester_map_is_rejected(self):
        self.workbook['Targets'].delete_rows(52)
        with self.assertRaisesRegex(ValueError, 'all eight categories'):
            self.save_and_load()

    def test_pilot_evidence_ids_are_checked_against_actual_records(self):
        self.workbook['Pilot links']['D5'] = 'MISSING-EXERCISE'
        self.workbook.save(self.path)
        with patch.object(build, 'PROGRESSION_PATH', self.path):
            with self.assertRaisesRegex(ValueError, 'unknown exercise evidence'):
                build.build_clean(ROOT / 'templates/clean.html', Path(self.directory.name) / 'clean.html')


if __name__ == '__main__':
    unittest.main()
