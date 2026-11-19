import unittest
from pathlib import Path
from openpyxl import load_workbook
import main

ROOT = Path(__file__).resolve().parents[1]
COMMON = ROOT / 'templates' / 'ΚΟΙΝΟΧΡΗΣΤΑ.xlsx'
LEDGER = ROOT / 'templates' / 'ΒΙΒΛΙΟ ΕΣΟΔΩΝ_ΕΞΟΔΩΝ.xlsx'

class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(COMMON.is_file(), f'Add blank template: {COMMON.name}')
        self.assertTrue(LEDGER.is_file(), f'Add blank template: {LEDGER.name}')
        self.a = load_workbook(COMMON)
        self.addCleanup(self.a.close)
        self.b = load_workbook(LEDGER)
        self.addCleanup(self.b.close)

    def test_template_schema(self):
        self.assertEqual(len(self.a.sheetnames),1)
        self.assertEqual(len(self.b.sheetnames),1)
        name = main.validate_workbook_pair(self.a,self.b)
        self.assertEqual(name,self.a.sheetnames[-1])
        for row in range(14,29):
            for col in 'STU':
                self.assertEqual(self.a[name][f'{col}{row}'].data_type,'f')
            for col in 'CDE':
                self.assertIn(self.a[name][f'{col}{row}'].value,(None,'',' '))
        for addr in ('F31','L31','L34'):
            self.assertEqual(self.b[name][addr].data_type,'f')

    def test_month_copy_preserves_original(self):
        name = main.validate_workbook_pair(self.a,self.b)
        original_formula = self.a[name]['I36'].value
        original_merges = set(self.a[name].merged_cells.ranges)
        original_cf = len(self.a[name].conditional_formatting)
        new,_,_,common,ledger = main.create_month_sheets(self.a,self.b,name,opening_cash=0)
        self.assertNotEqual(name,new)
        self.assertEqual(self.a[name]['I36'].value,original_formula)
        self.assertEqual(set(common.merged_cells.ranges),original_merges)
        self.assertEqual(len(common.conditional_formatting),original_cf)
        self.assertEqual(common['U28'].data_type,'f')
        self.assertEqual(ledger['L34'].data_type,'f')
