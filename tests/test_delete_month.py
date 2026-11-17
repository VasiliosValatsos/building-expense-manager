import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
import main

class DeleteTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.p = Path(temp.name)/'ΚΟΙΝΟΧΡΗΣΤΑ.xlsx'
        self.q = Path(temp.name)/'ΒΙΒΛΙΟ.xlsx'
        self.a,self.b = Workbook(),Workbook()
        self.addCleanup(self.a.close)
        self.addCleanup(self.b.close)
        self.a.active.title = self.b.active.title = 'December 2025'
        for row in range(14,29):
            for col in 'STU': self.a.active[f'{col}{row}']='=1'
        for addr in ('F31','L31','L34'): self.b.active[addr]='=1'
        main.create_month_sheets(self.a,self.b,'December 2025')
        self.a.save(self.p); self.b.save(self.q)
        self.pdf = main.month_pdf_path(self.p,main.parse_month_sheet('January 2026'))
        self.pdf.write_bytes(b'test placeholder')

    def test_delete(self):
        main.delete_month_files(self.a,self.p,self.b,self.q,'January 2026')
        self.assertFalse(self.pdf.exists())
        for path in (self.p,self.q):
            wb=load_workbook(path)
            try: self.assertEqual(wb.sheetnames,['December 2025'])
            finally: wb.close()

    def test_restore_pdf_on_failed_save(self):
        before=(self.p.read_bytes(),self.q.read_bytes())
        with patch.object(main,'save_workbooks_safely',side_effect=OSError('simulated failure')):
            with self.assertRaises(OSError):
                main.delete_month_files(self.a,self.p,self.b,self.q,'January 2026')
        self.assertTrue(self.pdf.exists())
        self.assertEqual(before,(self.p.read_bytes(),self.q.read_bytes()))

    def test_protect_last_sheet(self):
        del self.a['January 2026']; del self.b['January 2026']
        with self.assertRaises(ValueError): main.deletion_target(self.a,self.b)

    def test_reject_mismatched_month(self):
        self.b['January 2026'].title='February 2026'
        with self.assertRaises(ValueError): main.deletion_target(self.a,self.b)
