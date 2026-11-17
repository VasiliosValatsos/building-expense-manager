import unittest
from openpyxl import Workbook
import main

class HeatingTests(unittest.TestCase):
    def test_formula_keeps_zero_distinct_from_blank(self):
        sheet = Workbook().active
        sheet['H36'] = 0
        sheet['H37'] = None
        main.fix_heating_formulas(sheet)
        for row in range(36,51):
            self.assertIn(f'IF(H{row}="",1,H{row})', sheet[f'I{row}'].value.replace(' ',''))
        self.assertEqual(sheet['H36'].value, 0)
        self.assertIsNone(sheet['H37'].value)

    def test_no_participating_property(self):
        sheet = Workbook().active
        sheet['C14'] = 100
        sheet['H36'] = 0
        with self.assertRaises(ValueError):
            main.validate_heating_expense(sheet, 100)
        main.validate_heating_expense(sheet, 0)
        sheet['H36'] = None
        main.validate_heating_expense(sheet,100)
