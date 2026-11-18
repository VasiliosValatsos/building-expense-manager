import unittest
from decimal import Decimal

from openpyxl import Workbook
import main

class HeatingTests(unittest.TestCase):
    def test_formula_keeps_zero_distinct_from_blank(self):
        sheet = Workbook().active
        sheet['H36'] = 0
        sheet['H37'] = None
        main.fix_heating_formulas(sheet)
        for row in range(36,51):
            self.assertIn(f'IF(H{row}="",1,', sheet[f'C{row}'].value.replace(' ',''))
        self.assertEqual(sheet['H36'].value, 0)
        self.assertIsNone(sheet['H37'].value)

    def test_no_participating_property(self):
        sheet = Workbook().active
        sheet['C14'] = 100
        sheet['E14'] = "Fictional owner"
        sheet['H36'] = 0
        with self.assertRaises(ValueError):
            main.monthly_allocations(sheet, {"G4": 100})
        self.assertEqual(sum(main.monthly_allocations(sheet, {"G4": 0})["G4"]), 0)
        sheet['H36'] = None
        self.assertEqual(main.monthly_allocations(sheet, {"G4": 100})["G4"][0], 100)
        sheet['C15'] = 100
        sheet['E15'] = "Fictional owner 2"
        sheet['H37'] = Decimal("0.25")
        self.assertEqual(main.monthly_allocations(sheet, {"G4": 100})["G4"][:2], [80, 20])
