"""Tests for numeric input, payments and monthly sheet names."""
import unittest
import main

class CoreTests(unittest.TestCase):
    def test_decimal_inputs(self):
        for value in ('75,50', '75.50', 75.5):
            self.assertEqual(main.to_float(value), 75.5)
        self.assertEqual(main.to_float(None), 0)
        self.assertEqual(main.to_float(' '), 0)

    def test_bad_numbers(self):
        for value in ('not a number', 'NaN', 'inf', '=SUM(A1:A2)'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                main.to_float(value)

    def test_payment_allocation(self):
        for m,n,paid,expected in [(20,30,10,(10,30)),(20,30,25,(0,25)),(20,30,50,(0,0)),(20,30,60,(-10,0))]:
            with self.subTest(paid=paid):
                self.assertEqual(main.apply_payment(m,n,paid), expected)

    def test_month(self):
        self.assertEqual(main.parse_month_sheet('December 2025').month, 12)
        with self.assertRaises(ValueError):
            main.parse_month_sheet('invalid')
