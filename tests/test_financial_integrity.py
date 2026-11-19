import unittest
from contextlib import contextmanager
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import main


TEMPLATES = Path(__file__).resolve().parents[1] / "templates"


class CalculatedSheet:
    def __init__(self, name, values):
        self.name = name
        self.values = values

    def range(self, address):
        value = self.values.get(address)

        class Cell:
            def options(self, *, err_to_str):
                if not err_to_str:
                    raise AssertionError("Excel error values must remain visible")
                return self

        cell = Cell()
        cell.value = value
        return cell


def calculated_values(name, values):
    @contextmanager
    def provider(path, sheet_name):
        if sheet_name != name:
            raise AssertionError(sheet_name)
        yield CalculatedSheet(name, values)

    return patch.object(main, "calculated_sheet", provider)


class PaymentIntegrityTests(unittest.TestCase):
    def test_payment_and_credit_examples(self):
        cases = [
            (60, 40, 30, (30, 40)),
            (60, 40, 80, (0, 20)),
            (60, 40, 100, (0, 0)),
            (60, 40, 120, (-20, 0)),
            (-20, 100, 10, (0, 70)),
            (100, -20, 10, (70, 0)),
            (-20, 20, 0, (0, 0)),
            (-20, -10, 0, (-20, -10)),
            (0, -10, 0, (0, -10)),
        ]
        for tenant, owner, cash, expected in cases:
            with self.subTest(tenant=tenant, owner=owner, cash=cash):
                result = main.apply_payment(tenant, owner, cash)
                self.assertEqual(result, expected)
                self.assertEqual(sum(result), tenant + owner - cash)

    def test_full_settlement(self):
        for tenant, owner, expected in [
            (60, 40, (0, 0, 100)),
            (100, -20, (0, 0, 80)),
            (-20, 20, (0, 0, 0)),
            (-20, 10, (-10, 0, 0)),
            (0, -10, (0, -10, 0)),
        ]:
            with self.subTest(tenant=tenant, owner=owner):
                self.assertEqual(main.settle_balance(tenant, owner), expected)

    def test_negative_payment_rejected(self):
        with self.assertRaises(ValueError):
            main.apply_payment(10, 10, -1)

    def test_amounts_and_exact_allocation(self):
        self.assertEqual(main.money_cents("75,50"), 7550)
        for value in ("1.001", "1.0000000000000000000000000001", "1e-10000", "NaN", True, "#REF!"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                main.money_cents(value)
        self.assertEqual(main.allocate_cents(10000, [1, 1, 1]), [3334, 3333, 3333])
        self.assertEqual(
            main.allocate_cents(27500, [1, Decimal("0.5"), Decimal("0.25"), 0, 1]),
            [10000, 5000, 2500, 0, 10000],
        )
        self.assertEqual(main.allocate_cents(1, [0, 1, 1]), [0, 1, 0])
        with self.assertRaises(ValueError):
            main.allocate_cents(1, [0, 0])

    def test_ledger_records_only_actual_income(self):
        common, ledger = Workbook().active, Workbook().active
        main.record_payment(common, ledger, 14, -2000, 2000, 0)
        main.record_payment(common, ledger, 15, 0, -1000, 0, full=True)
        main.record_payment(common, ledger, 16, 10000, -2000, 0, full=True)
        self.assertEqual([ledger[f"F{row}"].value for row in (9, 10, 11)], [None, None, 80])
        self.assertEqual([common[f"N{row}"].value for row in (14, 15, 16)], [0, -10, 0])

    def test_languages_share_financial_behavior(self):
        try:
            for language in ("en", "el"):
                main.set_language(language)
                self.assertEqual(main.money_cents("10,25"), 1025)
                self.assertEqual(main.settle_balance(-20, 40), (0, 0, 20))
                self.assertEqual(main.allocate_cents(100, [1, 1, 1]), [34, 33, 33])
        finally:
            main.set_language("en")


class WorkbookIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.common_book = load_workbook(next(TEMPLATES.glob("Κ*.xlsx")))
        self.ledger_book = load_workbook(next(TEMPLATES.glob("Β*.xlsx")))
        self.addCleanup(self.common_book.close)
        self.addCleanup(self.ledger_book.close)

    def test_three_month_cash_carry_forward(self):
        name = self.ledger_book.active.title
        self.ledger_book.active["F8"] = 1000
        for opening, income, expense, closing in [(1000, 200, 100, 1100), (1100, 70, 90, 1080), (1080, 30, 60, 1050)]:
            ledger = self.ledger_book[name]
            self.assertEqual(ledger["F8"].value, opening)
            ledger["F9"] = income
            ledger["L8"] = expense
            cells = {addr: ledger[addr].value for addr in ("F8", "F9", "L8")}
            cells.update(F31=opening + income, L31=expense, L34=closing)
            with calculated_values(name, cells):
                transferred = main.read_closing_cash("ledger.xlsx", name)
            self.assertEqual(transferred, closing * 100)
            name, _, _, _, next_ledger = main.create_month_sheets(
                self.common_book, self.ledger_book, name, opening_cash=transferred,
            )
            self.assertEqual(next_ledger["F8"].value, closing)
            self.assertIsNone(next_ledger["F9"].value)
            self.assertIsNone(next_ledger["L8"].value)
            self.assertEqual(ledger["F8"].value, opening)
            self.assertEqual(ledger["F9"].value, income)

    def test_management_fee_is_one_owner_charge_and_one_cash_outflow(self):
        common, ledger = self.common_book.active, self.ledger_book.active
        common["C14"], common["E14"] = 35, "Fictional A"
        common["C15"], common["E15"] = 35, "Fictional B"
        common["D8"] = 60
        data = {
            "standard": dict(zip(main.STANDARD_BILLS, [10000, 5500, 2233, 4000, 500])),
            "urgent": ("=Literal fictional repair", 7500), "urgent2": ("", 0),
            "extra1": ("Fictional supplies", 2550), "extra2": ("", 0),
        }
        data["allocations"] = main.monthly_allocations(common, main.bill_amounts(data))
        self.assertEqual(data["allocations"]["D8"][:2], [3000, 3000])
        for _ in range(2):
            main.record_expenses(common, ledger, data)
        ledger["F8"], ledger["F9"] = 400, 100
        main.validate_month_cash(ledger)
        self.assertEqual(sum(sum(amounts) for amounts in data["allocations"].values()), 38283)
        self.assertEqual(sum(main.money_cents(ledger[f"L{row}"].value, blank_zero=True) for row in range(8, 21)), 38283)
        self.assertEqual(ledger["L16"].value, 60)

        self.assertEqual(ledger["K17"].data_type, "s")
        self.assertEqual(common["A6"].data_type, "s")
        cells = {f"F{row}": ledger[f"F{row}"].value for row in range(8, 24)}
        cells.update({f"L{row}": ledger[f"L{row}"].value for row in range(8, 21)})
        cells.update(F31=500, L31=382.83, L34=117.17)
        with calculated_values(ledger.title, cells):
            self.assertEqual(main.read_closing_cash("ledger.xlsx", ledger.title), 11717)
        base = ledger.title
        for _ in range(2):
            name, _, _, new_month, new_ledger = main.create_month_sheets(
                self.common_book, self.ledger_book, base, opening_cash=11717,
            )
            self.assertIsNone(new_ledger["L16"].value)
            main.record_management_fee(new_month, new_ledger)
            self.assertEqual(new_ledger["L16"].value, 60)
            del self.common_book[name]
            del self.ledger_book[name]
        self.assertEqual(ledger["L16"].value, 60)

    def test_previous_manual_expenses_are_not_carried_into_the_new_month(self):
        common = self.common_book.active
        common["D4"], common["D8"], common["G11"] = 42, 60, 14
        _, _, _, new_common, _ = main.create_month_sheets(
            self.common_book, self.ledger_book, common.title, opening_cash=0,
        )
        self.assertEqual([new_common[cell].value for cell in ("D4", "D8", "G11")], [None, 60, None])
        self.assertEqual([common[cell].value for cell in ("D4", "D8", "G11")], [42, 60, 14])

    def test_active_validation_and_property_balance_transition(self):
        area, owner, factor = main.apartment_details("20,12345", "  Fictional   Owner  ", "0,25")
        self.assertEqual((area, owner, factor), (Decimal("20.12345"), "Fictional Owner", Decimal("0.25")))
        for values in [(12, " ", None), (None, "Owner", None), (0, "Owner", None),
                       (-2, "Owner", None), ("NaN", "Owner", None), (12, "Owner", 1.1),
                       ("1e-10000", "Owner", None)]:
            with self.subTest(values=values), self.assertRaises(ValueError):
                main.apartment_details(*values)
        with self.assertRaises(ValueError):
            main.monthly_allocations(self.common_book.active, {"G5": 500})
        for balances in [(100, 0), (0, -100), (100, -100)]:
            with self.assertRaises(ValueError):
                main.validate_balance_transition("Current owner", "", *balances)
        main.validate_balance_transition("Current owner", "Replacement owner", 0, 0)

    def test_pending_settings_do_not_mutate_prior_financial_fields(self):
        common, ledger = self.common_book.active, self.ledger_book.active
        common["C14"], common["E14"], common["D8"] = 20, "Old owner", 60
        properties = [("", Decimal(30), "", "=Literal owner", None)] + [("", None, "", "", None)] * 14
        main.write_properties(common, ledger, properties, 7000, pending=True)
        self.assertEqual((common["E14"].value, common["D8"].value), ("Old owner", 60))
        name, _, _, new_common, new_ledger = main.create_month_sheets(
            self.common_book, self.ledger_book, common.title, opening_cash=0,
        )
        with self.assertRaises(ValueError):
            main.apply_pending_properties(new_common, new_ledger, [(0, -1, -1)] + [(0, 0, 0)] * 14)
        main.apply_pending_properties(new_common, new_ledger, [(0, 0, 0)] * 15)
        self.assertEqual((new_common["E14"].value, new_common["E14"].data_type), ("=Literal owner", "s"))
        self.assertEqual(new_ledger["E9"].data_type, "s")
        self.assertEqual(new_common["D8"].value, 70)
        self.assertIsNone(new_common["DF1"].value)
        self.assertEqual(common["E14"].value, "Old owner")

    def test_excel_errors_and_blank_required_values_have_cell_context(self):
        name = "December 2025"
        for value in ("#DIV/0!", "#VALUE!", "#REF!", "#N/A", None, "", "0", False):
            with self.subTest(value=value), calculated_values(name, {"S14": value}):
                with self.assertRaisesRegex(ValueError, r"common.xlsx / December 2025!S14"):
                    main.read_debts_from_excel("common.xlsx", name, [14])
        with calculated_values(name, {"S14": 0, "T14": 0, "U14": 0}):
            self.assertEqual(main.read_debts_from_excel("common.xlsx", name, [14]), [(0, 0, 0)])
        with calculated_values(name, {"F8": 0, "F31": 0, "L31": 0, "L34": "#REF!"}):
            with self.assertRaisesRegex(ValueError, r"ledger.xlsx / December 2025!L34"):
                main.read_closing_cash("ledger.xlsx", name)
        with calculated_values(name, dict(F8=1000, F9=200, L8=100, F31=1200, L31=100, L34=1000)):
            with self.assertRaisesRegex(ValueError, r"!L34"):
                main.read_closing_cash("ledger.xlsx", name)

    def test_legacy_totals_needing_rounding_confirmation(self):
        with patch.object(main.messagebox, "askyesno", return_value=False) as confirm:
            self.assertTrue(main.confirm_legacy_rounding([(101, -99, 2)]))
            confirm.assert_not_called()
            self.assertFalse(main.confirm_legacy_rounding([(101, -100, 2)]))
            self.assertEqual(confirm.call_args.kwargs["default"], "no")

    def test_setup_and_public_template_protection(self):
        self.assertTrue(main.setup_required(self.common_book.active, self.ledger_book.active))
        self.assertEqual(main.setup_values("1000,00", " One   Street ", 2026, 1), (100000, "One Street", "December 2025"))
        for settings in [("-1", "One Street", 2026, 1), ("0", "  ", 2026, 1),
                         ("NaN", "One Street", 2026, 1), ("0", "One Street", 2026, 13)]:
            with self.assertRaises(ValueError):
                main.setup_values(*settings)
        for path in TEMPLATES.glob("*.xlsx"):
            with self.assertRaises(ValueError):
                main.protect_template(path)

    def test_compatible_month_titles_do_not_allow_different_buildings(self):
        with TemporaryDirectory() as directory:
            common_path = Path(directory) / "common.xlsx"
            ledger_path = Path(directory) / "ledger.xlsx"
            self.common_book.active["E1"] = "Fictional address A"
            self.ledger_book.active["K3"] = "Fictional address B"
            self.common_book.save(common_path)
            self.ledger_book.save(ledger_path)
            with self.assertRaisesRegex(ValueError, "different building addresses"):
                main.open_workbook_pair(common_path, ledger_path)

    def test_corrupt_config_values_and_failed_atomic_replace(self):
        with TemporaryDirectory() as directory:
            config = Path(directory) / "config_paths.json"
            with patch.object(main, "CONFIG_FILE", str(config)):
                self.assertEqual(main.load_paths(), {})
                config.write_bytes(b"{broken")
                self.assertEqual(main.load_paths(), {})
                main.save_paths({"language": "el", "file1": [], "file2": "relative.xlsx"})
                self.assertEqual(main.load_paths(), {"language": "el"})
                before = config.read_bytes()
                with patch.object(main.os, "replace", side_effect=PermissionError):
                    with self.assertRaises(PermissionError):
                        main.save_paths({"language": "en"})
                self.assertEqual(config.read_bytes(), before)
