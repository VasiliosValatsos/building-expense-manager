import tempfile
import unittest
import os
import sys
import subprocess
import json
from contextlib import contextmanager
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

import main
from test_financial_integrity import CalculatedSheet


class StabilityTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def legacy_pair(self):
        with ZipFile(Path(__file__).with_name("baseline-82e289b.zip")) as fixture:
            fixture.extractall(self.folder)
        common = next(self.folder.glob("Κ*.xlsx"))
        ledger = next(self.folder.glob("Β*.xlsx"))
        for path, values in (
            (common, {"E1": "Fictional address", "C14": 10, "E14": "Owner A", "D14": "Tenant A", "D8": 60}),
            (ledger, {"K3": "Fictional address", "F8": 1000}),
        ):
            book = load_workbook(path)
            for address, value in values.items():
                book.active[address] = value
            book.save(path)
            book.close()
        return common, ledger

    def save_fixture(self):
        paths = [self.folder / f"{name}.xlsx" for name in ("common", "ledger")]
        books = [Workbook(), Workbook()]
        for book, path in zip(books, paths):
            self.addCleanup(book.close)
            book.save(path)
            book.active["A1"] = "Committed"
        return books, paths

    def test_saved_workbooks_survive_lock_cleanup_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            paths = [Path(directory) / f"{name}.xlsx" for name in ("common", "ledger")]
            books = [Workbook(), Workbook()]
            unlink = Path.unlink

            def fail_lock(path, *args, **kwargs):
                if path.suffix == ".lock":
                    raise PermissionError("lock is busy")
                return unlink(path, *args, **kwargs)

            for book, path in zip(books, paths):
                book.save(path)
                book.active["A1"] = "Committed"
            with patch.object(Path, "unlink", fail_lock):
                warnings = main.save_workbooks_safely(books[0], paths[0], books[1], paths[1])
            self.assertTrue(any(".lock" in warning for warning in warnings))
            for path in paths:
                reopened = load_workbook(path)
                self.assertEqual(reopened.active["A1"].value, "Committed")
                reopened.close()

    def test_expense_modified_after_allocation(self):
        values = {"G5": 120, "G12": 120, "H14": 50, "H15": 50,
                  "J14": 50, "J15": 50, "L14": 50, "L15": 50}
        common = CalculatedSheet("January 2026", values)
        ledger = CalculatedSheet(common.name, {"L9": 100})
        with self.assertRaisesRegex(ValueError, r"January 2026.*G5.*L9.*20\.00"):
            main.reconcile_expenses("common.xlsx", common, "ledger.xlsx", ledger, detailed=False)

    def test_reconciliation_rejects_old_shares_even_if_both_books_match(self):
        values = {"G5": 120, "G12": 120, "L12": 120, "H14": 50, "H15": 50,
                  "AZ14": 5000, "AZ15": 5000, "R11": "OK"}
        common = CalculatedSheet("January 2026", values)
        ledger = CalculatedSheet(common.name, {"L9": 120})
        with self.assertRaisesRegex(ValueError, r"January 2026.*G5.*AZ14:AZ28.*20\.00"):
            main.reconcile_expenses("common.xlsx", common, "ledger.xlsx", ledger, detailed=True)

    def test_near_equal_input_requires_supported_precision(self):
        with self.assertRaises(ValueError):
            main.apartment_details(Decimal("1.000000000001"), "Fictional", None)

    def test_legacy_first_month_and_two_more_issues_reopen_xlsx(self):
        paths = self.legacy_pair()
        def history_snapshot(sheet):
            values = {}
            for row in sheet:
                for cell in row:
                    if cell.value is None:
                        continue
                    value = cell.value
                    if hasattr(value, "text"):
                        value = (value.text, value.ref)
                    values[cell.coordinate] = value, cell.style_id
            return values, set(sheet.merged_cells.ranges), len(sheet.conditional_formatting), str(sheet.page_setup)

        original = load_workbook(paths[0])
        before = history_snapshot(original.active)
        original.close()

        @contextmanager
        def simulated_excel(path, title):
            book = load_workbook(path)
            try:
                sheet = book[title]
                literal = {cell.coordinate: cell.value for row in sheet for cell in row
                           if cell.data_type != "f" and cell.value is not None}
                if sheet["F8"].value is not None and sheet["A14"].value is None:
                    cash = sum(literal.get(f"F{row}", 0) for row in range(8, 24))
                    expenses = sum(literal.get(f"L{row}", 0) for row in range(8, 21))
                    literal.update(F31=cash, L31=expenses, L34=cash - expenses)
                else:
                    self.assertNotEqual(title, "December 2025", "Unissued fee must not be read from S:U")
                    expense = literal.get("G7", 0)
                    fee = literal.get("D8", 0)
                    tenant = expense + literal.get("M14", 0)
                    owner = fee + literal.get("N14", 0)
                    literal.update(D12=fee, G12=expense, L12=fee + expense, H14=expense,
                                   J14=expense, K14=fee, L14=fee + expense, O14=tenant + owner,
                                   H30=expense, J30=expense, K30=fee, L30=fee + expense, R11="OK")
                    for row in main.APARTMENT_ROWS:
                        literal.update({f"S{row}": tenant if row == 14 else 0,
                                        f"T{row}": owner if row == 14 else 0,
                                        f"U{row}": tenant + owner if row == 14 else 0})
                yield CalculatedSheet(title, literal)
            finally:
                book.close()

        for index, (paid, cleaning, expected) in enumerate(((0, 0, 940), (20000, 10000, 980), (10000, 4000, 980))):
            def bill_form(validate):
                data = dict.fromkeys(main.ADDITIONAL_BILLS, ("", 0))
                data["standard"] = dict.fromkeys(main.STANDARD_BILLS, 0)
                data["standard"]["Καθαρισμός κοιν/στων"] = cleaning
                data["allocations"] = validate(data)
                return dict(data, generate_pdf=False), "next"

            with (
                main.workbook_pair(*paths) as (common, ledger, revisions),
                patch.object(main, "calculated_sheet", simulated_excel),
                patch.object(main, "launch_dashboard", return_value=([paid], [False], "next")) as payments,
                patch.object(main, "launch_bills_dashboard", bill_form),
                patch.object(main.messagebox, "showinfo"),
                patch.object(main.messagebox, "askyesno", return_value=True),
            ):
                self.assertTrue(main.run_month(*paths, common, ledger, revisions))
                if index == 0:
                    self.assertEqual(payments.call_args.args[1:3], ([0], [0]))
            with main.workbook_pair(*paths) as (common, ledger, _), patch.object(main, "calculated_sheet", simulated_excel):
                self.assertEqual(main.read_closing_cash(paths[1], ledger.sheetnames[-1]), expected * 100)
                self.assertEqual(ledger.worksheets[-1]["L16"].value, 60)
                self.assertEqual(history_snapshot(common.worksheets[0]), before)

    def test_opening_debt_is_preserved_and_ambiguous_history_is_blocked(self):
        paths = self.legacy_pair()
        with main.workbook_pair(*paths) as (common, ledger, _):
            common.active["M14"], common.active["N14"] = -10, 25
            self.assertTrue(main.is_unissued_pair(common, ledger))
            self.assertEqual(main.opening_balances(common.active)[0], (-1000, 2500, 1500))
            ledger.active["E3"] = "Ambiguous report header"
            with self.assertRaisesRegex(ValueError, "cannot identify"):
                main.is_unissued_pair(common, ledger)

    def test_tenant_changes_preserve_debt_and_credit_on_cancel_or_accept(self):
        for amount in (25, -25):
            with self.subTest(amount=amount):
                paths = self.legacy_pair()
                with main.workbook_pair(*paths) as (common, ledger, revisions):
                    common.active["M14"] = amount
                    properties = [("", Decimal(10), "Tenant B", "Owner A", None)] + [
                        ("", None, "", "", None)
                    ] * 14
                    def form(ws, save, *, initial):
                        self.assertFalse(save(properties, 6000))
                        self.assertEqual(ws["D14"].value, "Tenant A")
                        self.assertEqual(ws["M14"].value, amount)
                        self.assertTrue(save(properties, 6000))
                    with patch.object(main, "launch_edit_dashboard", form), patch.object(
                        main.messagebox, "askyesno", side_effect=[False, True],
                    ) as dialog:
                        main.edit_properties(*paths, common, ledger, revisions)
                    self.assertEqual(dialog.call_count, 2)
                    self.assertEqual(dialog.call_args.kwargs["default"], "no")
                    self.assertEqual(common.active["M14"].value, amount)

    def test_marker_does_not_authorize_overwriting_custom_cell(self):
        paths = self.legacy_pair()
        with main.workbook_pair(*paths) as (common, ledger, _):
            main.set_expense_formulas(common.active)
            common.active["Z15"] = "Custom"
            with self.assertRaisesRegex(ValueError, "Z15"):
                main.validate_template_layout(common, ledger)

    def test_rollback_error_keeps_recovery_instructions_when_lock_cleanup_fails(self):
        books, paths = self.save_fixture()
        replace, unlink = os.replace, Path.unlink
        def fail_replace(src, dst):
            if dst == paths[1] or "rollback" in Path(src).name:
                raise PermissionError("replacement blocked")
            replace(src, dst)
        def fail_lock(path, *args, **kwargs):
            if path.suffix == ".lock":
                raise PermissionError("lock cleanup blocked")
            return unlink(path, *args, **kwargs)
        with patch.object(main.os, "replace", fail_replace), patch.object(Path, "unlink", fail_lock):
            with self.assertRaisesRegex(RuntimeError, "restore these backups") as error:
                main.save_workbooks_safely(books[0], paths[0], books[1], paths[1])
        self.assertIn("lock cleanup blocked", main.error_details(error.exception))
        self.assertTrue(main.transaction_path(paths[0]).is_file())
        self.assertTrue(list(self.folder.glob(".bem-rollback-common-*.xlsx")))

    def test_interrupted_pair_save_is_found_before_generic_validation(self):
        _, paths = self.save_fixture()
        before = [path.read_bytes() for path in paths]
        script = """
import os, sys
from pathlib import Path
from openpyxl import load_workbook
import main
a,b=map(Path, sys.argv[1:])
x,y=load_workbook(a),load_workbook(b)
x.active['A1']=y.active['A1']='New'
original=os.replace
def interrupted(src, dst):
    if dst == b:
        os._exit(42)
    original(src,dst)
main.os.replace=interrupted
main.save_workbooks_safely(x,a,y,b)
"""
        child = subprocess.run([sys.executable, "-B", "-c", script, *map(str, paths)], check=False)
        self.assertEqual(child.returncode, 42)
        self.assertNotEqual(paths[0].read_bytes(), before[0])
        self.assertEqual(paths[1].read_bytes(), before[1])
        journal = json.loads(main.transaction_path(paths[0]).read_text(encoding="utf-8"))
        self.assertEqual(Path(journal["workbooks"][0]["backup"]).read_bytes(), before[0])
        with self.assertRaisesRegex(ValueError, "incomplete operation") as issue:
            main.open_workbook_pair(*paths)
        self.assertIn(journal["workbooks"][1]["staged"], str(issue.exception))
        self.assertTrue(all(main.lock_path(path).exists() for path in paths))

    def test_failed_pdf_restoration_keeps_report_and_original_error(self):
        books, paths = self.save_fixture()
        pdf = self.folder / "Fictional.pdf"
        pdf.write_bytes(b"%PDF-Previous")
        replace = os.replace
        def fail_replace(src, dst):
            if dst == paths[1]:
                raise OSError("workbook replacement failed")
            if dst == pdf:
                raise OSError("PDF restore blocked")
            replace(src, dst)
        with patch.object(main.os, "replace", fail_replace), self.assertRaisesRegex(OSError, "workbook replacement failed") as error:
            main.save_workbooks_safely(books[0], paths[0], books[1], paths[1], delete_pdf=pdf)
        [staged] = self.folder.glob(".bem-delete-*.pdf")
        self.assertEqual(staged.read_bytes(), b"%PDF-Previous")
        self.assertIn(str(staged), main.error_details(error.exception))
        self.assertTrue(main.transaction_path(paths[0]).is_file())

    def test_failed_pdf_export_is_a_post_commit_warning_under_lock(self):
        books, paths = self.save_fixture()
        pdf = self.folder / "Fictional.pdf"
        def export():
            self.assertTrue(all(main.lock_path(path).is_file() for path in [*paths, pdf]))
            raise OSError("PDF locked")
        warnings = main.save_workbooks_safely(books[0], paths[0], books[1], paths[1],
                                              after_commit=export, pdf_destination=pdf)
        self.assertIn("Financial records were saved", warnings[0])
        self.assertIn("PDF locked", warnings[0])
        for path in paths:
            saved = load_workbook(path)
            self.assertEqual(saved.active["A1"].value, "Committed")
            saved.close()

    def test_existing_report_is_preserved_on_export(self):
        p = self.folder / "common.xlsx"
        pdf = main.month_pdf_path(p, main.parse_month_sheet("January 2026"))
        pdf.write_bytes(b"%PDF-Previous")
        with self.assertRaises(FileExistsError):
            main.export_month_pdf(p, main.parse_month_sheet("January 2026"), "January 2026")
        self.assertEqual(pdf.read_bytes(), b"%PDF-Previous")
