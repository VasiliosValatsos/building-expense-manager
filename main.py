#!/usr/bin/env python
# coding: utf-8

import hashlib
import json
import math
import os
import re
import shutil
import sys
import tempfile
import tkinter as tk
from contextlib import ExitStack, contextmanager
from copy import copy, deepcopy
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from fractions import Fraction
from functools import lru_cache
from math import gcd
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


GREEK_MONTHS = {
    1: "ΙΑΝΟΥΑΡΙΟΣ",
    2: "ΦΕΒΡΟΥΑΡΙΟΣ",
    3: "ΜΑΡΤΙΟΣ",
    4: "ΑΠΡΙΛΙΟΣ",
    5: "ΜΑΙΟΣ",
    6: "ΙΟΥΝΙΟΣ",
    7: "ΙΟΥΛΙΟΣ",
    8: "ΑΥΓΟΥΣΤΟΣ",
    9: "ΣΕΠΤΕΜΒΡΙΟΣ",
    10: "ΟΚΤΩΒΡΙΟΣ",
    11: "ΝΟΕΜΒΡΙΟΣ",
    12: "ΔΕΚΕΜΒΡΙΟΣ"
}


ENGLISH_MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
MONTH_NUMBERS = {month.lower(): index for index, month in enumerate(ENGLISH_MONTHS, 1)}

LANGUAGE = "en"


def tr(english, greek):
    """Choose the interface text without changing workbook data."""
    return greek if LANGUAGE == "el" else english


def set_language(language):
    global LANGUAGE, BILL_LABELS, FLOOR_LABELS, EXPENSE_LABELS, DISPLAY_MONTHS
    if language not in ("en", "el"):
        raise ValueError("Unsupported language")
    LANGUAGE = language
    BILL_LABELS = tr({'Φ.Α.': 'Natural gas / heating', 'Ηλεκτ. ρεύμα κοινοχρήστων': 'Common-area electricity', 'ΕΥΔΑΠ': 'Water', 'Καθαρισμός κοιν/στων': 'Common-area cleaning', 'Αναλώσιμα': 'Consumables'}, {'Φ.Α.': 'Φυσικό αέριο / θέρμανση', 'Ηλεκτ. ρεύμα κοινοχρήστων': 'Κοινόχρηστο ηλεκτρικό ρεύμα', 'ΕΥΔΑΠ': 'Ύδρευση (ΕΥΔΑΠ)', 'Καθαρισμός κοιν/στων': 'Καθαρισμός κοινοχρήστων', 'Αναλώσιμα': 'Αναλώσιμα'})
    FLOOR_LABELS = tr({'ΥΠ2': 'Basement 2', 'ΥΠ1': 'Basement 1', 'ΙΣ': 'Ground', "Α' ΟΡ": 'First floor', "Β' ΟΡ": 'Second floor', "Γ' ΟΡ": 'Third floor'}, {'ΥΠ2': 'ΥΠ2', 'ΥΠ1': 'ΥΠ1', 'ΙΣ': 'ΙΣ', "Α' ΟΡ": "Α' ΟΡ", "Β' ΟΡ": "Β' ΟΡ", "Γ' ΟΡ": "Γ' ΟΡ"})
    EXPENSE_LABELS = tr({'urgent': 'Owner expense 1', 'urgent2': 'Owner expense 2', 'extra1': 'Tenant expense 1', 'extra2': 'Tenant expense 2'}, {'urgent': 'Δαπάνη ιδιοκτητών 1', 'urgent2': 'Δαπάνη ιδιοκτητών 2', 'extra1': 'Δαπάνη ενοικιαστών 1', 'extra2': 'Δαπάνη ενοικιαστών 2'})
    DISPLAY_MONTHS = tr(list(ENGLISH_MONTHS), list(GREEK_MONTHS.values()))


set_language("en")

def get_app_data_path():
    """Return the per-user settings path for source and packaged builds."""
    if sys.platform == "win32":
        folder = os.environ.get("APPDATA") or os.path.join(Path.home(), "AppData", "Roaming")
        base = Path(folder) / "BuildingExpenseManager"
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / "BuildingExpenseManager"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "building-expense-manager"
    return str(base / "config_paths.json")


CONFIG_FILE = get_app_data_path()
APARTMENT_ROWS = range(14, 29)
OWNER_EXPENSES = [f"D{row}" for row in range(4, 12)]
TENANT_EXPENSES = [f"G{row}" for row in range(4, 12)]
PENDING_COLUMNS = ("DA", "DB", "DC", "DD", "DE")
STANDARD_BILLS = {
    "Φ.Α.": ("G4", "L10"),
    "Ηλεκτ. ρεύμα κοινοχρήστων": ("G5", "L9"),
    "ΕΥΔΑΠ": ("G6", "L11"),
    "Καθαρισμός κοιν/στων": ("G7", "L8"),
    "Αναλώσιμα": ("G8", "L12"),
}
ADDITIONAL_BILLS = {
    "urgent": ("A6", "D6", "K17", "L17"),
    "urgent2": ("A7", "D7", "K18", "L18"),
    "extra1": ("E9", "G9", "K13", "L13"),
    "extra2": ("E10", "G10", "K14", "L14"),
}
MANAGEMENT_FEE_CELLS = ("D8", "L16")
EXPENSE_PAIRS = [
    *STANDARD_BILLS.values(),
    *((charge, cash) for _, charge, _, cash in ADDITIONAL_BILLS.values()),
    MANAGEMENT_FEE_CELLS,
]
AREA_SCALE = 100_000
HEATING_SCALE = 10_000
EXACT_EXCEL_INTEGER = 999_999_999_999_999


def month_sheet_name(date: datetime) -> str:
    return f"{ENGLISH_MONTHS[date.month - 1]} {date.year}"


def parse_month_sheet(name: str) -> datetime:
    """Parse the English month name even when the app runs in Greek."""
    match = re.fullmatch(r"([A-Za-z]+) ([0-9]{4})", name.strip())
    if not match or match.group(1).lower() not in MONTH_NUMBERS:
        raise ValueError(tr(
            f"Invalid sheet name format: {name!r}; expected 'January 2026'",
            f"Μη έγκυρο όνομα φύλλου: {name!r}; αναμένεται 'January 2026'",
        ))
    return datetime(int(match.group(2)), MONTH_NUMBERS[match.group(1).lower()], 1)


def to_float(value, *, blank_zero=True):
    """Accept 75,50 and 75.50; reject formulas, errors and non-finite values."""
    if value is None or (isinstance(value, str) and not value.strip()):
        if blank_zero:
            return 0.0
        raise ValueError(tr('Enter a number.', 'Συμπληρώστε έναν αριθμό.'))
    if isinstance(value, bool):
        raise ValueError(tr('Boolean is not a numeric value', 'Η λογική τιμή δεν είναι αριθμός'))
    try:
        result = float(value.strip().replace(",", ".") if isinstance(value, str) else value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(tr(f'Invalid numeric value: {value!r}', f'Μη έγκυρη αριθμητική τιμή: {value!r}')) from exc
    if not math.isfinite(result):
        raise ValueError(tr('Number must be finite', 'Ο αριθμός πρέπει να είναι πεπερασμένος'))
    return result


def decimal_number(value):
    try:
        number = Decimal(str(value).strip().replace(",", "."))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(tr(f"Invalid number: {value!r}", f"Μη έγκυρος αριθμός: {value!r}")) from exc
    if not number.is_finite():
        raise ValueError(tr("Enter a finite number.", "Ο αριθμός πρέπει να είναι πεπερασμένος."))
    return number


def money_cents(value, *, blank_zero=False, calculated=False):
    if blank_zero and (value is None or str(value).strip() == ""):
        return 0
    amount = decimal_number(value)
    if abs(amount) > 9_999_999_999_999:
        raise ValueError(tr("This amount is too large for Excel to keep cent precision.", "Το ποσό είναι πολύ μεγάλο για να το αποθηκεύσει το Excel με ακρίβεια λεπτού."))
    rounded = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if not calculated and amount != rounded:
        raise ValueError(tr("Use up to two decimal places for amounts.", "Για ποσά χρησιμοποιήστε έως δύο δεκαδικά ψηφία."))
    return int(rounded * 100)


def euros(cents):
    return Decimal(cents).scaleb(-2)


def allocate_cents(amount, weights):
    weights = [Fraction(weight) for weight in weights]
    total = sum(weights)
    if amount < 0 or any(weight < 0 for weight in weights):
        raise ValueError(tr("Expense amounts and allocation weights must be zero or positive.", "Τα ποσά και οι συντελεστές επιμερισμού πρέπει να είναι μηδέν ή θετικοί."))
    if not amount:
        return [0] * len(weights)
    if not total:
        raise ValueError(tr("No eligible apartments for this expense.", "Καμία ιδιοκτησία δεν συμμετέχει σε αυτή τη δαπάνη."))
    quotas = [divmod(amount * weight, total) for weight in weights]
    whole = [int(part) for part, _ in quotas]
    eligible = sorted(
        (index for index, weight in enumerate(weights) if weight),
        key=lambda index: (-quotas[index][1], index),
    )
    for index in eligible[:amount - sum(whole)]:
        whole[index] += 1
    return whole


@contextmanager
def excel_app():
    """Open a desktop Excel session for recalculation or export."""
    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(tr('xlwings is required. Run: python -m pip install -r requirements.txt', 'Απαιτείται η βιβλιοθήκη xlwings. Εκτελέστε: python -m pip install -r requirements.txt')) from exc
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        yield app
    finally:
        app.quit()


def excel_number(path, sheet, address, *, blank_zero=False):
    try:
        value = sheet.range(address).options(err_to_str=True).value
        if value is None and blank_zero:
            return Decimal(0)
        if not isinstance(value, (int, float, Decimal)) or isinstance(value, bool):
            raise ValueError(tr(
                f"Expected a numeric Excel value, found {value!r}.",
                f"Αναμενόταν αριθμητική τιμή Excel, βρέθηκε {value!r}.",
            ))
        number = decimal_number(value)
        money_cents(number, calculated=True)
        return number
    except ValueError as exc:
        raise ValueError(f"{Path(path).name} / {sheet.name}!{address}: {exc}") from exc


@contextmanager
def calculated_sheet(path, sheet_name):
    with excel_app() as app:
        book = app.books.open(os.path.abspath(path), update_links=False, read_only=True)
        try:
            app.calculate()
            yield book.sheets[sheet_name]
        finally:
            book.close()


def read_debts_from_excel(path, sheet_name, rows):
    with calculated_sheet(path, sheet_name) as sheet:
        return calculated_balances(path, sheet, rows)


def calculated_balances(path, sheet, rows):
    result = []
    for row in rows:
        tenant, owner, total = (excel_number(path, sheet, f"{col}{row}") for col in "STU")
        verify_calculated_total(path, sheet, f"U{row}", total, tenant + owner)
        result.append(tuple(money_cents(amount, calculated=True) for amount in (tenant, owner, total)))
    return result


def reconcile_expenses(path, common, path2, ledger, *, detailed, expected_allocations=None):
    def amount(book, sheet, cell):
        return excel_number(book, sheet, cell, blank_zero=True)

    def agree(left, right, context):
        delta = money_cents(left - right, calculated=True)
        if delta:
            raise ValueError(tr(
                f"{common.name} / {context}: discrepancy EUR {euros(delta):+.2f}. Check the amounts for this month.",
                f"{common.name} / {context}: διαφορά EUR {euros(delta):+.2f}. Ελέγξτε τα ποσά του μήνα.",
            ))

    def c(cell):
        return amount(path, common, cell)

    def loc(cell):
        return f"{Path(path).name}!{cell}"

    labels = {slot[0]: BILL_LABELS[key] for key, slot in STANDARD_BILLS.items()}
    labels.update({slot[1]: EXPENSE_LABELS[key] for key, slot in ADDITIONAL_BILLS.items()})
    labels["D8"] = tr("Management fee", "Αμοιβή διαχείρισης")
    for charge, cash in EXPENSE_PAIRS:
        agree(c(charge), amount(path2, ledger, cash),
              f"{labels[charge]} · {loc(charge)} ↔ {Path(path2).name}!{cash}")
    owner = sum(map(c, OWNER_EXPENSES))
    tenant = sum(map(c, TENANT_EXPENSES))
    for cell, expected in (("D12", owner), ("G12", tenant), ("L12", owner + tenant)):
        agree(c(cell), expected, loc(cell))
    ledger_cells = range(8, 21)
    agree(owner + tenant, sum(amount(path2, ledger, f"L{row}") for row in ledger_cells),
          f"{loc('D12/G12')} ↔ {Path(path2).name}!L8:L20")

    columns = "HIJKLOMNSTU"
    rows = {row: {col: c(f"{col}{row}") for col in columns} for row in APARTMENT_ROWS}
    if detailed:
        expected = {row: [Decimal(0)] * 3 for row in APARTMENT_ROWS}
        for index, cell in enumerate(OWNER_EXPENSES + TENANT_EXPENSES):
            column = get_column_letter(25 + 3 * index)
            shares = [euros(money_cents(c(f"{column}{row}") / 100)) for row in APARTMENT_ROWS]
            agree(c(cell), sum(shares), loc(f"{cell} ↔ {column}14:{column}28"))
            if expected_allocations is not None:
                for row, share, expected_share in zip(APARTMENT_ROWS, shares, expected_allocations[cell]):
                    agree(share, euros(expected_share), loc(f"{cell} ↔ {column}{row}"))
            group = 0 if cell in OWNER_EXPENSES else 1 if cell == "G4" else 2
            for row, share in zip(APARTMENT_ROWS, shares):
                expected[row][group] += share
        for row, shares in expected.items():
            for col, share in zip("KIH", shares):
                agree(rows[row][col], share, loc(f"{col}{row} ↔ allocation"))
    for col, expected in (("H", tenant - c("G4")), ("I", c("G4")), ("J", tenant), ("K", owner), ("L", owner + tenant)):
        agree(sum(values[col] for values in rows.values()), expected, loc(f"{col}14:{col}28"))
        agree(c(f"{col}30"), expected, loc(f"{col}30"))
    for row, v in rows.items():
        for col, expected in (("J", v["H"] + v["I"]), ("L", v["J"] + v["K"]),
                              ("S", v["J"] + v["M"]), ("T", v["K"] + v["N"]), ("U", v["S"] + v["T"])):
            agree(v[col], expected, loc(f"{col}{row}"))
        # A historical inactive row can hide its report balance. Its S:U
        # remain visible to the transition check.
        if v["L"] or v["O"] or (v["U"] and common.range(f"C{row}").value):
            agree(v["O"], v["U"], loc(f"O{row} ↔ U{row}"))
    if detailed and common.range("R11").value != "OK":
        raise ValueError(tr(f"{common.name} / {loc('R11')}: reconciliation warning.",
                            f"{common.name} / {loc('R11')}: προειδοποίηση συμφωνίας."))


def opening_balances(common):
    result = []
    for row in APARTMENT_ROWS:
        try:
            tenant, owner = (money_cents(common[f"{col}{row}"].value, blank_zero=True) for col in "MN")
        except ValueError as exc:
            raise ValueError(f"{common.title} / M{row}:N{row}: {exc}") from exc
        result.append((tenant, owner, tenant + owner))
    return result


def is_unissued_pair(wb, wb2):
    common, ledger = wb.worksheets[-1], wb2.worksheets[-1]
    date = parse_month_sheet(common.title)
    title = f"{GREEK_MONTHS[date.month]} {date.year}"
    if common["O1"].value == ledger["E3"].value == title and not common["DF3"].value:
        return False
    if len(wb.worksheets) == 1 and is_opening_sheet(common, ledger):
        return True
    raise ValueError(tr(
        f"{common.title}: cannot identify an unissued opening sheet or a matching report. Review O1, E3 and the original cash movements before migration.",
        f"{common.title}: δεν αναγνωρίζεται αρχικό φύλλο ή εκδομένη αναφορά. Ελέγξτε τα O1, E3 και τις κινήσεις ταμείου πριν τη μετάβαση.",
    ))


def previous_finances(path, path2, wb, wb2):
    common, ledger = wb.worksheets[-1], wb2.worksheets[-1]
    unissued = is_unissued_pair(wb, wb2)
    if unissued:
        return opening_balances(common), read_closing_cash(path2, ledger.title)
    detailed = common["Y13"].value == "D4"
    expected = monthly_allocations(common, {}) if common["B36"].value == allocation_layout(False)["B36"] else None
    return reconcile_issued_period(path, path2, common.title, detailed=detailed, expected_allocations=expected)


def reconcile_issued_period(path, path2, title, *, detailed=True, expected_allocations=None):
    with calculated_sheet(path, title) as report, calculated_sheet(path2, title) as cash:
        reconcile_expenses(path, report, path2, cash, detailed=detailed, expected_allocations=expected_allocations)
        return calculated_balances(path, report, APARTMENT_ROWS), calculated_closing_cash(path2, cash)


def read_closing_cash(path, sheet_name):
    with calculated_sheet(path, sheet_name) as sheet:
        return calculated_closing_cash(path, sheet)


def calculated_closing_cash(path, sheet):
    closing = excel_number(path, sheet, "L34")
    income_with_opening = excel_number(path, sheet, "F31")
    expenses = excel_number(path, sheet, "L31")
    opening = excel_number(path, sheet, "F8")
    income = sum(excel_number(path, sheet, f"F{row}", blank_zero=True) for row in range(9, 24))
    outflow = sum(excel_number(path, sheet, f"L{row}", blank_zero=True) for row in range(8, 21))
    verify_calculated_total(path, sheet, "F31", income_with_opening, opening + income)
    verify_calculated_total(path, sheet, "L31", expenses, outflow)
    verify_calculated_total(path, sheet, "L34", closing, opening + income - outflow)
    return money_cents(closing, calculated=True)


def verify_calculated_total(path, sheet, address, actual, expected):
    # Older workbooks can contain fractions of a cent. Compare rounded totals.
    if abs(actual - expected) > Decimal("0.0000001") and money_cents(actual, calculated=True) != money_cents(expected, calculated=True):
        raise ValueError(tr(
            f"{Path(path).name} / {sheet.name}!{address}: the calculated total does not reconcile.",
            f"{Path(path).name} / {sheet.name}!{address}: το υπολογισμένο σύνολο δεν συμφωνεί.",
        ))


def literal_text(sheet, address, value):
    if value and len(value) > 32767:
        raise ValueError(tr("Text exceeds the Excel cell limit.", "Το κείμενο υπερβαίνει το όριο του κελιού Excel."))
    cell = sheet[address]
    cell.value = value or None
    if value:
        cell.data_type = "s"


def normalize_name(value):
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValueError(tr("Names must be text.", "Τα ονόματα πρέπει να είναι κείμενο."))
    return " ".join(value.split())


def excel_input_number(value, scale):
    if value is None or not str(value).strip():
        return None
    number = decimal_number(value)
    if abs(number) > 999_999 or number != number.quantize(Decimal(1) / scale):
        raise ValueError(tr(
            f"Use a value up to 999999 with at most {len(str(scale)) - 1} decimal places. Existing values are not rounded.",
            f"Χρησιμοποιήστε τιμή έως 999999 με έως {len(str(scale)) - 1} δεκαδικά. Οι υπάρχουσες τιμές δεν στρογγυλοποιούνται.",
        ))
    return number


def apartment_details(area, owner, factor):
    owner = normalize_name(owner)
    area = excel_input_number(area, AREA_SCALE)
    factor = excel_input_number(factor, HEATING_SCALE)
    if (area is not None and area <= 0) or (area is not None) != bool(owner):
        raise ValueError(tr(
            "Enter a positive area and an owner together, or leave both blank.",
            "Συμπληρώστε θετικά τετραγωνικά και ιδιοκτήτη μαζί ή αφήστε και τα δύο κενά.",
        ))
    if factor is not None and not 0 <= factor <= 1:
        raise ValueError(tr("Heating factor must be between 0 and 1.", "Ο συντελεστής θέρμανσης πρέπει να είναι από 0 έως 1."))
    return area, owner, factor


def apartment_weights(sheet):
    general, heating = [], []
    for row in APARTMENT_ROWS:
        try:
            area, _, factor = apartment_details(
                sheet[f"C{row}"].value, sheet[f"E{row}"].value, sheet[f"H{row + 22}"].value,
            )
        except ValueError as exc:
            raise ValueError(f"{sheet.title} / #{row - 13}: {exc}") from exc
        general.append(int((area or 0) * AREA_SCALE))
        heating.append(general[-1] * int((1 if factor is None else factor) * HEATING_SCALE))
    return tuple([weight // max(1, gcd(*weights)) for weight in weights] for weights in (general, heating))


def validate_balance_transition(previous_owner, next_owner, tenant, owner):
    if normalize_name(previous_owner) != normalize_name(next_owner) and (tenant or owner):
        raise ValueError(tr(
            "Before removing or changing the owner, settle both balances or record how they were reconciled.",
            "Πριν αφαιρέσετε ή αλλάξετε τον ιδιοκτήτη, τακτοποιήστε και τα δύο υπόλοιπα ή καταγράψτε πώς τα συμφωνήσατε.",
        ))


def confirm_legacy_rounding(balances):
    discrepancies = [
        f"#{row - 13}: {euros(total):.2f} → {euros(tenant + owner):.2f}"
        for row, (tenant, owner, total) in zip(APARTMENT_ROWS, balances)
        if tenant + owner != total
    ]
    if not discrepancies:
        return True
    return messagebox.askyesno(
        tr("Legacy balance rounding", "Στρογγυλοποίηση παλαιών υπολοίπων"),
        tr(
            "Rounding the tenant and owner balances separately changes the totals below. "
            "Continue using the two rounded balances?\n\n",
            "Αν στρογγυλοποιηθούν χωριστά τα υπόλοιπα ενοικιαστή και ιδιοκτήτη, αλλάζουν τα παρακάτω σύνολα. "
            "Θέλετε να συνεχίσετε με τα στρογγυλοποιημένα υπόλοιπα;\n\n",
        ) + "\n".join(discrepancies), default="no",
    )


def confirm_tenant_changes(rows):
    if not rows:
        return True
    return messagebox.askyesno(
        tr("Tenant change with a balance", "Αλλαγή ενοικιαστή με υπόλοιπο"),
        tr(
            "The existing balance stays with the apartment and will appear on the row with the new tenant name. "
            "This entry does not transfer legal liability. Continue?\n\n",
            "Το υπόλοιπο παραμένει στην ιδιοκτησία και θα εμφανίζεται στη γραμμή με το νέο όνομα ενοικιαστή. "
            "Η καταχώριση δεν μεταβιβάζει νομική ευθύνη. Συνέχεια;\n\n",
        ) + "\n".join(f"#{row - 13}: EUR {euros(balance):+.2f}" for row, balance in rows),
        default="no",
    )


def workbook_revision(path):
    with open(path, "rb") as source:
        return hashlib.file_digest(source, "sha256").digest()


def open_workbook_pair(path, path2):
    check_recovery(tuple(Path(item).resolve() for item in (path, path2)))
    with ExitStack() as cleanup:
        revisions = tuple(workbook_revision(item) for item in (path, path2))
        wb = load_workbook(path)
        cleanup.callback(wb.close)
        wb2 = load_workbook(path2)
        cleanup.callback(wb2.close)
        verify_revisions((path, path2), revisions)
        if wb.sheetnames != wb2.sheetnames:
            raise ValueError(tr("The workbooks have different month histories.", "Τα βιβλία έχουν διαφορετικό ιστορικό μηνών."))
        validate_template_layout(wb, wb2)
        parse_month_sheet(wb.sheetnames[-1])
        addresses = [normalize_name(book.worksheets[-1][cell].value).casefold()
                     for book, cell in ((wb, "E1"), (wb2, "K3"))]
        if all(addresses) and addresses[0] != addresses[1]:
            raise ValueError(tr("The workbooks show different building addresses.", "Τα βιβλία έχουν διαφορετικές διευθύνσεις πολυκατοικίας."))
        cleanup.pop_all()
        return wb, wb2, revisions


@contextmanager
def workbook_pair(path, path2):
    wb, wb2, revisions = open_workbook_pair(path, path2)
    try:
        yield wb, wb2, revisions
    finally:
        wb.close()
        wb2.close()


def verify_revisions(paths, revisions):
    for path, revision in zip(paths, revisions):
        if workbook_revision(path) != revision:
            raise ValueError(tr(
                f"{path} changed on disk. Reopen the workbooks before saving.",
                f"Το αρχείο {path} άλλαξε στον δίσκο. Ανοίξτε ξανά τα βιβλία πριν από την αποθήκευση.",
            ))


def lock_path(path):
    return path.with_name(f".bem-{path.name}.lock")


def transaction_path(path):
    return path.with_name(f".bem-{path.name}.transaction.json")


def check_recovery(paths):
    markers = [marker for path in paths for marker in (lock_path(path), transaction_path(path))
               if marker.exists()]
    entries = []
    for marker in markers:
        if marker.suffix == ".json":
            try:
                entries.append(json.dumps(json.loads(marker.read_text(encoding="utf-8")), ensure_ascii=False, indent=2))
            except (OSError, ValueError):
                pass
    if markers:
        candidates = sorted({str(item) for path in paths for item in path.parent.glob(".bem-*")})
        raise ValueError(tr(
            "A workbook is being updated or an incomplete operation needs recovery. Close Excel and check whether another copy of the app is running. Keep both original workbooks and every file listed below. Check the transaction record before you decide which files to restore or remove.",
            "Γίνεται ενημέρωση βιβλίου ή χρειάζεται επαναφορά μετά από διακοπή. Κλείστε το Excel και ελέγξτε αν είναι ανοιχτό άλλο παράθυρο της εφαρμογής. Κρατήστε τα δύο αρχικά βιβλία και όλα τα αρχεία της λίστας. Δείτε το αρχείο συναλλαγής πριν αποφασίσετε ποια αρχεία θα επαναφέρετε ή θα αφαιρέσετε.",
        ) + "\n" + "\n".join([*(str(path) for path in paths), *candidates, *entries]))


def cleanup_file(path, warnings):
    try:
        path.unlink(missing_ok=True)
    except OSError as exc:
        warnings.append(tr(f"Could not remove {path} after the operation: {exc}", f"Το {path} έμεινε στον φάκελο μετά την εργασία: {exc}"))


def error_details(exc):
    return "\n".join([str(exc), *getattr(exc, "__notes__", ())]).strip()


@contextmanager
def workbook_locks(paths, warnings=None):
    warnings = [] if warnings is None else warnings
    acquired = []
    try:
        for path in sorted(paths):
            lock = lock_path(path)
            try:
                fd = os.open(lock, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            except FileExistsError as exc:
                raise ValueError(tr(f"{path.name} is being updated: {lock}", f"Το {path.name} ενημερώνεται: {lock}")) from exc
            acquired.append(lock)
            os.close(fd)
        yield
    except BaseException as exc:
        for lock in reversed(acquired):
            cleanup_file(lock, warnings)
        for warning in warnings:
            exc.add_note(warning)
        raise
    else:
        for lock in reversed(acquired):
            cleanup_file(lock, warnings)


def save_workbooks_safely(wb, path, wb2, path2, *, revisions=None, delete_pdf=None, pdf_destination=None,
                          before_commit=None, after_commit=None):
    paths = [Path(path).resolve(), Path(path2).resolve()]
    if os.path.samefile(*paths):
        raise ValueError(tr('The two workbooks must be different files', 'Τα δύο αρχεία Excel πρέπει να είναι διαφορετικά'))
    for path in paths:
        protect_template(path)
    if wb.sheetnames != wb2.sheetnames:
        raise ValueError(tr("The workbooks have different month histories.", "Τα βιβλία έχουν διαφορετικό ιστορικό μηνών."))
    report = delete_pdf if delete_pdf is not None else pdf_destination
    locked_paths = paths + ([report.absolute()] if report is not None else [])
    check_recovery(locked_paths)
    warnings = []
    with workbook_locks(locked_paths, warnings):
        if any(transaction_path(path).exists() for path in paths):
            check_recovery(locked_paths)
        save_locked_workbooks((wb, wb2), paths, revisions, warnings, delete_pdf, before_commit)
        if after_commit is not None:
            try:
                after_commit()
            except (Exception, KeyboardInterrupt) as exc:
                warnings.append(tr(
                    "Financial records were saved. PDF export failed. Open the saved monthly sheet in the common expenses workbook and export A1:P31 to PDF. Do not re-enter payments.\n",
                    "Οι οικονομικές εγγραφές αποθηκεύτηκαν. Η εξαγωγή PDF απέτυχε. Ανοίξτε στο Excel το αποθηκευμένο φύλλο του μήνα στα κοινόχρηστα και εξαγάγετε την περιοχή A1:P31 σε PDF. Μην καταχωρίσετε ξανά πληρωμές.\n",
                ) + error_details(exc))
    return warnings


def stage_path(original, prefix):
    fd, filename = tempfile.mkstemp(prefix=prefix, suffix=original.suffix, dir=original.parent)
    os.close(fd)
    return Path(filename)


def save_locked_workbooks(workbooks, paths, revisions, warnings, delete_pdf, before_commit):
    staged, backups, staged_revisions = [], [], []
    preserve = set()
    journal = transaction_path(paths[0])
    pdf_stage = None
    pdf_moved = False
    journal_created = False
    revisions = revisions if revisions is not None else tuple(workbook_revision(path) for path in paths)
    try:
        verify_revisions(paths, revisions)
        for workbook, original in zip(workbooks, paths):
            staged.append(stage_path(original, ".bem-new-"))
            workbook.save(staged[-1])
            staged_revisions.append(workbook_revision(staged[-1]))
            backups.append(stage_path(original, f".bem-rollback-{original.stem}-"))
            shutil.copy2(original, backups[-1])
            for prepared in (staged[-1], backups[-1]):
                with prepared.open("r+b") as file:
                    os.fsync(file.fileno())
        if before_commit is not None:
            before_commit(*staged)
        if delete_pdf is not None and os.path.lexists(delete_pdf):
            if not delete_pdf.is_file() or delete_pdf.is_symlink():
                raise ValueError(tr(f"Invalid PDF: {delete_pdf}", f"Μη έγκυρο PDF: {delete_pdf}"))
            pdf_stage = stage_path(delete_pdf, ".bem-delete-")
        with journal.open("x", encoding="utf-8") as record:
            journal_created = True
            json.dump({
                "workbooks": [
                    {"original": str(path), "staged": str(new), "backup": str(old),
                     "before_sha256": before.hex(), "prepared_sha256": after.hex()}
                    for path, new, old, before, after in zip(paths, staged, backups, revisions, staged_revisions)
                ],
                "pdf": {"original": str(delete_pdf), "staged": str(pdf_stage)} if pdf_stage else None,
            }, record, ensure_ascii=False, indent=2)
            record.flush()
            os.fsync(record.fileno())
        verify_revisions(paths, revisions)
        if pdf_stage is not None:
            os.replace(delete_pdf, pdf_stage)
            pdf_moved = True
        for new, original, revision in zip(staged, paths, revisions):
            verify_revisions((original,), (revision,))
            os.replace(new, original)
    except BaseException as exc:
        failures = []
        for new, backup, original, prepared in zip(staged, backups, paths, staged_revisions):
            if not new.exists():
                try:
                    verify_revisions((original,), (prepared,))
                    os.replace(backup, original)
                except BaseException:
                    preserve.add(backup)
                    failures.append(f"{backup} → {original}")
        if pdf_moved:
            try:
                if delete_pdf.exists():
                    raise FileExistsError(delete_pdf)
                os.replace(pdf_stage, delete_pdf)
            except OSError as restore_error:
                preserve.add(pdf_stage)
                exc.add_note(f"{pdf_stage} → {delete_pdf}: {restore_error}")
        if failures:
            error = RuntimeError(tr(
                "Save and rollback failed. Close Excel and restore these backups before continuing:\n",
                "Απέτυχαν η αποθήκευση και η επαναφορά. Κλείστε το Excel και επαναφέρετε τα αντίγραφα πριν συνεχίσετε:\n",
            ) + "\n".join(failures) + f"\n{journal}")
            for note in getattr(exc, "__notes__", ()):
                error.add_note(note)
            raise error from exc
        raise
    finally:
        for temporary in staged + backups + ([pdf_stage] if pdf_stage is not None else []):
            if temporary not in preserve and not (preserve and temporary in staged):
                cleanup_file(temporary, warnings)
        if journal_created and not preserve:
            cleanup_file(journal, warnings)


def validate_template_layout(wb, wb2):
    """Check the layout and reserved cells of the 15-apartment templates."""
    common = wb[wb.sheetnames[-1]]
    ledger = wb2[wb2.sheetnames[-1]]
    if [common[f"A{row}"].value for row in APARTMENT_ROWS] != list(range(1, 16)):
        raise ValueError(tr("Expected apartment numbers 1–15 in A14:A28.", "Αναμένονται αριθμοί ιδιοκτησιών 1–15 στα A14:A28."))
    validate_fee_slot(ledger)
    validate_allocation_slots(common)
    if common["DF1"].value != "NEXT_MONTH":
        staged = (common[f"{col}{row}"] for col in PENDING_COLUMNS for row in APARTMENT_ROWS)
        if common["DF1"].value is not None or common["DF2"].value is not None or any(cell.value is not None for cell in staged):
            raise ValueError(tr(
                "The workbook contains custom data in cells reserved for next-month settings.",
                "Το βιβλίο έχει πρόσθετα στοιχεία σε κελιά ρυθμίσεων του επόμενου μήνα.",
            ))
    if common["DF3"].value not in (None, "OPENING"):
        raise ValueError(tr("Cell DF3 is reserved for setup.", "Το κελί DF3 προορίζεται για την αρχική ρύθμιση."))
    if common["DF1"].value == "NEXT_MONTH":
        group = common.column_dimensions.get("DA")
        if not group or not group.hidden or group.max != column_index_from_string("DF"):
            raise ValueError(tr("DA:DF do not match the saved settings layout.", "Τα DA:DF δεν συμφωνούν με τη διάταξη ρυθμίσεων."))
        fee = money_cents(common["DF2"].value)
        if fee < 0:
            raise ValueError(tr("Invalid scheduled fee in DF2.", "Μη έγκυρη προγραμματισμένη αμοιβή στο DF2."))
        for row in APARTMENT_ROWS:
            for col in PENDING_COLUMNS:
                cell = common[f"{col}{row}"]
                if cell.data_type == "f" or (col in ("DA", "DC", "DD") and cell.value is not None and cell.data_type != "s") or (
                    col in ("DB", "DE") and cell.value is not None and cell.data_type != "n"
                ):
                    raise ValueError(tr(f"Unexpected data in {cell.coordinate}.", f"Μη αναμενόμενα δεδομένα στο {cell.coordinate}."))
    for row in range(1, common.max_row + 1):
        for col in (*PENDING_COLUMNS, "DF"):
            if (row in APARTMENT_ROWS and col != "DF") or (row <= 3 and col == "DF"):
                continue
            cell = common[f"{col}{row}"]
            if cell.value is not None:
                raise ValueError(tr(f"Reserved settings cell {cell.coordinate} is occupied.",
                                    f"Το δεσμευμένο κελί ρυθμίσεων {cell.coordinate} περιέχει δεδομένα."))
    for address in ("D12", "G12", "L12", "I54"):
        if common[address].data_type != "f":
            raise ValueError(tr(f"Expected a common-expense formula in {address}.", f"Αναμένεται τύπος κοινοχρήστων στο {address}."))
    for row in range(14, 29):
        for col in ("S", "T", "U"):
            cell = common[f"{col}{row}"]
            if cell.data_type != "f":
                raise ValueError(
                    tr(f'Invalid ΚΟΙΝΟΧΡΗΣΤΑ template: expected a formula in {cell.coordinate}', f'Μη έγκυρο πρότυπο ΚΟΙΝΟΧΡΗΣΤΑ: απαιτείται τύπος στο {cell.coordinate}')
                )
    for address in ("F31", "L31", "L34"):
        if ledger[address].data_type != "f":
            raise ValueError(tr(f'Invalid ΒΙΒΛΙΟ template: expected a formula in {address}', f'Μη έγκυρο πρότυπο ΒΙΒΛΙΟ: απαιτείται τύπος στο {address}'))


def validate_workbook_pair(wb, wb2):
    if not wb.sheetnames or not wb2.sheetnames:
        raise ValueError(tr('Both workbooks must contain at least one sheet', 'Κάθε αρχείο πρέπει να περιέχει τουλάχιστον ένα φύλλο'))
    first, second = wb.sheetnames[-1], wb2.sheetnames[-1]
    if wb.sheetnames != wb2.sheetnames:
        raise ValueError(tr(f'Last sheets do not match: {first!r} vs {second!r}', f'Τα τελευταία φύλλα δεν συμφωνούν: {first!r} έναντι {second!r}'))
    expected_next = month_sheet_name(parse_month_sheet(first) + relativedelta(months=1))
    if expected_next.casefold() in {name.strip().casefold() for name in wb.sheetnames}:
        raise ValueError(tr(f'Sheet {expected_next!r} already exists', f'Το φύλλο {expected_next!r} υπάρχει ήδη'))
    validate_template_layout(wb, wb2)
    return first


def fix_heating_formulas(sheet, *, legacy=False):
    for row in APARTMENT_ROWS:
        r = row + 22
        owner = f'TRIM(SUBSTITUTE(CLEAN(E{row}),CHAR(160)," "))'
        active = f'AND(ISNUMBER(C{row}),C{row}>0,ISTEXT(E{row}),LEN({owner})>0)'
        if legacy:
            sheet[f"B{r}"] = f'=IF({active},C{row},0)'
            sheet[f"C{r}"] = f'=IF(B{r}=0,0,B{r}*IF(H{r}="",1,IF(AND(ISNUMBER(H{r}),H{r}>=0,H{r}<=1),H{r},NA())))'
        else:
            valid = f"AND({active},C{row}<=999999,ROUND(C{row},5)=C{row})"
            sheet[f"B{r}"] = f'=IF(AND(C{row}="",LEN({owner})=0),0,IF({valid},ROUND(C{row}*{AREA_SCALE},0),NA()))'
            sheet[f"C{r}"] = f'=IF(B{r}=0,0,B{r}*ROUND(IF(H{r}="",1,IF(AND(ISNUMBER(H{r}),H{r}>=0,H{r}<=1,ROUND(H{r},4)=H{r}),H{r},NA()))*{HEATING_SCALE},0))'
        general_total = "$F$54" if legacy else "SUM($B$36:$B$50)"
        heating_total = "$I$54" if legacy else "SUM($C$36:$C$50)"
        sheet[f"F{r}"] = f'=IF(B{r}>0,B{r}/{general_total}*100,"")'
        sheet[f"I{r}"] = f'=IF(B{r}>0,IF({heating_total}>0,C{r}/{heating_total}*100,0),"")'
    sheet["F54"] = "=SUM(B36:B50)" if legacy else f"=SUM(B36:B50)/{AREA_SCALE}"
    sheet["I54"] = "=SUM(C36:C50)" if legacy else f"=SUM(C36:C50)/{AREA_SCALE * HEATING_SCALE}"


def hide_working_columns(sheet, first, last):
    for key, dimension in list(sheet.column_dimensions.items()):
        start = dimension.min or column_index_from_string(key)
        end = dimension.max or start
        if start > last or end < first:
            continue
        del sheet.column_dimensions[key]
        for left, right in ((start, first - 1), (last + 1, end)):
            if start <= left <= right <= end:
                remaining = copy(dimension)
                remaining.index = get_column_letter(left)
                remaining.min, remaining.max = left, right
                sheet.column_dimensions[remaining.index] = remaining
    sheet.column_dimensions.group(get_column_letter(first), get_column_letter(last), outline_level=0, hidden=True)


def set_expense_formulas(sheet, allocations=None, *, legacy=False):
    """Write the weight, allocation and report cells for the common workbook."""
    fix_heating_formulas(sheet, legacy=legacy)
    hide_working_columns(sheet, 23, 70)
    allocations = allocations if allocations is not None else {}
    result_columns = {}
    for index, address in enumerate(OWNER_EXPENSES + TENANT_EXPENSES):
        first_column = 23 + index * 3
        quota, whole, result = (get_column_letter(first_column + offset) for offset in range(3))
        result_columns[address] = result
        sheet[f"{result}13"] = address
        if address in allocations:
            for row in APARTMENT_ROWS:
                sheet[f"{quota}{row}"] = None
                sheet[f"{whole}{row}"] = None
                sheet[f"{result}{row}"] = allocations[address][row - APARTMENT_ROWS.start]
        else:
            write_template_allocation(sheet, address, quota, whole, result, legacy=legacy)
    write_expense_report(sheet, result_columns, legacy=legacy)


def write_template_allocation(sheet, address, quota, whole, result, *, legacy):
    # Legacy formulas are also used to recognize previously generated sheets.
    weight_col, denominator = ("C", "$I$54") if address == "G4" else ("B", "$F$54")
    cell = f"${address[0]}${address[1:]}"
    amount = (
        f'IF({cell}="",0,IF(AND(ISNUMBER({cell}),{cell}>=0,ROUND({cell},2)={cell}),'
        f'ROUND({cell}*100,0),NA()))'
    )
    quota_range = f"${quota}$14:${quota}$28"
    whole_range = f"${whole}$14:${whole}$28"
    for row in APARTMENT_ROWS:
        weight = f"${weight_col}${row + 22}"
        quotas_through_row = f"${quota}$14:{quota}{row}"
        wholes_through_row = f"${whole}$14:{whole}{row}"
        if legacy:
            sheet[f"{quota}{row}"] = f'=IF({amount}=0,0,IF({denominator}>0,{amount}*{weight}/{denominator},NA()))'
            sheet[f"{whole}{row}"] = f"=INT({quota}{row})"
            remainder = f"ROUND({quota}{row}-{whole}{row},12)"
            all_remainders = f"ROUND({quota_range}-{whole_range},12)"
            remainders_through_row = f"ROUND({quotas_through_row}-{wholes_through_row},12)"
        else:
            weights = f"${weight_col}$36:${weight_col}$50"
            divisor = f"MAX(1,GCD({weights}))"
            total = f"SUMPRODUCT({weights}/{divisor})"
            product = f"({amount}*({weight}/{divisor}))"
            sheet[f"{quota}{row}"] = (
                f"=IF({amount}=0,0,IF(AND({total}>0,{total}<={EXACT_EXCEL_INTEGER},"
                f"{product}<={EXACT_EXCEL_INTEGER}),{product},NA()))"
            )
            sheet[f"{whole}{row}"] = f"=IF({quota}{row}=0,0,QUOTIENT({quota}{row},{total}))"
            remainder = f"({quota}{row}-{whole}{row}*{total})"
            all_remainders = f"({quota_range}-{whole_range}*{total})"
            remainders_through_row = f"({quotas_through_row}-{wholes_through_row}*{total})"
        # Count larger remainders and ties up to this apartment's row.
        rank = (
            f"SUMPRODUCT(--({all_remainders}>{remainder}))"
            f"+SUMPRODUCT(--({remainders_through_row}={remainder}),--({quotas_through_row}>0))"
        )
        sheet[f"{result}{row}"] = (
            f"={whole}{row}+IF({quota}{row}>0,--(({rank})<={amount}-SUM({whole_range})),0)"
        )


def write_expense_report(sheet, result_columns, *, legacy):
    for row in APARTMENT_ROWS:
        def category_total(addresses):
            return "SUM(" + ",".join(f"{result_columns[address]}{row}" for address in addresses) + ")/100"

        active = f"B{row + 22}>0"
        sheet[f"H{row}"] = f'=IF({active},{category_total(TENANT_EXPENSES[1:])},"")'
        sheet[f"I{row}"] = f'=IF({active},{result_columns["G4"]}{row}/100,"")'
        sheet[f"J{row}"] = f'=IF({active},ROUND(SUM(H{row}:I{row}),2),"")'
        sheet[f"K{row}"] = f'=IF({active},{category_total(OWNER_EXPENSES)},"")'
        sheet[f"L{row}"] = f'=IF({active},ROUND(SUM(J{row}:K{row}),2),"")'
        sheet[f"O{row}"] = f'=IF({active},U{row},"")'
        sheet[f"S{row}"] = f"=ROUND(SUM(J{row},M{row}),2)"
        sheet[f"T{row}"] = f"=ROUND(SUM(K{row},N{row}),2)"
        sheet[f"U{row}"] = f"=ROUND(SUM(S{row}:T{row}),2)"
    sheet["A8"] = "Αμοιβή διαχείρισης"
    sheet["R11"] = '=IF(AND(ROUND(SUM(J14:J28)-G12,2)=0,ROUND(SUM(K14:K28)-D12,2)=0),"OK","Ελέγξτε τον επιμερισμό")'
    if not legacy:
        for col, expense_cells in (("D", OWNER_EXPENSES), ("G", TENANT_EXPENSES)):
            sheet[f"{col}12"] = f"=ROUND(SUM({expense_cells[0]}:{expense_cells[-1]}),2)"
        sheet["L12"] = "=ROUND(D12+G12,2)"
        for col in "HIJKL":
            sheet[f"{col}30"] = f"=ROUND(SUM({col}14:{col}28),2)"


@lru_cache(maxsize=2)
def allocation_layout(legacy):
    book = Workbook()
    set_expense_formulas(book.active, legacy=legacy)
    return {cell.coordinate: cell.value for row in book.active for cell in row if cell.value is not None}


def validate_allocation_slots(sheet):
    working = [sheet["R11"], *(sheet[f"{col}{row}"] for row in range(36, 51) for col in "BC")]
    working.extend(sheet.cell(row, col) for row in range(13, 29) for col in range(23, 71))
    for row in sheet.iter_rows(min_col=23, max_col=70, max_row=sheet.max_row):
        for cell in row:
            if (cell.row < 13 or cell.row > 28) and cell.value is not None:
                raise ValueError(tr(f"{sheet.title}!{cell.coordinate}: reserved allocation column is occupied.",
                                    f"{sheet.title}!{cell.coordinate}: η δεσμευμένη στήλη επιμερισμού περιέχει δεδομένα."))
    if all(cell.value is None for cell in working):
        return
    current, previous = allocation_layout(False), allocation_layout(True)
    layout = current if sheet["B36"].value == current["B36"] else previous
    for cell in working:
        if cell.column >= 23 and cell.row in APARTMENT_ROWS:
            start = 23 + (cell.column - 23) // 3 * 3
            quota, whole, result = [sheet.cell(cell.row, start + i).value for i in range(3)]
            if quota is None and whole is None and type(result) is int and result >= 0:
                continue
        if cell.value != layout.get(cell.coordinate):
            raise ValueError(tr(
                f"{sheet.title}!{cell.coordinate}: unexpected content in a reserved allocation cell.",
                f"{sheet.title}!{cell.coordinate}: μη αναμενόμενο περιεχόμενο σε δεσμευμένο κελί επιμερισμού.",
            ))


def monthly_allocations(sheet, standard):
    general, heating = apartment_weights(sheet)
    amounts = {address: money_cents(sheet[address].value, blank_zero=True)
               for address in OWNER_EXPENSES + TENANT_EXPENSES}
    amounts.update(standard)
    if amounts["G4"] > 0 and not any(heating):
        raise ValueError(tr(
            "Heating cannot be allocated. No active apartment has a positive heating weight.",
            "Η θέρμανση δεν μπορεί να επιμεριστεί. Καμία ενεργή ιδιοκτησία δεν έχει θετικό βάρος θέρμανσης.",
        ))
    for address, amount in amounts.items():
        weights = heating if address == "G4" else general
        if amount and max(sum(weights), amount * max(weights)) > EXACT_EXCEL_INTEGER:
            raise ValueError(tr(
                f"{sheet.title}!{address}: expense × allocation weight exceeds Excel's exact integer range.",
                f"{sheet.title}!{address}: το γινόμενο δαπάνης επί βάρος υπερβαίνει την ακρίβεια ακεραίων του Excel.",
            ))
    return {
        address: allocate_cents(amount, heating if address == "G4" else general)
        for address, amount in amounts.items()
    }


def create_month_sheets(wb, wb2, previous_name, *, opening_cash):
    """Prepare both worksheets without saving the workbooks."""
    opening = euros(opening_cash)
    money_cents(opening)
    previous_ledger = wb2[previous_name]
    validate_fee_slot(previous_ledger)
    previous_month = parse_month_sheet(previous_name)
    new_month = previous_month + relativedelta(months=1)
    new_name = month_sheet_name(new_month)
    title = f"{GREEK_MONTHS[new_month.month]} {new_month.year}"

    common = wb.copy_worksheet(wb[previous_name])
    common.title = new_name
    common.conditional_formatting = deepcopy(wb[previous_name].conditional_formatting)
    for description, _, _, _ in ADDITIONAL_BILLS.values():
        common[description] = None
    for address in OWNER_EXPENSES + TENANT_EXPENSES:
        if address != MANAGEMENT_FEE_CELLS[0]:
            common[address] = None

    ledger = wb2.copy_worksheet(previous_ledger)
    ledger.title = new_name
    ledger.conditional_formatting = deepcopy(wb2[previous_name].conditional_formatting)
    for row in range(9, 24):
        ledger[f"D{row}"] = None
        ledger[f"F{row}"] = None
    for row in range(8, 21):
        ledger[f"J{row}"] = None
        ledger[f"L{row}"] = None
    # Older ledgers also use these rows for extra expense descriptions.
    for row in (*range(13, 16), *range(17, 21)):
        ledger[f"K{row}"] = None
    ledger["K16"] = "Αμοιβή διαχείρισης"
    ledger["F8"] = opening
    ledger["F31"] = "=ROUND(SUM(F8:F23),2)"
    ledger["L31"] = "=ROUND(SUM(L8:L20),2)"
    ledger["L34"] = "=ROUND(F31-L31,2)"
    for row in APARTMENT_ROWS:
        common[f"M{row}"] = None
        common[f"N{row}"] = None

    common["O1"] = title
    common["DF3"] = None
    ledger["E3"] = title
    ledger["E8"] = f"{GREEK_MONTHS[previous_month.month]} {previous_month.year}"
    return new_name, new_month, title, common, ledger


def validate_fee_slot(ledger):
    label = ledger["K16"].value
    known = ("", "Εξοδα διαχείρισης", "Έξοδα διαχείρισης", "Αμοιβή διαχείρισης")
    if normalize_name(label).casefold() not in {name.casefold() for name in known} or (
        not label and ledger["L16"].value is not None
    ):
        raise ValueError(tr("Ledger K16 is reserved for the management fee.", "Το K16 του βιβλίου προορίζεται για την αμοιβή διαχείρισης."))


def record_management_fee(common, ledger):
    validate_fee_slot(ledger)
    amount = money_cents(common["D8"].value, blank_zero=True)
    if amount < 0:
        raise ValueError(tr("The management fee cannot be negative.", "Η αμοιβή διαχείρισης δεν μπορεί να είναι αρνητική."))
    ledger["K16"] = "Αμοιβή διαχείρισης"
    ledger["L16"] = euros(amount) if amount else None


def apply_payment(tenant, owner, paid):
    """Allocate cents within a single apartment."""
    if paid < 0:
        raise ValueError(tr("Payments cannot be negative.", "Οι πληρωμές δεν μπορεί να είναι αρνητικές."))
    offset = min(max(-tenant, 0), max(owner, 0)) - min(max(-owner, 0), max(tenant, 0))
    tenant += offset
    owner -= offset
    tenant_cash = min(max(tenant, 0), paid)
    owner_cash = min(max(owner, 0), paid - tenant_cash)
    return tenant - (paid - owner_cash), owner - owner_cash


def settle_balance(tenant, owner):
    paid = max(tenant + owner, 0)
    return (*apply_payment(tenant, owner, paid), paid)


def setup_required(common, ledger):
    return ledger["F8"].value is None or not normalize_name(common["E1"].value) or not normalize_name(ledger["K3"].value)


def is_opening_sheet(common, ledger):
    movement_cells = [f"F{row}" for row in range(9, 24)] + [f"L{row}" for row in range(8, 21)]
    charges = [cell for cell in OWNER_EXPENSES + TENANT_EXPENSES if cell != "D8"]
    return (
        (common["DF3"].value == "OPENING" or common["O1"].value in (None, "ΜΗΝΑΣ ΧΡΟΝΟΣ"))
        and ledger["E3"].value in (None, "ΜΗΝΑΣ ΧΡΟΝΟΣ")
        and all(ledger[cell].value in (None, 0) for cell in movement_cells)
        and all(common[cell].value in (None, 0) for cell in charges)
    )


def setup_values(balance, address, year, month):
    balance = money_cents(balance)
    address = normalize_name(address)
    if balance < 0 or not address:
        raise ValueError(tr(
            "Enter a building address and a non-negative opening balance.",
            "Συμπληρώστε διεύθυνση και μη αρνητικό αρχικό υπόλοιπο.",
        ))
    first_month = datetime(year, month, 1)
    return balance, address, month_sheet_name(first_month - relativedelta(months=1))


def collect_first_time_setup():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(tr('Setup', 'Ρύθμιση'), tr('First-time setup', 'Αρχική ρύθμιση'))

    def open_dialog():
        dialog = tk.Toplevel()
        dialog.title(tr('First-time setup', 'Αρχική ρύθμιση'))

        tk.Label(dialog, text=tr('Opening cash balance:', 'Αρχικό υπόλοιπο ταμείου:')).pack(pady=5)
        entry1 = tk.Entry(dialog)
        entry1.pack(pady=5)

        tk.Label(dialog, text=tr('Street and number:', 'Οδός και αριθμός:')).pack(pady=5)
        entry2 = tk.Entry(dialog)
        entry2.pack(pady=5)

        tk.Label(dialog, text=tr('Select first report month:', 'Επιλέξτε τον πρώτο μήνα έκδοσης:')).pack(pady=5)
        months = DISPLAY_MONTHS
        month_cb = ttk.Combobox(dialog, values=months, state="readonly")
        month_cb.current(0)
        month_cb.pack(pady=5)

        tk.Label(dialog, text=tr('Year:', 'Έτος:')).pack(pady=5)
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 5, current_year + 6)]
        year_cb = ttk.Combobox(dialog, values=years, state="readonly")
        year_cb.set(str(current_year))
        year_cb.pack(pady=5)

        result = {}

        def submit():
            try:
                result["setup"] = setup_values(
                    entry1.get(), entry2.get(), int(year_cb.get()), month_cb.current() + 1,
                )
                dialog.destroy()

            except Exception as e:
                messagebox.showerror(tr('Error', 'Σφάλμα'), tr(f'Invalid input: {e}', f'Μη έγκυρη εισαγωγή: {e}'))

        tk.Button(dialog, text=tr('Submit', 'Καταχώριση'), command=submit).pack(pady=10)

        dialog.grab_set()
        dialog.wait_window()

        return result.get("setup")

    try:
        return open_dialog()
    finally:
        root.destroy()


def month_pdf_path(path, month):
    return Path(path).parent / f"ΚΟΙΝΟΧΡΗΣΤΑ {GREEK_MONTHS[month.month]} {month.year}.pdf"


def deletion_target(wb, wb2):
    """Only delete a matching last month, never the sole remaining template."""
    if len(wb.sheetnames) < 2 or len(wb2.sheetnames) < 2:
        raise ValueError(tr('The only remaining sheet cannot be deleted.', 'Δεν επιτρέπεται η διαγραφή του μοναδικού φύλλου.'))
    if wb.sheetnames != wb2.sheetnames:
        raise ValueError(tr('The workbook sheet histories differ. Deletion canceled.', 'Τα ιστορικά φύλλα των δύο Excel δεν συμφωνούν. Η διαγραφή ακυρώθηκε.'))
    name = wb.sheetnames[-1]
    month = parse_month_sheet(name)
    previous = parse_month_sheet(wb.sheetnames[-2])
    if previous + relativedelta(months=1) != month:
        raise ValueError(tr('The last two sheets are not consecutive months.', 'Τα δύο τελευταία φύλλα δεν είναι διαδοχικοί μήνες.'))
    validate_template_layout(wb, wb2)
    return name, month


def delete_month_files(wb, path, wb2, path2, expected_name, *, revisions=None):
    name, month = deletion_target(wb, wb2)
    if name != expected_name:
        raise ValueError(tr('The last month changed. Try again.', 'Ο τελευταίος μήνας άλλαξε. Δοκιμάστε ξανά.'))
    pdf = month_pdf_path(path, month)
    del wb[name]
    del wb2[name]
    validate_template_layout(wb, wb2)
    warnings = save_workbooks_safely(wb, path, wb2, path2, revisions=revisions, delete_pdf=pdf)
    if warnings:
        return tr("The month was deleted. Cleanup needs attention:\n",
                  "Ο μήνας διαγράφηκε. Ο καθαρισμός χρειάζεται έλεγχο:\n") + "\n".join(warnings)
    return None


def launch_delete_last_month(path, path2):
    try:
        with workbook_pair(path, path2) as (wb, wb2, revisions):
            name, month = deletion_target(wb, wb2)
            pdf = month_pdf_path(path, month)
            detail = f"PDF: {pdf.name}" if pdf.exists() else tr('No matching PDF was found.', 'Δεν βρέθηκε αντίστοιχο PDF.')
            try:
                previous_finances(path, path2, wb, wb2)
            except Exception as exc:
                detail += "\n\n" + tr(
                    "The issued month could not be reconciled. Confirm deletion only to correct or recreate it.\n",
                    "Ο εκδομένος μήνας δεν μπόρεσε να συμφωνηθεί. Επιβεβαιώστε τη διαγραφή μόνο για διόρθωση ή επανέκδοση.\n",
                ) + error_details(exc)
            if not messagebox.askyesno(
                tr('Confirm deletion', 'Επιβεβαίωση διαγραφής'),
                tr(f"Are you sure you want to delete the sheets '{name}' ({DISPLAY_MONTHS[month.month - 1]} {month.year}) from both Excel workbooks?\n\n{detail}\n\nDeletion is permanent.", f"Είστε σίγουρος ότι θέλετε να διαγράψετε τις καρτέλες '{name}' ({DISPLAY_MONTHS[month.month - 1]} {month.year}) και από τα δύο Excel;\n\n{detail}\n\nΗ διαγραφή είναι οριστική."), default="no"
            ):
                return
            warning = delete_month_files(wb, path, wb2, path2, name, revisions=revisions)
            if warning:
                messagebox.showwarning(tr('Deletion', 'Διαγραφή'), warning)
            else:
                messagebox.showinfo(tr('Deletion', 'Διαγραφή'), tr(f'Month {name} was deleted from both Excel workbooks.\n', f'Ο μήνας {name} διαγράφηκε και από τα δύο Excel.\n') +
                                    (tr('The matching PDF no longer exists.', 'Το αντίστοιχο PDF δεν υπάρχει πλέον.')))
    except Exception as exc:
        messagebox.showerror(tr('Deletion error', 'Σφάλμα διαγραφής'), error_details(exc))


def launch_welcome_dashboard():
    root = tk.Tk()
    root.title(tr('Building Expense Manager', 'Διαχείριση κοινοχρήστων'))
    root.geometry("560x480")

    title = tk.Label(
        root,
        text=tr('Building Expense Manager', 'Διαχείριση κοινοχρήστων'),
        font=("Helvetica", 20, "bold"),
    )
    title.pack(pady=40)

    subtitle = tk.Label(
        root,
        text=tr('Monthly expenses and payments', 'Μηνιαίες δαπάνες και πληρωμές'),
        font=("Helvetica", 12),
    )
    subtitle.pack(pady=10)

    choice = {"action": None}

    def start():
        choice["action"] = "start"
        root.destroy()

    def edit():
        choice["action"] = "edit"
        root.destroy()

    def delete():
        choice["action"] = "delete"
        root.destroy()

    def change_language():
        choice["action"] = "language"
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=12)

    tk.Button(btn_frame, text=tr('Start next month', 'Έναρξη νέου μήνα'), width=20, command=start).pack(pady=10)
    tk.Button(btn_frame, text=tr('Edit apartment details', 'Επεξεργασία στοιχείων'), width=24, command=edit).pack(pady=10)
    tk.Button(btn_frame, text=tr('Delete last month', 'Διαγραφή τελευταίου μήνα'), width=28, command=delete).pack(pady=10)

    ttk.Button(btn_frame, text="Language / Γλώσσα", command=change_language, width=28).pack(pady=10)

    def choose_files():
        choice["action"] = "files"
        root.destroy()

    ttk.Button(btn_frame, text=tr("Change workbooks", "Αλλαγή αρχείων Excel"), command=choose_files).pack(pady=6)
    root.mainloop()
    return choice["action"]


def apartment_cells(row: int, *, pending: bool) -> list[str]:
    if pending:
        return [f"{col}{row}" for col in PENDING_COLUMNS]
    return [f"{col}{row}" for col in "BCDE"] + [f"H{row + 22}"]


def property_values(sheet, row):
    pending = sheet["DF1"].value == "NEXT_MONTH"
    return [sheet[address].value for address in apartment_cells(row, pending=pending)]


def management_fee_value(sheet):
    return sheet["DF2"].value if sheet["DF1"].value == "NEXT_MONTH" else sheet["D8"].value


def write_properties(common, ledger, properties, fee, *, pending=False):
    if pending:
        common["DF1"] = "NEXT_MONTH"
        common["DF2"] = euros(fee)
        hide_working_columns(common, column_index_from_string("DA"), column_index_from_string("DF"))
    else:
        common["D8"] = euros(fee)
    for row, (floor, area, tenant, owner, factor) in zip(APARTMENT_ROWS, properties):
        if not pending:
            literal_text(ledger, f"E{row - 5}", owner)
        floor_cell, area_cell, tenant_cell, owner_cell, factor_cell = apartment_cells(row, pending=pending)
        literal_text(common, floor_cell, floor)
        common[area_cell] = area
        literal_text(common, tenant_cell, tenant)
        literal_text(common, owner_cell, owner)
        common[factor_cell] = factor


def apply_pending_properties(common, ledger, balances):
    properties = [property_values(common, row) for row in APARTMENT_ROWS]
    for row, details, (tenant, owner, _) in zip(APARTMENT_ROWS, properties, balances):
        try:
            validate_balance_transition(common[f"E{row}"].value, details[3], tenant, owner)
        except ValueError as exc:
            raise ValueError(f"{common.title} / #{row - 13}: {exc}") from exc
    if common["DF1"].value == "NEXT_MONTH":
        write_properties(common, ledger, properties, money_cents(management_fee_value(common)))
        for row in APARTMENT_ROWS:
            for address in apartment_cells(row, pending=True):
                common[address] = None
        common["DF1"] = None
        common["DF2"] = None


def launch_edit_dashboard(ws, save_properties, *, initial=False):
    root = tk.Tk()
    root.title(tr('Edit apartment details', 'Επεξεργασία στοιχείων'))
    root.geometry(tr('820x650', '920x690'))

    entries = []
    result = {"status": "quit"}

    frame = tk.Frame(root)
    frame.pack(pady=10)

    headers = [tr('No.', 'Α/Α'), tr('Floor', 'Όροφος'), tr('Sqm', 'Τ.μ.'), tr('Tenant', 'Ενοικιαστής'), tr('Owner', 'Ιδιοκτήτης'), tr('Heating factor (0–1)', 'Θέρμανση (0–1)')]
    for col, h in enumerate(headers):
        tk.Label(frame, text=h, font=("Arial", 10, "bold")).grid(row=0, column=col)

    floor_options = ["ΥΠ2", "ΥΠ1", "ΙΣ", "Α' ΟΡ", "Β' ΟΡ", "Γ' ΟΡ"]

    for i in range(15):
        row_entries = []

        tk.Label(frame, text=str(i + 1), width=3).grid(row=i + 1, column=0)

        f, c, d, e, h = ["" if value is None else value for value in property_values(ws, 14 + i)]

        floor_var = tk.StringVar(value=FLOOR_LABELS.get(str(f), str(f)))
        floor_cb = ttk.Combobox(frame, textvariable=floor_var, values=[FLOOR_LABELS[x] for x in floor_options], state="readonly", width=13)
        floor_cb.grid(row=i + 1, column=1)
        row_entries.append(floor_var)

        for j, val in enumerate([c, d, e, h]):
            var = tk.StringVar(value=str(val))
            tk.Entry(frame, textvariable=var, width=15).grid(row=i + 1, column=j + 2)
            row_entries.append(var)

        entries.append(row_entries)

    fee = management_fee_value(ws)
    fee_var = tk.StringVar(value=str(fee if fee is not None else ""))
    if not initial:
        tk.Label(root, text=tr(
            "Changes apply to the next month.",
            "Οι αλλαγές ισχύουν από τον επόμενο μήνα.",
        )).pack()

    tk.Label(root, text=tr('Heating weight: 0 exempt, 0.25 quarter weight, 1 or blank full weight', 'Βάρος θέρμανσης: 0 απαλλαγή, 0,25 ένα τέταρτο, 1 ή κενό πλήρες βάρος'),
             wraplength=760).pack(pady=5)
    tk.Label(root, text=tr('Management Fee', 'Αμοιβή διαχείρισης')).pack(pady=10)
    tk.Entry(root, textvariable=fee_var).pack()

    def finish():
        if not messagebox.askyesno(tr('Confirm', 'Επιβεβαίωση'), tr('Save changes?', 'Αποθήκευση αλλαγών;')):
            return
        try:
            properties = []
            for i, row in enumerate(entries, start=1):
                try:
                    area, owner, factor = apartment_details(row[1].get(), row[3].get(), row[4].get())
                except ValueError as exc:
                    raise ValueError(f"#{i}: {exc}") from exc
                floor = next((key for key, label in FLOOR_LABELS.items() if label == row[0].get()), row[0].get())
                properties.append((floor, area, normalize_name(row[2].get()), owner, factor))
            if initial and not any(area for _, area, *_ in properties):
                raise ValueError(tr("Set up at least one active apartment.", "Ρυθμίστε τουλάχιστον μία ενεργή ιδιοκτησία."))
            fee = money_cents(fee_var.get(), blank_zero=True)
            if fee < 0:
                raise ValueError(tr("The management fee cannot be negative.", "Η αμοιβή διαχείρισης δεν μπορεί να είναι αρνητική."))
            if fee and not any(area for _, area, *_ in properties):
                raise ValueError(tr("The management fee needs an active apartment.", "Η αμοιβή διαχείρισης απαιτεί ενεργή ιδιοκτησία."))
            if save_properties(properties, fee) is False:
                return
        except Exception as exc:
            messagebox.showerror(tr('Check apartment details', 'Έλεγχος στοιχείων ιδιοκτησίας'), error_details(exc))
            return
        result["status"] = "next"
        root.destroy()

    def on_close():
        result["status"] = "quit"
        if root.winfo_exists():
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Button(root, text=tr('Finish', 'Ολοκλήρωση'), command=finish).pack(pady=20)

    root.mainloop()
    return result["status"]


def launch_dashboard(owners, tenant_debts, owner_debts, tenants=None):
    tenants = tenants if tenants is not None else [None] * len(owners)
    result = {"status": "quit"}

    root = tk.Tk()
    root.title(tr('Common expense payments', 'Πληρωμές κοινοχρήστων'))
    root.geometry("920x580")

    payments = [0] * len(owners)
    full_paid_flags = [False] * len(owners)

    payment_vars = []
    full_vars = []

    scroll_frame = tk.Frame(root)
    scroll_frame.pack(fill="both", expand=True)

    bottom_frame = tk.Frame(root)
    bottom_frame.pack(side="bottom", fill="x", pady=10)

    headers = [tr('Owner', 'Ιδιοκτήτης'), tr('Tenant', 'Ενοικιαστής'), tr('Net balance (€)', 'Καθαρό υπόλοιπο (€)'), tr('Payment', 'Πληρωμή'), tr('Full settlement', 'Εξόφληση')]
    for col, text in enumerate(headers):
        tk.Label(scroll_frame, text=text, font=("Arial", 10, "bold")).grid(row=0, column=col, padx=5, pady=5)

    for i, owner in enumerate(owners):
        total_debt = tenant_debts[i] + owner_debts[i]

        tk.Label(scroll_frame, text=owner, wraplength=235, justify="left").grid(row=i+1, column=0, sticky="w", padx=5)
        tk.Label(scroll_frame, text=tenants[i] or "—", wraplength=235, justify="left").grid(row=i+1, column=1, sticky="w", padx=5)
        tk.Label(scroll_frame, text=f"{euros(total_debt):.2f}").grid(row=i+1, column=2)

        pay_var = tk.StringVar()
        tk.Entry(scroll_frame, textvariable=pay_var, width=10).grid(row=i+1, column=3)
        payment_vars.append(pay_var)

        full_var = tk.BooleanVar()
        tk.Checkbutton(scroll_frame, variable=full_var).grid(row=i+1, column=4)
        full_vars.append(full_var)

    def submit():
        try:
            for i in range(len(owners)):
                val = payment_vars[i].get().strip()
                full_flag = full_vars[i].get()

                if val and full_flag:
                    messagebox.showerror(
                        tr('Error', 'Σφάλμα'),
                        tr(f'{owners[i]}: Enter a payment amount or choose full settlement.', f'{owners[i]}: Βάλτε ποσό πληρωμής ή επιλέξτε εξόφληση.')
                    )
                    return

                payment = money_cents(val, blank_zero=True)

                if payment < 0:
                    messagebox.showerror(
                        tr('Error', 'Σφάλμα'),
                        tr(f'{owners[i]}: Negative payments are not allowed.', f'{owners[i]}: Δεν επιτρέπονται αρνητικές πληρωμές.')
                    )
                    return

                payments[i] = payment
                full_paid_flags[i] = full_flag

            result["status"] = "next"
            root.destroy()
        except ValueError as exc:
            messagebox.showerror(tr('Error', 'Σφάλμα'), error_details(exc))

    def on_close():
        result["status"] = "quit"
        if root.winfo_exists():
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Button(bottom_frame, text=tr('Next Step', 'Επόμενο βήμα'), command=submit).pack(pady=5)

    root.mainloop()

    if result["status"] == "quit":
        return None, None, "quit"

    return payments, full_paid_flags, "next"


def launch_bills_dashboard(validate_expenses):
    root = tk.Tk()
    root.title(tr('Enter expenses', 'Καταχώριση δαπανών'))
    root.geometry("530x780")

    entries = {}

    tk.Label(root, text=tr('Enter expenses', 'Καταχώριση δαπανών'), font=("Arial", 12, "bold")).pack(pady=10)

    frame = tk.Frame(root)
    frame.pack()

    for name in STANDARD_BILLS:
        row = tk.Frame(frame)
        row.pack(fill="x", pady=2)

        tk.Label(row, text=BILL_LABELS[name], width=34, anchor="w").pack(side="left")

        var = tk.StringVar()
        tk.Entry(row, textvariable=var, width=10).pack(side="right")

        entries[name] = var

    tk.Label(root, text=tr('Owners pay urgent expenses.\nTenants pay extra expenses.', 'Οι έκτακτες δαπάνες χρεώνονται στους ιδιοκτήτες.\nΟι πρόσθετες δαπάνες χρεώνονται στους ενοικιαστές.'), wraplength=480, justify="left").pack(pady=8)

    additional_titles = {
        "urgent": tr("Owner expense 1", "Έκτακτη δαπάνη ιδιοκτητών 1"),
        "urgent2": tr("Owner expense 2", "Έκτακτη δαπάνη ιδιοκτητών 2"),
        "extra1": tr("Extra tenant expense 1", "Πρόσθετη δαπάνη ενοικιαστών 1"),
        "extra2": tr("Extra tenant expense 2", "Πρόσθετη δαπάνη ενοικιαστών 2"),
    }
    additional_entries = {}
    for key in ADDITIONAL_BILLS:
        description, amount = tk.StringVar(), tk.StringVar()
        tk.Label(root, text=additional_titles[key]).pack(pady=5)
        tk.Entry(root, textvariable=description).pack()
        tk.Entry(root, textvariable=amount).pack()
        additional_entries[key] = description, amount

    result = {}

    def collect_bills_data(generate_pdf=False):
        try:
            data = {}
            data["standard"] = {}

            for k, v in entries.items():
                val = money_cents(v.get(), blank_zero=True)

                if val < 0:
                    messagebox.showerror(tr('Error', 'Σφάλμα'), tr(f'{BILL_LABELS[k]}: Negative values not allowed', f'{BILL_LABELS[k]}: Δεν επιτρέπονται αρνητικές τιμές'))
                    return None

                data["standard"][k] = val

            for key, (description, amount) in additional_entries.items():
                data[key] = description.get(), money_cents(amount.get(), blank_zero=True)

            for key in ADDITIONAL_BILLS:
                description, amount = data[key]
                if amount < 0:
                    raise ValueError(tr(
                        f"{EXPENSE_LABELS[key]}: negative values not allowed",
                        f"{EXPENSE_LABELS[key]}: δεν επιτρέπονται αρνητικές τιμές",
                    ))
                if amount > 0 and not description.strip():
                    raise ValueError(tr(
                        f"{EXPENSE_LABELS[key]}: description is required",
                        f"{EXPENSE_LABELS[key]}: απαιτείται περιγραφή",
                    ))

            data["allocations"] = validate_expenses(data)
            data["generate_pdf"] = generate_pdf

            return data

        except Exception as e:
            messagebox.showerror(tr('Error', 'Σφάλμα'), f"{e}")
            return None

    def submit(generate_pdf=False):
        data = collect_bills_data(generate_pdf)
        if data is not None:
            result.update(data)
            root.destroy()

    def on_close():
        result["status"] = "quit"
        if root.winfo_exists():
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Button(root, text=tr('Finish without PDF', 'Ολοκλήρωση χωρίς PDF'), command=submit).pack(pady=5)
    tk.Button(root, text=tr('Finish and export PDF', 'Ολοκλήρωση με PDF'), command=lambda: submit(generate_pdf=True)).pack(pady=10)

    root.mainloop()

    if result.get("status") == "quit":
        return None, "quit"

    return result, "next"


def valid_config(data):
    if not isinstance(data, dict):
        return {}
    result = {}
    if data.get("language") in ("en", "el"):
        result["language"] = data["language"]
    for key in ("file1", "file2"):
        path = data.get(key)
        if isinstance(path, str) and "\0" not in path and Path(path).is_absolute() and Path(path).suffix.lower() == ".xlsx":
            result[key] = path
    return result


def load_paths():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as source:
            return valid_config(json.load(source))
    except (OSError, ValueError):
        return {}


def save_paths(data):
    settings = load_paths()
    settings.update(data)
    directory = Path(CONFIG_FILE).parent
    directory.mkdir(parents=True, exist_ok=True)
    fd, filename = tempfile.mkstemp(prefix=".bem-config-", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as target:
            json.dump(valid_config(settings), target, ensure_ascii=False, indent=2)
            target.flush()
            os.fsync(target.fileno())
        os.replace(filename, CONFIG_FILE)
    finally:
        Path(filename).unlink(missing_ok=True)


def protect_template(path):
    for original in (Path(__file__).resolve().parent / "templates").glob("*.xlsx"):
        if path.resolve() == original.resolve() or (path.exists() and os.path.samefile(path, original)):
            raise ValueError(tr(
                "Make private working copies of both public templates and select the copies.",
                "Δημιουργήστε ιδιωτικά αντίγραφα των δύο δημόσιων προτύπων και επιλέξτε τα αντίγραφα.",
            ))


def check_workbook_paths(path, path2):
    if path and path2:
        check_recovery(tuple(Path(item).resolve() for item in (path, path2)))
    for item in (path, path2):
        candidate = Path(item or "")
        if not candidate.is_absolute() or candidate.suffix.lower() != ".xlsx" or not candidate.is_file():
            raise ValueError(tr(f"Select a valid Excel file: {item}", f"Επιλέξτε έγκυρο αρχείο Excel: {item}"))
        protect_template(candidate)
    if os.path.samefile(path, path2):
        raise ValueError(tr("Select two different workbooks.", "Επιλέξτε δύο διαφορετικά βιβλία."))


def select_workbook_pair():
    while True:
        path = select_file(tr("Select the common expenses workbook (ΚΟΙΝΟΧΡΗΣΤΑ)", "Επιλέξτε το αρχείο ΚΟΙΝΟΧΡΗΣΤΑ"), "ΚΟΙΝΟΧΡΗΣΤΑ")
        if not path:
            return None
        path2 = select_file(tr("Select the accounting ledger (ΒΙΒΛΙΟ)", "Επιλέξτε το αρχείο ΒΙΒΛΙΟ"), "ΒΙΒΛΙΟ")
        if not path2:
            return None
        try:
            check_workbook_paths(path, path2)
            with workbook_pair(path, path2):
                pass
        except Exception as exc:
            messagebox.showerror(tr("Workbook selection", "Επιλογή βιβλίων"), error_details(exc))
            continue
        try:
            save_paths({"file1": path, "file2": path2})
        except OSError as exc:
            messagebox.showwarning(tr("Settings", "Ρυθμίσεις"), tr(
                f"Could not remember the workbook paths: {exc}",
                f"Δεν αποθηκεύτηκαν οι διαδρομές των βιβλίων: {exc}",
            ))
        return path, path2


def select_file(title, expected_keyword):
    root = tk.Tk()
    root.withdraw()

    while True:
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[(tr('Excel files', 'Αρχεία Excel'), "*.xlsx")]
        )

        if not file_path:
            root.destroy()
            return None

        filename = os.path.basename(file_path)

        if expected_keyword.casefold() not in filename.casefold():
            messagebox.showerror(
                tr('Wrong File', 'Λανθασμένο αρχείο'),
                tr(f'Choose an Excel file with {expected_keyword} in its name.', f'Επιλέξτε αρχείο Excel που έχει το {expected_keyword} στο όνομά του.')
            )
            continue

        root.destroy()
        return file_path


def choose_language():
    """Return a language and preference, or None if the window is closed."""
    root = tk.Tk()
    root.title("Language / Γλώσσα")
    root.geometry("440x290")
    root.resizable(False, False)
    selected = {"choice": None}
    remember = tk.BooleanVar(master=root, value=False)
    tk.Label(root, text="Building Expense Manager", font=("Helvetica", 18, "bold")).pack(pady=(25, 8))
    tk.Label(root, text="Select your language / Επιλέξτε γλώσσα").pack(pady=8)

    def select(language):
        selected["choice"] = (language, remember.get())
        root.destroy()

    ttk.Button(root, text="Ελληνικά", command=lambda: select("el"), width=25).pack(pady=6)
    ttk.Button(root, text="English", command=lambda: select("en"), width=25).pack(pady=6)
    ttk.Checkbutton(root, text="Remember my choice / Αποθήκευση επιλογής", variable=remember).pack(pady=12)
    root.mainloop()
    return selected["choice"]


def configure_language(*, force=False):
    saved = load_paths().get("language")
    if not force and saved in ("en", "el"):
        set_language(saved)
        return True
    choice = choose_language()
    if choice is None:
        return False
    language, remember = choice
    set_language(language)
    try:
        save_paths({"language": language if remember else None})
    except OSError as exc:
        messagebox.showwarning(tr("Settings", "Ρυθμίσεις"), tr(f"Could not save language preference: {exc}", f"Δεν αποθηκεύτηκε η επιλογή γλώσσας: {exc}"))
    return True


def edit_properties(path, path2, wb, wb2, revisions, *, setup=None):
    common, ledger = wb.worksheets[-1], wb2.worksheets[-1]
    has_history = not is_unissued_pair(wb, wb2)
    balances = previous_finances(path, path2, wb, wb2)[0] if has_history else opening_balances(common)

    def save(properties, fee):
        tenant_changes = []
        for row, details, (tenant, owner, _) in zip(APARTMENT_ROWS, properties, balances):
            try:
                validate_balance_transition(common[f"E{row}"].value, details[3], tenant, owner)
            except ValueError as exc:
                raise ValueError(f"#{row - 13}: {exc}") from exc
            if tenant and normalize_name(common[f"D{row}"].value) != details[2]:
                tenant_changes.append((row, tenant))
        if not confirm_tenant_changes(tenant_changes):
            return False
        if setup is not None:
            cash, address, base_month = setup
            common.title = ledger.title = base_month
            ledger["F8"] = euros(cash)
            literal_text(ledger, "K3", address)
            literal_text(common, "E1", address)
            common["DF3"] = "OPENING"
        if not has_history:
            write_properties(common, ledger, properties, 0)
            set_expense_formulas(common)
        write_properties(common, ledger, properties, fee, pending=True)
        warnings = save_workbooks_safely(wb, path, wb2, path2, revisions=revisions)
        if warnings:
            messagebox.showwarning(tr("Settings saved", "Οι ρυθμίσεις αποθηκεύτηκαν"), "\n".join(warnings))
        return True

    return launch_edit_dashboard(common, save, initial=setup is not None)


def bill_amounts(data):
    amounts = {STANDARD_BILLS[key][0]: value for key, value in data["standard"].items()}
    amounts.update({charge: data[key][1] for key, (_, charge, _, _) in ADDITIONAL_BILLS.items()})
    return amounts


def record_expenses(common, ledger, data):
    for key, amount in data["standard"].items():
        charge, cash = STANDARD_BILLS[key]
        common[charge] = ledger[cash] = euros(amount) if amount else None
    for key, (desc_cell, charge, ledger_desc, cash) in ADDITIONAL_BILLS.items():
        description, amount = data[key]
        literal_text(common, desc_cell, description.strip() if amount else "")
        literal_text(ledger, ledger_desc, description.strip() if amount else "")
        common[charge] = ledger[cash] = euros(amount) if amount else None
    record_management_fee(common, ledger)
    set_expense_formulas(common, data["allocations"])


def record_payment(common, ledger, row, tenant, owner, actual_payment, *, full=False):
    if full:
        tenant, owner, actual_payment = settle_balance(tenant, owner)
    else:
        tenant, owner = apply_payment(tenant, owner, actual_payment)
    common[f"M{row}"] = euros(tenant)
    common[f"N{row}"] = euros(owner)
    ledger[f"F{row - 5}"] = euros(actual_payment) if actual_payment else None


def validate_month_cash(ledger):
    opening = money_cents(ledger["F8"].value)
    income = sum(money_cents(ledger[f"F{row}"].value, blank_zero=True) for row in range(9, 24))
    expenses = sum(money_cents(ledger[f"L{row}"].value, blank_zero=True) for row in range(8, 21))
    for amount in (opening + income, expenses, opening + income - expenses):
        money_cents(euros(amount))


def export_month_pdf(path, month, sheet_name):
    destination = month_pdf_path(path, month)
    if os.path.lexists(destination):
        raise FileExistsError(tr(f"An existing PDF was preserved: {destination}",
                                 f"Το ήδη υπάρχον PDF διατηρήθηκε: {destination}"))
    temporary = stage_path(destination, ".bem-pdf-")
    try:
        with excel_app() as app:
            book = app.books.open(os.path.abspath(path), update_links=False, read_only=True)
            try:
                sheet = book.sheets[sheet_name]
                app.calculate()
                sheet.page_setup.print_area = "$A$1:$P$31"
                sheet.to_pdf(path=str(temporary), show=False)
            finally:
                book.close()
        with temporary.open("rb") as output:
            if output.read(5) != b"%PDF-":
                raise ValueError(tr("Excel produced an incomplete PDF.", "Το Excel παρήγαγε ελλιπές PDF."))
        os.replace(temporary, destination)
    except BaseException as exc:
        exc.add_note(f"{temporary} → {destination}")
        raise
    return destination


def run_month(path, path2, wb, wb2, revisions):
    previous = validate_workbook_pair(wb, wb2)
    old_layout = wb[previous]["B36"].value != allocation_layout(False)["B36"]
    unissued = is_unissued_pair(wb, wb2)
    if not unissued and old_layout and not messagebox.askyesno(
        tr("Check legacy cash", "Έλεγχος παλαιού ταμείου"),
        tr("Confirm that the legacy closing balance matches actual cash before continuing.",
           "Επιβεβαιώστε ότι το κλείσιμο του παλαιού βιβλίου συμφωνεί με το πραγματικό ταμείο πριν συνεχίσετε."),
        default="no",
    ):
        return False
    balances, opening_cash = previous_finances(path, path2, wb, wb2)
    if not confirm_legacy_rounding(balances):
        return False
    source = wb[previous]
    if old_layout and source["DF1"].value == "NEXT_MONTH" and not confirm_tenant_changes([
        (row, tenant) for row, (tenant, _, _) in zip(APARTMENT_ROWS, balances)
        if tenant and normalize_name(source[f"D{row}"].value) != normalize_name(property_values(source, row)[2])
    ]):
        return False
    name, month, _, common, ledger = create_month_sheets(
        wb, wb2, previous, opening_cash=opening_cash,
    )
    apply_pending_properties(common, ledger, balances)
    weights, _ = apartment_weights(common)
    rows = [row for row, weight in zip(APARTMENT_ROWS, weights) if weight]
    if unissued and not rows:
        raise ValueError(tr("Set up an active apartment before issuing the first month.",
                            "Ρυθμίστε ενεργή ιδιοκτησία πριν από την έκδοση του πρώτου μήνα."))
    for row, weight, (tenant, owner, _) in zip(APARTMENT_ROWS, weights, balances):
        if not weight and (tenant or owner):
            raise ValueError(tr(
                f"{previous} / #{row - 13}: resolve the inactive apartment's balances before continuing.",
                f"{previous} / #{row - 13}: τακτοποιήστε τα υπόλοιπα της ανενεργής ιδιοκτησίας πριν συνεχίσετε.",
            ))
    messagebox.showinfo(tr("New month", "Νέος μήνας"), f"{DISPLAY_MONTHS[month.month - 1]} {month.year}")
    if rows and not collect_month_payments(common, ledger, rows, balances):
        return False

    bills, status = launch_bills_dashboard(lambda data: monthly_allocations(common, bill_amounts(data)))
    if status == "quit":
        return False
    record_expenses(common, ledger, bills)
    validate_month_cash(ledger)
    if not messagebox.askyesno(tr("Confirm", "Επιβεβαίωση"), tr(
        "Save changes to both Excel workbooks?", "Να αποθηκευτούν οι αλλαγές και στα δύο βιβλία Excel;",
    )):
        return False
    save_month(path, path2, wb, wb2, revisions, name=name, month=month, bills=bills)
    return True


def collect_month_payments(common, ledger, rows, balances) -> bool:
    owners = [f"#{row - 13} · {common[f'E{row}'].value}" for row in rows]
    tenants = [common[f"D{row}"].value for row in rows]
    debts = [balances[row - APARTMENT_ROWS.start] for row in rows]
    payments, full_settlements, status = launch_dashboard(
        owners, [tenant for tenant, _, _ in debts], [owner for _, owner, _ in debts], tenants,
    )
    if status == "quit":
        return False
    for row, (tenant, owner, _), payment, full in zip(rows, debts, payments, full_settlements):
        record_payment(common, ledger, row, tenant, owner, payment, full=full)
    return True


def save_month(path, path2, wb, wb2, revisions, *, name, month, bills):
    export = (lambda: export_month_pdf(path, month, name)) if bills["generate_pdf"] else None

    def check_prepared(staged, cash):
        try:
            reconcile_issued_period(staged, cash, name, expected_allocations=bills["allocations"])
        except ValueError as exc:
            message = str(exc).replace(staged.name, Path(path).name).replace(cash.name, Path(path2).name)
            raise ValueError(message) from exc

    warnings = save_workbooks_safely(
        wb, path, wb2, path2, revisions=revisions,
        before_commit=check_prepared,
        after_commit=export,
        pdf_destination=month_pdf_path(path, month) if export else None,
    )
    if warnings:
        messagebox.showwarning(tr("Financial records saved", "Οι οικονομικές εγγραφές αποθηκεύτηκαν"), "\n".join(warnings))
    elif export:
        pdf = month_pdf_path(path, month)
        messagebox.showinfo(tr("New file", "Νέο αρχείο"), tr(f"PDF generated:\n{pdf}", f"Δημιουργήθηκε PDF:\n{pdf}"))
    else:
        messagebox.showinfo(tr("Saved", "Αποθηκεύτηκε"), tr("Both workbooks were saved.", "Τα δύο βιβλία αποθηκεύτηκαν."))


def main():
    if not configure_language():
        return
    settings = load_paths()
    paths = settings.get("file1"), settings.get("file2")
    try:
        check_workbook_paths(*paths)
        with workbook_pair(*paths):
            pass
    except Exception as exc:
        if all(paths):
            messagebox.showwarning(tr("Workbooks", "Βιβλία"), error_details(exc))
        paths = select_workbook_pair()
    while paths:
        action = launch_welcome_dashboard()
        if action is None:
            return
        if action == "language":
            configure_language(force=True)
            continue
        if action == "files":
            paths = select_workbook_pair() or paths
            continue
        path, path2 = paths
        try:
            if action == "delete":
                launch_delete_last_month(path, path2)
                continue
            with workbook_pair(path, path2) as (wb, wb2, revisions):
                common, ledger = wb.worksheets[-1], wb2.worksheets[-1]
                if setup_required(common, ledger):
                    if len(wb.sheetnames) != 1 or not is_opening_sheet(common, ledger):
                        raise ValueError(tr(
                            "Existing accounts need a valid address and opening balance. Review both workbooks.",
                            "Το υπάρχον ιστορικό απαιτεί έγκυρη διεύθυνση και αρχικό υπόλοιπο. Ελέγξτε τα δύο βιβλία.",
                        ))
                    setup = collect_first_time_setup()
                    if setup is not None:
                        edit_properties(path, path2, wb, wb2, revisions, setup=setup)
                elif action == "edit":
                    edit_properties(path, path2, wb, wb2, revisions)
                elif run_month(path, path2, wb, wb2, revisions):
                    return
        except (KeyboardInterrupt, Exception) as exc:
            messagebox.showerror(tr("Error", "Σφάλμα"), error_details(exc) or tr("Process interrupted", "Η διαδικασία διακόπηκε"))


if __name__ == "__main__":
    main()
