#!/usr/bin/env python
# coding: utf-8

import os
import sys
from datetime import datetime


from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook

import tkinter as tk
from tkinter import messagebox, ttk, filedialog

import json
import math
import re
import tempfile
import shutil
from pathlib import Path
from contextlib import contextmanager
from copy import deepcopy


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


# One shared implementation; language affects presentation only.
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
    DISPLAY_MONTHS = tr(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], list(GREEK_MONTHS.values()))


set_language("en")

# Cells in the common-expenses workbook that must be cleared for a new month.
BILL_CELLS_TO_CLEAR = [
    "A6", "D6", "A7", "D7",
    "G4", "G5", "G6", "G7", "G8",
    "E9", "G9",
    "E10", "G10"
]

# Cells in the accounting workbook that must be cleared for a new month.
CELLS_TO_CLEAR_2 = [
    "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23",
    "L8", "L9", "L10", "L11", "L12", "L13", "L14", "L15", "L17", "L18", "L19", "L20",
    "K13", "K14", "K15", "K17", "K18", "K19", "K20"
]


def get_app_data_path():
    """Per-user writable settings; works for source and frozen executables."""
    if sys.platform == "win32":
        folder = os.environ.get("APPDATA") or os.path.join(Path.home(), "AppData", "Roaming")
        base = Path(folder) / "BuildingExpenseManager"
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / "BuildingExpenseManager"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "building-expense-manager"
    return str(base / "config_paths.json")


CONFIG_FILE = get_app_data_path()


def month_sheet_name(date):
    return f"{['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][date.month - 1]} {date.year}"


def parse_month_sheet(name):
    """Independent of the operating system locale."""
    months = {m.lower(): i for i, m in enumerate(
        ["January", "February", "March", "April", "May", "June",
         "July", "August", "September", "October", "November", "December"], 1)}
    match = re.fullmatch(r"([A-Za-z]+) ([0-9]{4})", name.strip())
    if not match or match.group(1).lower() not in months:
        raise ValueError(tr(f"Invalid sheet name format: {name!r}; expected 'January 2026'", f"Μη έγκυρο όνομα φύλλου: {name!r}; αναμένεται 'January 2026'"))
    return datetime(int(match.group(2)), months[match.group(1).lower()], 1)


def to_float(value, *, blank_zero=True):
    """Accept 75,50 and 75.50; reject formulas, errors and non-finite values."""
    if value is None or (isinstance(value, str) and not value.strip()):
        if blank_zero:
            return 0.0
        raise ValueError(tr('A numeric value is required', 'Απαιτείται αριθμητική τιμή'))
    if isinstance(value, bool):
        raise ValueError(tr('Boolean is not a numeric value', 'Η λογική τιμή δεν είναι αριθμός'))
    try:
        result = float(value.strip().replace(",", ".") if isinstance(value, str) else value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(tr(f'Invalid numeric value: {value!r}', f'Μη έγκυρη αριθμητική τιμή: {value!r}')) from exc
    if not math.isfinite(result):
        raise ValueError(tr('Number must be finite', 'Ο αριθμός πρέπει να είναι πεπερασμένος'))
    return result


@contextmanager
def excel_app():
    """Use locally installed desktop Microsoft Excel. Never save opened source files."""
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


def read_debts_from_excel(path, sheet_name, rows):
    """Read freshly recalculated formula values before openpyxl modifies the file."""
    with excel_app() as app:
        book = app.books.open(os.path.abspath(path), update_links=False, read_only=True)
        try:
            app.calculate()
            sheet = book.sheets[sheet_name]
            data = []
            for row in rows:
                numbers = [to_float(sheet.range(f"{col}{row}").value)
                           for col in ("S", "T", "U")]
                data.append(tuple(numbers))
            return data
        finally:
            book.close()


def save_workbooks_safely(wb, path, wb2, path2):
    """Stage both saves and attempt rollback if either replacement fails.

    Temporary rollback copies are removed immediately, not retained as backups.
    No cross-file atomicity is possible after sudden power loss.
    """
    paths = [Path(path).resolve(), Path(path2).resolve()]
    if paths[0] == paths[1]:
        raise ValueError(tr('The two workbooks must be different files', 'Τα δύο αρχεία Excel πρέπει να είναι διαφορετικά'))
    staged, originals, committed = [], [], []
    try:
        for workbook, original in zip((wb, wb2), paths):
            if not original.is_file():
                raise FileNotFoundError(original)
            fd, tmp = tempfile.mkstemp(prefix=".bem-new-", suffix=".xlsx", dir=original.parent)
            os.close(fd)
            staged.append(Path(tmp))
            workbook.save(tmp)
            fd, old = tempfile.mkstemp(prefix=".bem-rollback-", suffix=".xlsx", dir=original.parent)
            os.close(fd)
            originals.append(Path(old))
            shutil.copy2(original, old)
        for new, original in zip(staged, paths):
            os.replace(new, original)
            committed.append(original)
    except Exception as exc:
        rollback_errors = []
        for i, original in enumerate(paths):
            if original in committed and i < len(originals):
                try:
                    shutil.copy2(originals[i], original)
                except OSError as rollback_exc:
                    rollback_errors.append(f"{original}: {rollback_exc}")
        if rollback_errors:
            raise RuntimeError(tr('Save failed AND rollback failed: ', 'Απέτυχαν η αποθήκευση και η επαναφορά: ') + "; ".join(rollback_errors)) from exc
        raise
    finally:
        for tmp in staged + originals:
            tmp.unlink(missing_ok=True)


def validate_template_layout(wb, wb2):
    """Ensure both workbooks match the 15-apartment templates this app supports."""
    common = wb[wb.sheetnames[-1]]
    ledger = wb2[wb2.sheetnames[-1]]
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
    if first != second:
        raise ValueError(tr(f'Last sheets do not match: {first!r} vs {second!r}', f'Τα τελευταία φύλλα δεν συμφωνούν: {first!r} έναντι {second!r}'))
    parse_month_sheet(first)
    expected_next = month_sheet_name(parse_month_sheet(first) + relativedelta(months=1))
    if expected_next in wb.sheetnames or expected_next in wb2.sheetnames:
        raise ValueError(tr(f'Sheet {expected_next!r} already exists', f'Το φύλλο {expected_next!r} υπάρχει ήδη'))
    validate_template_layout(wb, wb2)
    return first


def fix_heating_formulas(sheet):
    """Zero means exempt; blank means full participation. Keep all other formulas."""
    for row in range(36, 51):
        c = row - 22
        sheet[f"I{row}"] = f'=IF(C{c}>0,IF($I$54>0,C{c}*IF(H{row}="",1,H{row})/$I$54*100,0),"")'


def validate_heating_expense(sheet, amount):
    if amount > 0 and not any(
        to_float(sheet[f"C{r}"].value) > 0 and
        (sheet[f"H{r+22}"].value in (None, "") or to_float(sheet[f"H{r+22}"].value) > 0)
        for r in range(14, 29)
    ):
        raise ValueError(tr('Heating expense cannot be allocated: all apartments have zero participation.', 'Η δαπάνη θέρμανσης δεν μπορεί να κατανεμηθεί: δεν υπάρχει ιδιοκτησία με θετική συμμετοχή.'))


def create_month_sheets(wb, wb2, previous_name):
    """Prepare the next month only in memory; caller decides whether to commit."""
    previous_month = parse_month_sheet(previous_name)
    new_month = previous_month + relativedelta(months=1)
    new_name = month_sheet_name(new_month)
    title = f"{GREEK_MONTHS[new_month.month]} {new_month.year}"

    common = wb.copy_worksheet(wb[previous_name])
    common.title = new_name
    fix_heating_formulas(common)
    common.conditional_formatting = deepcopy(wb[previous_name].conditional_formatting)
    for address in BILL_CELLS_TO_CLEAR:
        common[address] = None

    ledger = wb2.copy_worksheet(wb2[previous_name])
    ledger.title = new_name
    ledger.conditional_formatting = deepcopy(wb2[previous_name].conditional_formatting)
    for address in CELLS_TO_CLEAR_2:
        ledger[address] = None

    common["O1"] = title
    ledger["E3"] = title
    ledger["E8"] = f"{GREEK_MONTHS[previous_month.month]} {previous_month.year}"
    return new_name, new_month, title, common, ledger


def apply_payment(m, n, paid):
    if paid >= m:
        paid -= m
        m = 0
    else:
        m -= paid
        paid = 0

    if paid != 0:
        n -= paid

    return m, n


def check_first_time_setup(path, path2):
    wb = load_workbook(path)
    ws = wb[wb.sheetnames[-1]]

    owners = ws["E14":"E28"]
    all_empty = all(cell.value in (None, "", " ") for row in owners for cell in row)

    wb.close()

    if not all_empty:
        return "not_first_time"

    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(tr('Setup', 'Ρύθμιση'), tr('First-time setup', 'Αρχική ρύθμιση'))

    def open_dialog():
        dialog = tk.Toplevel()
        dialog.title(tr('First-time setup', 'Αρχική ρύθμιση'))

        tk.Label(dialog, text=tr('Enter available management balance:', 'Διαθέσιμο υπόλοιπο διαχείρισης:')).pack(pady=5)
        entry1 = tk.Entry(dialog)
        entry1.pack(pady=5)

        tk.Label(dialog, text=tr('Street Name and Number:', 'Οδός και αριθμός:')).pack(pady=5)
        entry2 = tk.Entry(dialog)
        entry2.pack(pady=5)

        tk.Label(dialog, text=tr('Select first report month:', 'Επιλέξτε τον πρώτο μήνα έκδοσης:')).pack(pady=5)
        months = DISPLAY_MONTHS
        month_cb = ttk.Combobox(dialog, values=months, state="readonly")
        month_cb.current(0)
        month_cb.pack(pady=5)

        tk.Label(dialog, text=tr('Select Year:', 'Επιλέξτε έτος:')).pack(pady=5)
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 5, current_year + 6)]
        year_cb = ttk.Combobox(dialog, values=years, state="readonly")
        year_cb.set(str(current_year))
        year_cb.pack(pady=5)

        result = {}

        def submit():
            try:
                val1 = to_float(entry1.get(), blank_zero=False)
                if val1 < 0:
                    raise ValueError

                val2 = entry2.get()
                month = month_cb.get()
                year = int(year_cb.get())

                month_index = months.index(month) + 1

                if month_index == 1:
                    prev_month = 12
                    prev_year = year - 1
                else:
                    prev_month = month_index - 1
                    prev_year = year

                sheet_name = month_sheet_name(datetime(prev_year, prev_month, 1))

                result["value1"] = val1
                result["value2"] = val2
                result["sheet_name"] = sheet_name

                dialog.destroy()

            except Exception as e:
                messagebox.showerror(tr('Error', 'Σφάλμα'), tr(f'Invalid input: {e}', f'Μη έγκυρη εισαγωγή: {e}'))

        tk.Button(dialog, text=tr('Submit', 'Καταχώριση'), command=submit).pack(pady=10)

        dialog.grab_set()
        dialog.wait_window()

        return result

    data = open_dialog()
    root.destroy()

    if not data:
        return "cancel"

    wb1 = load_workbook(path)
    wb2 = load_workbook(path2)

    if len(wb1.sheetnames) != len(wb2.sheetnames):
        wb1.close()
        wb2.close()
        raise ValueError(tr('Workbook sheet counts differ; first-time setup canceled', 'Διαφορετικός αριθμός φύλλων. Η αρχική ρύθμιση ακυρώθηκε.'))
    ws1 = wb1[wb1.sheetnames[-1]]
    ws2 = wb2[wb2.sheetnames[-1]]

    ws2["F8"].value = data['value1']
    ws2["K3"].value = data['value2']
    ws1["E1"].value = data['value2']

    if len(wb1.sheetnames) == 1:
        ws1.title = data['sheet_name']
    if len(wb2.sheetnames) == 1:
        ws2.title = data['sheet_name']

    save_workbooks_safely(wb1, path, wb2, path2)

    wb1.close()
    wb2.close()

    return "first_time_done"


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


def delete_month_files(wb, path, wb2, path2, expected_name):
    name, month = deletion_target(wb, wb2)
    if name != expected_name:
        raise ValueError(tr('The last month changed. Try again.', 'Ο τελευταίος μήνας άλλαξε. Δοκιμάστε ξανά.'))
    pdf = month_pdf_path(path, month)
    staged_pdf = None
    if pdf.exists():
        if not pdf.is_file() or pdf.is_symlink():
            raise ValueError(tr(f'Not a regular PDF file: {pdf}', f'Μη έγκυρο αρχείο PDF: {pdf}'))
        # Rename before committing Excel so an open/locked PDF stops deletion.
        fd, temp = tempfile.mkstemp(prefix=".bem-delete-", suffix=".pdf", dir=pdf.parent)
        os.close(fd)
        staged_pdf = Path(temp)
        try:
            os.replace(pdf, staged_pdf)
        except Exception:
            staged_pdf.unlink(missing_ok=True)
            raise
    try:
        del wb[name]
        del wb2[name]
        validate_template_layout(wb, wb2)
        save_workbooks_safely(wb, path, wb2, path2)
    except BaseException:
        if staged_pdf is not None:
            os.replace(staged_pdf, pdf)
        raise
    if staged_pdf is not None:
        try:
            staged_pdf.unlink()
        except OSError as exc:
            return tr(f'The sheets were deleted, but the temporary PDF must be deleted manually: {staged_pdf} ({exc})', f'Τα φύλλα διαγράφηκαν, αλλά απαιτείται χειροκίνητη διαγραφή του προσωρινού PDF: {staged_pdf} ({exc})')
    return None


def launch_delete_last_month(path, path2):
    try:
        with_context = (load_workbook(path), load_workbook(path2))
        wb, wb2 = with_context
        try:
            name, month = deletion_target(wb, wb2)
            pdf = month_pdf_path(path, month)
            detail = f"PDF: {pdf.name}" if pdf.exists() else tr('No matching PDF was found.', 'Δεν βρέθηκε αντίστοιχο PDF.')
            if not messagebox.askyesno(
                tr('Confirm deletion', 'Επιβεβαίωση διαγραφής'),
                tr(f"Are you sure you want to delete the sheets '{name}' ({DISPLAY_MONTHS[month.month - 1]} {month.year}) from both Excel workbooks?\n\n{detail}\n\nDeletion is permanent.", f"Είστε σίγουρος ότι θέλετε να διαγράψετε τις καρτέλες '{name}' ({DISPLAY_MONTHS[month.month - 1]} {month.year}) και από τα δύο Excel;\n\n{detail}\n\nΗ διαγραφή είναι οριστική."), default="no"
            ):
                return
            warning = delete_month_files(wb, path, wb2, path2, name)
            if warning:
                messagebox.showwarning(tr('Deletion', 'Διαγραφή'), warning)
            else:
                messagebox.showinfo(tr('Deletion', 'Διαγραφή'), tr(f'Month {name} was deleted from both Excel workbooks.\n', f'Ο μήνας {name} διαγράφηκε και από τα δύο Excel.\n') +
                                    (tr('The matching PDF no longer exists.', 'Το αντίστοιχο PDF δεν υπάρχει πλέον.')))
        finally:
            wb.close()
            wb2.close()
    except Exception as exc:
        messagebox.showerror(tr('Deletion error', 'Σφάλμα διαγραφής'), str(exc))


def launch_welcome_dashboard():
    root = tk.Tk()
    root.title(tr('Shared Debt Management', 'Διαχείριση Κοινοχρήστων'))
    root.geometry("560x480")

    title = tk.Label(
        root,
        text=tr('Shared Debt Management', 'Διαχείριση Κοινοχρήστων'),
        font=("Helvetica", 20, "bold"),
    )
    title.pack(pady=40)

    subtitle = tk.Label(
        root,
        text=tr('Monthly Utilities Automation System', 'Μηνιαία αυτοματοποίηση κοινοχρήστων'),
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

    tk.Button(btn_frame, text=tr('Start Program', 'Έναρξη προγράμματος'), width=20, command=start).pack(pady=10)
    tk.Button(btn_frame, text=tr('Edit Data', 'Επεξεργασία στοιχείων'), width=24, command=edit).pack(pady=10)
    tk.Button(btn_frame, text=tr('Delete last month', 'Διαγραφή τελευταίου φύλλου'), width=28, command=delete).pack(pady=10)

    ttk.Button(btn_frame, text="Language / Γλώσσα", command=change_language, width=28).pack(pady=10)

    root.mainloop()
    return choice["action"]


def launch_edit_dashboard(path, path2):
    root = tk.Tk()
    root.title(tr('Edit Data', 'Επεξεργασία στοιχείων'))
    root.geometry(tr('820x650', '920x690'))

    wb = load_workbook(path)
    ws = wb[wb.sheetnames[-1]]

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

        f = ws[f"B{14 + i}"].value if ws[f"B{14 + i}"].value is not None else ""
        c = ws[f"C{14 + i}"].value if ws[f"C{14 + i}"].value is not None else ""
        d = ws[f"D{14 + i}"].value or ""
        e = ws[f"E{14 + i}"].value or ""
        h = ws[f"H{36 + i}"].value if ws[f"H{36 + i}"].value is not None else ""

        floor_var = tk.StringVar(value=FLOOR_LABELS.get(str(f), str(f)))
        floor_cb = ttk.Combobox(frame, textvariable=floor_var, values=[FLOOR_LABELS[x] for x in floor_options], state="readonly", width=13)
        floor_cb.grid(row=i + 1, column=1)
        row_entries.append(floor_var)

        for j, val in enumerate([c, d, e, h]):
            var = tk.StringVar(value=str(val))
            tk.Entry(frame, textvariable=var, width=15).grid(row=i + 1, column=j + 2)
            row_entries.append(var)

        entries.append(row_entries)

    fee_var = tk.StringVar(value=str(ws["D8"].value if ws["D8"].value is not None else ""))
    wb.close()

    tk.Label(root, text=tr('Heating: 0 = 0% | 0.25 = 25% | 1 or blank = 100%', 'Θέρμανση: 0 = 0% | 0,25 = 25% | 1 ή κενό = 100%'),
             wraplength=760).pack(pady=5)
    tk.Label(root, text=tr('Management Fee', 'Αμοιβή διαχείρισης')).pack(pady=10)
    tk.Entry(root, textvariable=fee_var).pack()

    def finish():
        if not messagebox.askyesno(tr('Confirm', 'Επιβεβαίωση'), tr('Save changes?', 'Αποθήκευση αλλαγών;')):
            return
        try:
            square_metres = []
            heating = []
            for i, row in enumerate(entries, start=1):
                sqm = row[1].get().strip()
                sqm_number = to_float(sqm, blank_zero=False) if sqm else None
                if sqm_number is not None and sqm_number <= 0:
                    raise ValueError(tr(f'Apartment {i}: square metres must be positive', f'Ιδιοκτησία {i}: τα τετραγωνικά πρέπει να είναι θετικά'))
                square_metres.append(sqm_number)
                heating_text = row[4].get().strip()
                factor = to_float(heating_text) if heating_text else None
                if factor is not None and not 0 <= factor <= 1:
                    raise ValueError(tr(f'Apartment {i}: heating factor must be between 0 and 1 (25% = 0.25)', f'Ιδιοκτησία {i}: ο συντελεστής θέρμανσης πρέπει να είναι από 0 έως 1 (25% = 0,25)'))
                heating.append(factor)
            management_fee = to_float(fee_var.get()) if fee_var.get().strip() else None
            if management_fee is not None and management_fee < 0:
                raise ValueError(tr('Management fee cannot be negative', 'Η αμοιβή διαχείρισης δεν μπορεί να είναι αρνητική'))
            wb = load_workbook(path)
            wb2 = load_workbook(path2)
            try:
                ws = wb[wb.sheetnames[-1]]
                ws2 = wb2[wb2.sheetnames[-1]]
                for i, row in enumerate(entries):
                    ws[f"B{14+i}"].value = next((k for k, v in FLOOR_LABELS.items() if v == row[0].get()), row[0].get())
                    ws[f"C{14+i}"].value = square_metres[i]
                    ws[f"D{14+i}"].value = row[2].get()
                    ws[f"E{14+i}"].value = row[3].get()
                    ws[f"H{36+i}"].value = heating[i]
                    ws2[f"E{9+i}"].value = row[3].get()
                ws["D8"].value = management_fee
                fix_heating_formulas(ws)
                validate_heating_expense(ws, to_float(ws["G4"].value))
                save_workbooks_safely(wb, path, wb2, path2)
            finally:
                wb.close()
                wb2.close()
        except Exception as exc:
            messagebox.showerror(tr('Edit Data Error', 'Σφάλμα επεξεργασίας στοιχείων'), str(exc))
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
    root.title(tr('Utilities Dashboard', 'Πληρωμές κοινοχρήστων'))
    root.geometry("920x580")

    payments = [0] * len(owners)
    full_paid_flags = [False] * len(owners)

    payment_vars = []
    full_vars = []

    scroll_frame = tk.Frame(root)
    scroll_frame.pack(fill="both", expand=True)

    bottom_frame = tk.Frame(root)
    bottom_frame.pack(side="bottom", fill="x", pady=10)

    headers = [tr('Owner', 'Ιδιοκτήτης'), tr('Tenant', 'Ενοικιαστής'), tr('Total Debt', 'Συνολική οφειλή'), tr('Payment', 'Πληρωμή'), tr('Full Payment', 'Εξόφληση')]
    for col, text in enumerate(headers):
        tk.Label(scroll_frame, text=text, font=("Arial", 10, "bold")).grid(row=0, column=col, padx=5, pady=5)

    for i, owner in enumerate(owners):
        total_debt = tenant_debts[i] + owner_debts[i]

        tk.Label(scroll_frame, text=owner, wraplength=235, justify="left").grid(row=i+1, column=0, sticky="w", padx=5)
        tk.Label(scroll_frame, text=tenants[i] or "—", wraplength=235, justify="left").grid(row=i+1, column=1, sticky="w", padx=5)
        tk.Label(scroll_frame, text=f"{total_debt:.2f}").grid(row=i+1, column=2)

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
                        tr(f'{owners[i]}: Cannot enter payment AND select full payment.', f'{owners[i]}: Δεν επιτρέπεται ταυτόχρονη καταχώριση πληρωμής και εξόφλησης.')
                    )
                    return

                payment = to_float(val) if val else 0.0

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
        except ValueError:
            messagebox.showerror(tr('Error', 'Σφάλμα'), tr('Invalid number entered', 'Μη έγκυρος αριθμός'))

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


def launch_bills_dashboard():
    root = tk.Tk()
    root.title(tr('Bills Input', 'Καταχώριση δαπανών'))
    root.geometry("530x780")

    entries = {}

    bills = [
        "Φ.Α.",
        "Ηλεκτ. ρεύμα κοινοχρήστων",
        "ΕΥΔΑΠ",
        "Καθαρισμός κοιν/στων",
        "Αναλώσιμα"
    ]

    tk.Label(root, text=tr('Bills Input', 'Καταχώριση δαπανών'), font=("Arial", 12, "bold")).pack(pady=10)

    frame = tk.Frame(root)
    frame.pack()

    for name in bills:
        row = tk.Frame(frame)
        row.pack(fill="x", pady=2)

        tk.Label(row, text=BILL_LABELS[name], width=34, anchor="w").pack(side="left")

        var = tk.StringVar()
        tk.Entry(row, textvariable=var, width=10).pack(side="right")

        entries[name] = var

    tk.Label(root, text=tr('Urgent expenses are charged to owners.\nExtra expenses are charged to tenants.', 'Έκτακτες δαπάνες ιδιοκτητών: χρεώνονται στους ιδιοκτήτες.\nΠρόσθετες δαπάνες ενοικιαστών: χρεώνονται στους ενοικιαστές.'), wraplength=480, justify="left").pack(pady=8)

    urgent_desc = tk.StringVar()
    urgent_amount = tk.StringVar()

    tk.Label(root, text=tr('Urgent expense 1 (owner)', 'Έκτακτη δαπάνη ιδιοκτητών 1')).pack(pady=5)
    tk.Entry(root, textvariable=urgent_desc).pack()
    tk.Entry(root, textvariable=urgent_amount).pack()

    urgent2_desc = tk.StringVar()
    urgent2_amount = tk.StringVar()

    tk.Label(root, text=tr('Urgent expense 2 (owner)', 'Έκτακτη δαπάνη ιδιοκτητών 2')).pack(pady=5)
    tk.Entry(root, textvariable=urgent2_desc).pack()
    tk.Entry(root, textvariable=urgent2_amount).pack()

    extra1_desc = tk.StringVar()
    extra1_amount = tk.StringVar()

    tk.Label(root, text=tr('Extra expense 1 (tenant)', 'Πρόσθετη δαπάνη ενοικιαστών 1')).pack(pady=5)
    tk.Entry(root, textvariable=extra1_desc).pack()
    tk.Entry(root, textvariable=extra1_amount).pack()

    extra2_desc = tk.StringVar()
    extra2_amount = tk.StringVar()

    tk.Label(root, text=tr('Extra expense 2 (tenant)', 'Πρόσθετη δαπάνη ενοικιαστών 2')).pack(pady=5)
    tk.Entry(root, textvariable=extra2_desc).pack()
    tk.Entry(root, textvariable=extra2_amount).pack()

    result = {}

    def collect_bills_data(generate_pdf=False):
        try:
            data = {}
            data["standard"] = {}

            for k, v in entries.items():
                val = to_float(v.get())

                if val < 0:
                    messagebox.showerror(tr('Error', 'Σφάλμα'), tr(f'{BILL_LABELS[k]}: Negative values not allowed', f'{BILL_LABELS[k]}: Δεν επιτρέπονται αρνητικές τιμές'))
                    return None

                data["standard"][k] = val

            data["urgent"] = (
                urgent_desc.get(),
                to_float(urgent_amount.get())
            )

            data["urgent2"] = (
                urgent2_desc.get(),
                to_float(urgent2_amount.get())
            )

            data["extra1"] = (
                extra1_desc.get(),
                to_float(extra1_amount.get())
            )

            data["extra2"] = (
                extra2_desc.get(),
                to_float(extra2_amount.get())
            )

            for key in ("urgent", "urgent2", "extra1", "extra2"):
                if data[key][1] < 0:
                    raise ValueError(tr(f'{EXPENSE_LABELS[key]}: negative values not allowed', f'{EXPENSE_LABELS[key]}: δεν επιτρέπονται αρνητικές τιμές'))
                if data[key][1] > 0 and not data[key][0].strip():
                    raise ValueError(tr(f'{EXPENSE_LABELS[key]}: description is required', f'{EXPENSE_LABELS[key]}: απαιτείται περιγραφή'))

            if generate_pdf:
                data["generate_pdf"] = True

            return data

        except Exception as e:
            messagebox.showerror(tr('Error', 'Σφάλμα'), f"{e}")
            return None

    def submit():
        data = collect_bills_data(generate_pdf=False)
        if data is None:
            return

        result.update(data)
        root.destroy()

    def submit_and_generate():
        data = collect_bills_data(generate_pdf=True)
        if data is None:
            return

        result.update(data)
        root.destroy()

    def on_close():
        result["status"] = "quit"
        if root.winfo_exists():
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Button(root, text=tr('Finish without PDF', 'Ολοκλήρωση χωρίς PDF'), command=submit).pack(pady=5)
    tk.Button(root, text=tr('Generate PDF & Finish', 'Δημιουργία PDF και ολοκλήρωση'), command=submit_and_generate).pack(pady=10)

    root.mainloop()

    if result.get("status") == "quit":
        return None, "quit"

    return result, "next"


def load_paths():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}
    return {}


def save_paths(data):
    settings = load_paths()
    settings.update(data)
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


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
                tr(f'Selected file is not valid.\n\nExpected file containing:\n{expected_keyword}', f'Το επιλεγμένο αρχείο δεν είναι έγκυρο.\n\nΑναμένεται όνομα αρχείου που περιέχει:\n{expected_keyword}')
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


def main():
    if not configure_language():
        return
    new_name = None

    paths = load_paths()

    path = paths.get("file1")
    path2 = paths.get("file2")

    if not path or not os.path.exists(path):
        messagebox.showinfo(tr('Select File', 'Επιλογή αρχείου'), tr('Please select the common expenses Excel file (ΚΟΙΝΟΧΡΗΣΤΑ).', 'Επιλέξτε το αρχείο Excel ΚΟΙΝΟΧΡΗΣΤΑ.'))
        path = select_file(tr('Select ΚΟΙΝΟΧΡΗΣΤΑ file', 'Επιλογή αρχείου ΚΟΙΝΟΧΡΗΣΤΑ'), "ΚΟΙΝΟΧΡΗΣΤΑ")

    if not path2 or not os.path.exists(path2):
        messagebox.showinfo(tr('Select File', 'Επιλογή αρχείου'), tr('Please select the accounting ledger Excel file (ΒΙΒΛΙΟ).', 'Επιλέξτε το αρχείο Excel ΒΙΒΛΙΟ.'))
        path2 = select_file(tr('Select ΒΙΒΛΙΟ file', 'Επιλογή αρχείου ΒΙΒΛΙΟ'), "ΒΙΒΛΙΟ")

    if not path or not path2:
        return

    if os.path.samefile(path, path2):
        messagebox.showerror(tr('Invalid Files', 'Μη έγκυρα αρχεία'), tr('Choose two different Excel files', 'Επιλέξτε δύο διαφορετικά αρχεία Excel'))
        return

    save_paths({
        "file1": path,
        "file2": path2
    })

    while True:
        try:
            while True:
                action = launch_welcome_dashboard()

                if action == "language":
                    configure_language(force=True)
                    continue

                if action == "edit":
                    launch_edit_dashboard(path, path2)
                    continue

                elif action == "delete":
                    launch_delete_last_month(path, path2)
                    continue

                elif action == "start":
                    result = check_first_time_setup(path, path2)

                    if result == "cancel":
                        continue

                    elif result == "first_time_done":
                        launch_edit_dashboard(path, path2)
                        continue

                    elif result == "not_first_time":
                        break

                if action is None:
                    return

            # ---------------------------
            # STEP 2: VALIDATE AND CREATE MONTH IN MEMORY
            # ---------------------------
            wb = load_workbook(path)
            wb2 = load_workbook(path2)
            last_sheet_name = validate_workbook_pair(wb, wb2)
            prev_dt = parse_month_sheet(last_sheet_name)
            new_name, new_dt, greek_month_text, new_ws, new_ws2 = create_month_sheets(
                wb, wb2, last_sheet_name
            )

            messagebox.showinfo(
                tr('Setup', 'Ρύθμιση'),
                "\n" + "=" * 48 + "\n"
                + tr('ENTER DATA FOR THE MONTH:\n', 'ΕΙΣΑΓΩΓΗ ΣΤΟΙΧΕΙΩΝ ΓΙΑ ΤΟΝ ΜΗΝΑ:\n')
                + f"{DISPLAY_MONTHS[new_dt.month-1]} {new_dt.year}\n"
                + "=" * 48
            )

            # Workbooks remain in memory until the user completes both dialogs.

            # ---------------------------
            # STEP 4: READ DATA FROM OPENPYXL SHEETS
            # ---------------------------
            owners_with_rows = [
                (row, new_ws[f"E{row}"].value)
                for row in range(14, 29)
                if new_ws[f"E{row}"].value not in (None, "", " ")
            ]
            OWNERS = [owner for _, owner in owners_with_rows]
            row_numbers = [row for row, _ in owners_with_rows]
            tenants = [new_ws[f"D{row}"].value for row in row_numbers]
            debts = read_debts_from_excel(path, last_sheet_name, row_numbers)
            tenant_debts = [triple[0] for triple in debts]
            owner_debts = [triple[1] for triple in debts]
            AMOUNTS_OWED = [triple[2] for triple in debts]

            # ---------------------------
            # STEP 5: PAYMENTS GUI
            # ---------------------------
            payments, full_flags, status = launch_dashboard(OWNERS, tenant_debts, owner_debts, tenants)

            if status == "quit":
                wb.close()
                wb2.close()
                continue

            for i in range(len(OWNERS)):
                m = tenant_debts[i]
                n = owner_debts[i]

                if full_flags[i]:
                    paid = AMOUNTS_OWED[i]
                    m, n = 0.0, 0.0
                else:
                    paid = payments[i]
                    if paid > 0:
                        m, n = apply_payment(m, n, paid)

                r = row_numbers[i]
                new_ws[f"M{r}"].value = m
                new_ws[f"N{r}"].value = n

                new_ws2[f"F{r - 5}"].value = paid if paid != 0 else None

            # ---------------------------
            # STEP 6: BILLS
            # ---------------------------
            bills_data, status = launch_bills_dashboard()

            if status == "quit":
                wb.close()
                wb2.close()
                continue

            validate_heating_expense(new_ws, bills_data["standard"]["Φ.Α."])
            generate_pdf_flag = bills_data.get("generate_pdf", False)

            bills_map = {
                "Φ.Α.": "G4",
                "Ηλεκτ. ρεύμα κοινοχρήστων": "G5",
                "ΕΥΔΑΠ": "G6",
                "Καθαρισμός κοιν/στων": "G7",
                "Αναλώσιμα": "G8"
            }

            bills_map2 = {
                "Φ.Α.": "L10",
                "Ηλεκτ. ρεύμα κοινοχρήστων": "L9",
                "ΕΥΔΑΠ": "L11",
                "Καθαρισμός κοιν/στων": "L8",
                "Αναλώσιμα": "L12"
            }

            for name, amount in bills_data["standard"].items():
                if amount == 0:
                    new_ws[bills_map[name]].value = None
                    new_ws2[bills_map2[name]].value = None
                else:
                    new_ws[bills_map[name]].value = amount
                    new_ws2[bills_map2[name]].value = amount

            desc, amount = bills_data["urgent"]
            if amount > 0:
                new_ws["A6"].value = desc
                new_ws["D6"].value = amount
                new_ws2["K17"].value = desc
                new_ws2["L17"].value = amount

            desc, amount = bills_data["urgent2"]
            if amount > 0:
                new_ws["A7"].value = desc
                new_ws["D7"].value = amount
                new_ws2["K18"].value = desc
                new_ws2["L18"].value = amount

            desc, amount = bills_data["extra1"]
            if amount > 0:
                new_ws["E9"].value = desc
                new_ws["G9"].value = amount
                new_ws2["K13"].value = desc
                new_ws2["L13"].value = amount

            desc, amount = bills_data["extra2"]
            if amount > 0:
                new_ws["E10"].value = desc
                new_ws["G10"].value = amount
                new_ws2["K14"].value = desc
                new_ws2["L14"].value = amount

            if not messagebox.askyesno(tr('Confirm', 'Επιβεβαίωση'), tr('Save both updated Excel workbooks?', 'Αποθήκευση και των δύο ενημερωμένων αρχείων Excel;')):
                wb.close()
                wb2.close()
                continue
            save_workbooks_safely(wb, path, wb2, path2)
            wb.close()
            wb2.close()

            # ---------------------------
            # STEP 7: PDF
            # ---------------------------
            if generate_pdf_flag:
                pdf_path = str(month_pdf_path(path, new_dt))
                try:
                    with excel_app() as app:
                        book_xw = app.books.open(os.path.abspath(path), update_links=False)
                        try:
                            sheet = book_xw.sheets[new_name]
                            app.calculate()
                            sheet.page_setup.print_area = "$A$1:$P$31"
                            sheet.to_pdf(path=pdf_path, show=False)
                        finally:
                            book_xw.close()
                    messagebox.showinfo(tr('New File', 'Νέο αρχείο'), tr(f'PDF generated:\n{pdf_path}', f'Δημιουργήθηκε PDF:\n{pdf_path}'))
                except Exception as exc:
                    messagebox.showerror(
                        tr('PDF Export Error', 'Σφάλμα εξαγωγής PDF'),
                        tr(f'Excel files were saved successfully, but PDF export failed.\nReason: {exc}\nYou can retry PDF export manually in Excel.', f'Τα αρχεία Excel αποθηκεύτηκαν, αλλά η εξαγωγή PDF απέτυχε.\nΑιτία: {exc}\nΜπορείτε να επαναλάβετε την εξαγωγή PDF χειροκίνητα στο Excel.')
                    )

            # ---------------------------
            # CLEAN EXIT PER LOOP
            # ---------------------------
            wb.close()
            wb2.close()
            return

        except (KeyboardInterrupt, Exception) as exc:
            for workbook_name in ("wb", "wb2"):
                workbook = locals().get(workbook_name)
                if workbook is not None:
                    try:
                        workbook.close()
                    except Exception:
                        pass
            messagebox.showerror(tr('Error', 'Σφάλμα'), str(exc) if str(exc) else tr('Process interrupted', 'Η διαδικασία διακόπηκε'))
            return


if __name__ == "__main__":
    main()