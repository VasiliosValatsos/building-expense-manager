# Testing

## Evidence and scope

The author reports completion of manual acceptance tests on Windows 10 and macOS. The exact final unified `main.py` was not supplied when this companion pack was prepared. The automated tests below are candidate regression tests based on the previously agreed helper interface. They must be run against the actual final code; no passing result is claimed here. If helper names or signatures changed, adapt the tests to the new interface rather than changing the working program merely to satisfy outdated assumptions.

The test runner uses the Python standard library `unittest`. Use `python -m unittest discover -s tests -v` after adding `main.py` and both templates. On macOS replace `python` with `python3` if necessary. CI uses Python 3.11 and 3.12 on Windows and macOS, with no desktop Excel execution. It does not certify Windows 10 compatibility or GUI rendering.

## Acceptance checklist (use copies, not live or published templates)

- [ ] Launch and select both languages; verify all labels and dialogs.
- [ ] Initial setup and all 15 slots; decimal comma/point entries; cancel edits without saving.
- [ ] Payments: zero, partial, full, overpayment, negative rejection, checkbox conflicts and gaps in occupied slots.
- [ ] Heating: zero, 0.25, 0.5, 1, blank, invalid factors and no participant with a positive bill.
- [ ] For five equal-area properties with factors 1, 0.5, 0.25, 0 and blank, allocate a heating bill of 275 EUR: expected charges 100, 50, 25, 0, 100 EUR.
- [ ] Standard, owner and tenant expenses map to the intended cells and reconcile with the ledger.
- [ ] Create three successive months including December to January; historical sheet contents remain unchanged.
- [ ] Cancel at payments and expenses; no new monthly sheets should remain.
- [ ] PDF export, correct month in filename, legibility and print layout; save without PDF.
- [ ] Delete latest month: cancel, confirm, missing PDF, mismatched workbooks, protect sole sheet, recreate month.
- [ ] Confirm failure handling with closed/open documents and temporary save errors using disposable copies.
- [ ] Verify no actual personal data in templates, screenshots, sample PDF or metadata.

## Record exact tested versions

Run `python --version` and `python -m pip show openpyxl xlwings python-dateutil`. Record versions from both operating systems below or in a release note. Dependency ranges are provisional, not a reproducibility lockfile.

## Tested Development Environment

| Component | Version |
|---|---|
| Operating System | Windows 10 |
| Python | 3.15.2 |
| openpyxl | 3.1.5 |
| xlwings | 0.37.10 |
| python-dateutil | 2.9.0.post0 |
| pywin32 | 312 |
| six | 1.17.0 |
| et-xmlfile | 2.0.0 |

Dependencies were successfully installed in an isolated Python virtual environment.

The application has also been manually tested on macOS.

Automated regression tests are executed using Python's built-in unittest framework. Microsoft Excel integration and PDF generation are validated separately through manual testing.