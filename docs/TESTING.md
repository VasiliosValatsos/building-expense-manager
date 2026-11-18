# Testing

## Evidence and scope

The unittest suite exercises cent-level payments, settlement and allocation, the cash roll-forward, the management fee, apartment transitions, template layout, error handling and workbook recovery. Tests for calculated cells use a simulated Excel reader. Use the checklist below for native Excel and GUI verification.

The test runner uses the Python standard library `unittest`. Use `python -m unittest discover -s tests -v` after adding `main.py` and both templates. On macOS replace `python` with `python3` if necessary. CI uses Python 3.11 and 3.12 on Windows and macOS, with no desktop Excel execution. It does not certify Windows 10 compatibility or GUI rendering.

## Acceptance checklist (use copies, not live or published templates)

- [ ] Launch and select both languages; verify all labels and dialogs.
- [ ] Initial setup and all 15 slots; decimal comma/point entries; cancel edits without saving; confirm a public template cannot be selected as a working file.
- [ ] Payments: zero, partial, full, overpayment, negative rejection, checkbox conflicts and gaps in occupied slots. Check a negative net balance remains credit after full settlement and enters no cash income.
- [ ] Heating: zero, 0.25, 0.5, 1, blank, invalid factors and no participant with a positive bill.
- [ ] For five equal-area properties with factors 1, 0.5, 0.25, 0 and blank, allocate a heating bill of 275 EUR: expected charges 100, 50, 25, 0, 100 EUR.
- [ ] Standard, owner and tenant expenses map to the intended cells. When other expenses are EUR 322.83 and the fee is EUR 60, check total charges and cash expenses of EUR 382.83. The fee is included once in `D8` and `L16`, independently of any receipts.
- [ ] Start with EUR 1,000 and a EUR 60 fee. Create three successive months including December to January. Enter receipts of EUR 200, EUR 70 and EUR 30 and cleaning expenses of EUR 40, EUR 30 and zero. Confirm closing cash of EUR 1,100, EUR 1,080 and EUR 1,050. Confirm the new `F8` equals the prior `L34` and historical amounts remain unchanged.
- [ ] Cancel at payments and expenses; no new monthly sheets should remain.
- [ ] PDF export, correct month in filename, legibility and print layout; save without PDF.
- [ ] Delete latest month: cancel, confirm, missing PDF, mismatched workbooks, protect sole sheet, recreate month.
- [ ] Confirm failure handling with closed/open documents and temporary save errors using disposable copies. Modify a copy in Excel while the application is entering a month, then confirm that the application rejects the stale save.
- [ ] Use the blank `82e289b` regression fixture with one fictional owner and a EUR 60 configured fee. Issue the first month without payment. Verify zero prior debt, EUR 60 in `L16`, and EUR 60 outstanding.
- [ ] In an issued copy, change a EUR 100 expense to EUR 120 while leaving the two EUR 50 shares or ledger amount unchanged. Confirm the error names the month, cells and EUR 20 difference. Verify a confirmed deletion can still recover the latest period.
- [ ] Test the supported weight precision in the standalone template and the application. Two areas of 1 and 1.000000000001 are outside the shared precision contract. Check equal weights, EUR 0.01, fractional and zero heating, and a large in-range expense.
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
