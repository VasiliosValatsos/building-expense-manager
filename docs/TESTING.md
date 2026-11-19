# Testing

## Automated tests

The regression suite checks payments and settlement, cent-level allocations, cash carry-forward, the management fee, changes to apartment records, template layout and recovery after a failed save. Tests that read calculated cells use a simulated Excel reader.

Run `python -m unittest discover -s tests -v` from the repository root with the dependencies installed. On macOS, use `python3` if needed. The CI workflow is configured for Python 3.11 and 3.12 on Windows and macOS. It runs without desktop Excel or the GUI.

## Manual checks with working copies

- [ ] Open the app in each language and check the labels and dialogs.
- [ ] Test initial setup and all 15 slots. Try both decimal separators, cancel an edit, and check that the app rejects public templates as working files.
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

Run `python --version` and `python -m pip show openpyxl xlwings python-dateutil` to record the versions used for a test. The install requirements specify version ranges.

## Windows test environment

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

All 47 regression tests passed in this virtual environment. A separate check used desktop Excel 16.0 with simulated form inputs to compare two monthly workbooks and the PDF text with the previous checkpoint. The interactive GUI checklist, the current macOS run and the current CI matrix remain to be checked. The developer reported manual macOS checks on an earlier version.

For the portfolio example, the same Windows setup ran the application's monthly workflow with fictional inputs for two consecutive months. Desktop Excel recalculated both books and exported the February PDF. The checked values include the apartment allocations, receipts, management fee and closing balance. The [demo walkthrough](../examples/sample-output/README.md) lists the inputs and expected results. The automated CI matrix still runs without Excel or the GUI.
