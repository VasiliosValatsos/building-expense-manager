# Changelog

## 1.0.0 development milestones

- Added monthly expense allocation and updates to both Excel workbooks.
- Accepted decimal commas and points. Used desktop Excel to read recalculated debts.
- Showed owners and tenants during payment entry, validated inputs and added recovery for workbook saves.
- Required confirmation to delete the latest month and its PDF.
- Used a heating factor of zero for exemption and a blank field for full weight.
- Split owner and tenant expenses, updated the blank templates and kept their conditional formatting.
- Combined the Greek and English interfaces in one app.
- Added documentation, regression tests and a GitHub Actions workflow.
- The developer reported manual checks of an earlier version on Windows 10 and macOS.
- Added cent-level allocations, explicit credit handling, closing cash carry-forward and a ledger entry for the management fee.
- Added reconciliation before saving, recovery records and checks for changes to open workbooks.
- Separated the Python routines that build Excel formulas, write reports and save a month. Compared the resulting workbooks and PDF text with the previous checkpoint.
