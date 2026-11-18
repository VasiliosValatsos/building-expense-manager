# User guide

## Before starting

Install Python, the requirements and Microsoft Excel desktop as described in the main README. Copy the blank templates to a private folder. Close the working Excel files and PDF before launching the application. Keep independent backups.

## Initial setup and property details

Choose Greek or English and select private copies of both templates. Enter the first reporting month, building address and opening cash balance. At least one apartment needs a positive area and an owner name. A tenant name is optional. You can use a comma or a point for decimals. Monetary amounts accept up to two decimal places. Keep the workbook layout and use **Change workbooks** to select a different pair.

After a month has been generated, changes to properties and the management fee apply to the next month. An apartment with a remaining debt or credit cannot be removed or assigned to a different owner until both balances are explicitly reconciled. Record the resolution in your accounting records before proceeding. Keep separate working folders for different buildings.

Changing a tenant name while there is a tenant debt or credit requires confirmation. The balance stays with the apartment and appears on the row with the new name. This entry does not transfer legal liability.

## Heating

Enter `0` for exemption, `0.25` for 25% weighting, `0.5` for half weighting and `1` or blank for full weighting. The final charge is allocated according to relative weighted area. A positive heating bill requires at least one participating property.

## Monthly operation

Choose Start Program. Review the apartment number, owner, tenant and net balance. Enter the actual cash received or choose full settlement. Existing credits offset debts within the same apartment before cash is applied. Cash pays tenant debts first, then owner debts. Overpayments become tenant credits. Full settlement collects only positive net debt and preserves remaining credit.

Enter actual expenses and confirm the month. Each expense is allocated to eligible apartments to the cent, with ties resolved in apartment row order. The monthly management fee is an owner charge and is paid from cash once. The ledger's opening cash comes from the previous closing balance. Its new closing balance is opening cash plus actual receipts minus actual expenses. Review the charges and ledger separately, since amounts billed can differ from amounts collected. Choose PDF export if needed.

An older report can contain fractions of a cent. If rounding its tenant and owner balances separately changes the displayed combined total, the application shows both totals and asks before proceeding. The previous worksheet keeps its original values and formulas.

Reconcile the legacy closing balance in `L34` with actual cash before upgrading. Internal worksheet totals can agree even if an older release carried forward an incorrect `F8`. On an unissued opening sheet, actual debts and credits come from `M:N`. The configured fee in `D8` applies when the first month is issued. Review any ambiguous header or cash activity before proceeding.

If an issued legacy month shows a fee in `D8` but omits the corresponding outflow in `L16`, check whether the fee was paid. The application stops so the manager can reconcile the original records.

Before correcting an issued period, back up both workbooks and the PDF. A manual edit must leave the expense, allocation, apartment charges and corresponding ledger movement in agreement. Keep the reserved calculation cells intact. If the latest month needs a new allocation, confirm deletion and recreate it with actual receipts and corrected expenses. For an older period, reconcile its effect on later carry-forwards before changing the records.

## Delete the latest month

Close the two Excel files and PDF. Choose Delete last month and verify the displayed month carefully. Confirm only if deletion is intended. The corresponding sheets and matching report are removed, subject to the application's safety checks. A missing PDF does not prevent removal of the matching sheets; the last remaining base sheet is protected. Back up financial records independently before destructive actions.

## Troubleshooting

- Missing module: install `requirements.txt` with the same Python interpreter used to launch the app.
- Excel automation or PDF error: check that desktop Excel opens locally, that the documents are closed and that required macOS automation permissions are granted.
- Invalid template: use the matching blank workbooks, with the original layout and month names.
- Incorrect file selection: check remembered local settings before testing on another dataset. Do not run the program against live files by accident.
- If an operation fails, inspect both files and do not repeat a payment blindly.
- After an interruption, close Excel and check that no second application instance is active. Preserve copies of the originals, temporary files, backups and PDF. The transaction record maps original, staged and backup paths and lists before and prepared SHA‑256 hashes. For rollback, confirm the backup matches the original's recorded before hash and review the staged PDF. Preserve multiple or unidentified candidates until the pair can be reconciled. Confirm the recovery choice before replacing files or removing locks and transaction markers. Two independently stored XLSX files cannot be replaced as one atomic operation.
- If the application reports that financial records were saved, keep the existing entries. To regenerate a PDF, close the application, open the saved month in Excel, set the print area to `A1:P31`, and export to a new filename. Review the result before replacing an existing report.
