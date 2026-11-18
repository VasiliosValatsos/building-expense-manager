# User guide

## Before starting

Follow the installation steps in the README. Copy both blank templates to a private folder and keep separate backups. Close the working Excel files and PDF before opening the app.

## Initial setup and property details

Choose Greek or English and select your working copies. Enter the first reporting month, the building address and the opening cash balance. Enter an owner name and a positive area for at least one apartment. Leave the tenant field blank if needed. You can use a decimal comma or point. Enter money to at most two decimal places. Keep the workbook layout intact. Use **Change workbooks** to switch to another pair.

After a month has been generated, changes to properties and the management fee apply to the next month. An apartment with a remaining debt or credit cannot be removed or assigned to a different owner until both balances are explicitly reconciled. Record the resolution in your accounting records before proceeding. Keep separate working folders for different buildings.

Changing a tenant name while there is a tenant debt or credit requires confirmation. The balance stays with the apartment and appears on the row with the new name. This entry does not transfer legal liability.

## Heating

Enter `0` for exemption, `0.25` for 25% weighting, `0.5` for half weighting and `1` or blank for full weighting. The final charge is allocated according to relative weighted area. A positive heating bill requires at least one participating property.

## Monthly operation

Choose **Start next month**. Check the apartment number, owner, tenant and net balance. Enter the amount you received or choose full settlement. The app first offsets credits against debts for the same apartment. It applies payments to the tenant's debt first, then to the owner's. Any overpayment becomes a tenant credit. Full settlement collects the net amount due and leaves any remaining credit in place.

Enter the expenses, choose **Finish without PDF** or **Finish and export PDF**, then confirm the save. The app allocates each expense to the cent. If two apartments have the same remainder, the earlier row gets the spare cent. It charges the management fee to owners and records it once as a cash expense. The ledger opens with the previous closing balance. It adds actual receipts and subtracts actual expenses to calculate the new balance. Check both the amounts billed and the amounts received.

An older report can contain fractions of a cent. If rounding its tenant and owner balances separately changes the displayed combined total, the application shows both totals and asks before proceeding. The previous worksheet keeps its original values and formulas.

Reconcile the legacy closing balance in `L34` with actual cash before upgrading. Internal worksheet totals can agree even if an older release carried forward an incorrect `F8`. On an unissued opening sheet, actual debts and credits come from `M:N`. The configured fee in `D8` applies when the first month is issued. Review any ambiguous header or cash activity before proceeding.

If an issued legacy month shows a fee in `D8` but omits the corresponding outflow in `L16`, check whether the fee was paid. The application stops so the manager can reconcile the original records.

Before correcting an issued period, back up both workbooks and the PDF. A manual edit must leave the expense, allocation, apartment charges and corresponding ledger movement in agreement. Keep the reserved calculation cells intact. If the latest month needs a new allocation, confirm deletion and recreate it with actual receipts and corrected expenses. For an older period, reconcile its effect on later carry-forwards before changing the records.

## Delete the latest month

Back up both workbooks and the PDF, then close them. Choose **Delete last month** and check the month in the confirmation dialog. The app removes that month's sheets from both workbooks and deletes the matching PDF if it exists. It keeps the sole remaining opening sheet.

## Troubleshooting

- Missing module: install `requirements.txt` with the same Python interpreter used to launch the app.
- Excel automation or PDF error: check that desktop Excel opens locally, that the documents are closed and that required macOS automation permissions are granted.
- Invalid template: use the matching blank workbooks, with the original layout and month names.
- Incorrect file selection: check remembered local settings before testing on another dataset. Do not run the program against live files by accident.
- If an operation fails, inspect both files and do not repeat a payment blindly.
- After an interruption, close Excel and check whether another copy of the app is running. Keep the original workbooks, temporary files, backups and PDF. The transaction record lists their paths and the SHA‑256 hashes from before and after preparation. To roll back, match the backup to the original path and its before hash. Check the temporary PDF too. If several files could be the right copy, keep them all until you reconcile the pair. Check your choice before replacing files or removing locks and transaction markers. An interruption can leave one XLSX updated before the other.
- If the application reports that financial records were saved, keep the existing entries. To regenerate a PDF, close the application, open the saved month in Excel, set the print area to `A1:P31`, and export to a new filename. Review the result before replacing an existing report.
