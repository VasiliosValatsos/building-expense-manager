# User guide

## Before starting

Install Python, the requirements and Microsoft Excel desktop as described in the main README. Copy the blank templates to a private folder. Close the working Excel files and PDF before launching the application. Keep independent backups.

## Initial setup and property details

Choose Greek or English, select both workbooks when prompted and enter the requested starting month, building address and opening balance. Enter the property floor, area, tenant and owner in the appropriate slot. Enter decimals with a comma or a point. Do not change the workbook layout. Confirm the property settings. Edits to the latest sheet can affect its existing calculated charges.

## Heating

Enter `0` for exemption, `0.25` for 25% weighting, `0.5` for half weighting and `1` or blank for full weighting. The final charge is allocated according to relative weighted area. A positive heating bill requires at least one participating property.

## Monthly operation

Choose Start Program. Review the owner, tenant and outstanding balance per property. Enter the actual payment or choose full payment, not both. Enter recurring amounts and descriptions for additional expenses: urgent expenses are allocated to owners; extra expenses to tenants. Review and confirm the final action. Choose PDF export if needed. Check both workbooks and compare their totals and the report to actual receipts.

## Delete the latest month

Close the two Excel files and PDF. Choose Delete last month and verify the displayed month carefully. Confirm only if deletion is intended. The corresponding sheets and matching report are removed, subject to the application's safety checks. A missing PDF does not prevent removal of the matching sheets; the last remaining base sheet is protected. Back up financial records independently before destructive actions.

## Troubleshooting

- Missing module: install `requirements.txt` with the same Python interpreter used to launch the app.
- Excel automation or PDF error: check that desktop Excel opens locally, that the documents are closed and that required macOS automation permissions are granted.
- Invalid template: use the matching blank workbooks, with the original layout and month names.
- Incorrect file selection: check remembered local settings before testing on another dataset. Do not run the program against live files by accident.
- If an operation fails, inspect both files and do not repeat a payment blindly.
