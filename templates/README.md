# Blank Excel templates

Required filenames:

- `ΚΟΙΝΟΧΡΗΣΤΑ.xlsx`
- `ΒΙΒΛΙΟ ΕΣΟΔΩΝ_ΕΞΟΔΩΝ.xlsx`

Copy both blank workbooks to a private folder before using the app. Keep the sheet names, cells and formulas in place. Store real records in the copies.

In the common expenses template, enter an owner and positive area for each participating apartment. A blank heating factor has full weight and zero has none. The standalone calculation distributes each expense to the cent. In the ledger, `F8` is opening cash, `K16:L16` is reserved for the management fee and `L34` is closing cash. The application fills these cells when it creates the next month. Keep reserved calculation columns and the two month histories intact.

Enter up to five decimal places for areas, with a maximum of 999,999 m². Heating factors range from 0 to 1 and accept up to four decimals. Correct out-of-range values before issuing a month. Excel and Python turn the weights into integers, divide by their common divisor and rank the integer remainders. In a tie, the earlier apartment row comes first. For a positive expense, both the sum of the reduced weights and each reduced weight multiplied by the amount in cents must be at most 999,999,999,999,999. This keeps the arithmetic exact in Excel. The standalone workbook can show `#N/A` for unsupported input.
