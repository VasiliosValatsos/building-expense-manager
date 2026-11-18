# Blank Excel templates

Required filenames:

- `ΚΟΙΝΟΧΡΗΣΤΑ.xlsx`
- `ΒΙΒΛΙΟ ΕΣΟΔΩΝ_ΕΞΟΔΩΝ.xlsx`

Use the supplied matching blank templates. Do not rename monthly sheets, rearrange cells or remove formulas. Copy both files to a private working directory before running the application. Never replace these publicly tracked files with working accounting data.

In the common expenses template, enter an owner and positive area for each participating apartment. A blank heating factor has full weight and zero has none. The standalone calculation distributes each expense to the cent. In the ledger, `F8` is opening cash, `K16:L16` is reserved for the management fee and `L34` is closing cash. The application fills these cells when it creates the next month. Keep reserved calculation columns and the two month histories intact.

Areas support up to five decimal places and 999,999 m². Heating factors support up to four decimals from 0 to 1. Inputs outside this contract must be corrected explicitly before issuing a new month. Excel and Python convert these values to integer weights, divide by their common divisor, then rank integer remainders. Equal remainders are resolved by apartment row order. The sum of reduced weights and the product of expense cents by each reduced weight must fit within 999,999,999,999,999 for a positive expense. This bound keeps the shared Excel arithmetic exact. An unsupported input can surface as `#N/A` in standalone Excel.
