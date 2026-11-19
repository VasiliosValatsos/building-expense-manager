# Demo month and workbooks

[February 2026 report (PDF)](sample-report.pdf) · [Both Excel workbooks (ZIP)](demo-workbooks.zip)

The application generated these files with desktop Excel 16.0 on Windows. The building and all names are fictional. The PDF header reads “DEMO - Fictional building”. Extract the workbooks to a separate folder and open them in desktop Excel to recalculate the formulas. Use fresh copies of the blank templates if you want to reproduce the run.

## Inputs

Set the first reporting month to January 2026, the starting cash to €600 and the monthly management fee to €45. Use the building name `DEMO - Fictional building`. Enter `Demo Owner 1` to `Demo Owner 3` and `Demo Tenant 1` to `Demo Tenant 3` in the first three rows. Leave the other rows blank.

| Apartment | Floor | Area | Heating factor |
| --- | --- | ---: | ---: |
| 1 | Ground | 50 m² | 1 |
| 2 | First floor | 75 m² | 0.5 |
| 3 | Second floor | 100 m² | 0 |

In January, enter €50 from apartment 3 and leave the other payments blank. Enter €90 for common-area cleaning and leave all other expenses blank. The €45 management fee is already set. January closes at **€515** (€600 + €50 − €135).

## February payments

| Apartment | Tenant balance from January | Owner balance from January | February input | Cash received | Prior debt left |
| --- | ---: | ---: | --- | ---: | ---: |
| 1 | €20 | €10 | Enter €25 | €25 | €5 |
| 2 | €30 | €15 | Choose **Full settlement** | €45 | €0 |
| 3 | −€10 | €20 | Leave payment blank | €0 | €10 |

Apartment 3's €10 tenant credit offsets €10 of owner debt. The cash ledger records no receipt for that offset.

## February expenses and balances

Enter €70 for natural gas and €100 for common-area cleaning. Leave the other fields blank. The €45 management fee makes total expenses **€215**. Choose **Finish and export PDF** and confirm the save.

| Apartment | Cleaning | Heating | Management fee | New charges | Due after February |
| --- | ---: | ---: | ---: | ---: | ---: |
| 1 | €22.22 | €40.00 | €10.00 | €72.22 | €77.22 |
| 2 | €33.33 | €30.00 | €15.00 | €78.33 | €78.33 |
| 3 | €44.45 | €0.00 | €20.00 | €64.45 | €74.45 |
| Total | €100.00 | €70.00 | €45.00 | €215.00 | €230.00 |

Floor area sets the cleaning and management shares. The €100 cleaning bill leaves one cent after rounding the shares down, and apartment 3 has the largest remainder. Heating uses weights of 50, 37.5 and 0.

The cash ledger opens at **€515** (`F8`). Receipts of **€70** bring the income-side total to **€585** (`F31`, which includes opening cash). Outflows are **€215** (`L31`), including the **€45** fee in `L16`. Closing cash is therefore **€370** (`L34`). The **€230** due from apartments is a separate total.

## Το παράδειγμα στα ελληνικά

Τρεις ιδιοκτησίες 50, 75 και 100 τ.μ. έχουν συντελεστές θέρμανσης 1, 0,5 και 0. Ο Φεβρουάριος ξεκινά με ταμείο €515. Οι εισπράξεις των €70 και οι εκροές των €215 δίνουν κλείσιμο **€370**. Στην τρίτη ιδιοκτησία, η πίστωση €10 του ενοικιαστή μειώνει την οφειλή του ιδιοκτήτη χωρίς είσπραξη. Οι νέες χρεώσεις είναι €215 και οι συνολικές οφειλές μετά τη μεταφορά είναι €230.
