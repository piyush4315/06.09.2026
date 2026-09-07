# 06.09.2026

MSTC combined bid sheet (auctions 21977, 21978, 21979, 21980).

## Workbook contents

| Sheet | Layout | What it is |
| --- | --- | --- |
| `Final Calculation Sheet` | — | The original working sheet. Untouched — the single source of truth. |
| `Buyer Pivot` | **buyers → columns**, details ↓ rows | One column per buyer, 30 detail rows down the side, `TOTAL — ALL BUYERS` at the end. Every cell a live SUMIF/COUNTIF on the buyer name in row 3. |
| `Lot-wise Vertical` | **everything downwards** | Buyer ▸ lot ▸ field, one below the other. Three fold levels: buyer → lot → field. Each buyer ends with its own `∑ TOTALS` block, the sheet ends with an all-buyers grand total. |
| `Collapsible - Lots Across` | fields ↓ rows, lots → columns | Colour block per buyer, its lots side by side, every field down the rows. Folds by buyer and by section. Buyer index with jump links at the bottom. |
| `Collapsible - Lots Down` | lots ↓ rows, fields → columns | Colour block per buyer, one row per lot, `∑ TOTALS` row per buyer. Folds by buyer and by column category. |
| `Ledger Filter - Lots Down` | lots ↓ rows, fields → columns | One flat table, AutoFilter on all 35 columns, lots grouped under filter aware `SUBTOTAL` buyer rows, KPI band on top. |
| `Ledger Filter - Lots Across` | fields ↓ rows, lots → columns | One flat table, all 37 lots as columns, AutoFilter on the FIELD column, colour coded buyer band across the top. |

Every figure on the five views is a **live formula** pointing at `Final Calculation Sheet` —
nothing is typed in, so they all update the moment the source sheet changes.

## Buyer Pivot — buyers across, details down

| FIELD ▸ / BUYER → | AL HAMD | F R KHANS | … | WATAN | TOTAL — ALL BUYERS |
| --- | --- | --- | --- | --- | --- |
| Lots | 1 | 2 | … | 3 | 37 |
| Mat. Value ₹ | 790,089 | 238,768 | … | 609,100 | 15,130,598 |
| Total Received ₹ | … | … | … | … | … |
| Outstanding ₹ | … | … | … | … | … |
| Payment Status | … | … | … | … | … |
| Quantity (total) / Avg Rate ₹ / Lots pending invoice | | | | | |
| 11 financial lines incl. GST, TCS, both TDS, service charge, receivables | | | | | |
| SD Expected / Received / Outstanding | | | | | |
| FP Expected / Received / Outstanding | | | | | |
| LPP Expected / Received | | | | | |
| Invoices raised / SAP Docs posted | | | | | |
| Collection % | | | | | |

* Each buyer column is driven by the buyer name in **row 3** — rename that cell and the whole
  column re-points at the new buyer.
* `−` / `+` in the left margin folds a whole detail band (OVERVIEW, LOT INFORMATION,
  FINANCIALS, SECURITY DEPOSIT, FINAL PAYMENT, LPP, DOCUMENT, RECOVERY).
* Filter arrow on the FIELD column; conditional formatting on Outstanding (red/green),
  Lots pending invoice (amber) and a red-amber-green colour scale on Collection %.

## Lot-wise Vertical — the fully vertical sheet

Reads top to bottom only, two columns (`FIELD` and `VALUE`):

```
▼  NATIONAL ENTERPRISES  •  13 LOT(S)  •  MAT. VALUE ₹…  •  OUTSTANDING ₹…      ← fold level 1
   ▸  LOT 1874  •  Scrap of Empty oil drum  •  MAT. VALUE ₹…  •  OUTSTANDING ₹…  ← fold level 2
      ▸ LOT INFORMATION                                                          ← fold level 3
            Lot No.                    1874
            Lot Name                   Scrap of Empty oil drum
            Bid Sheet / Auction       21977
            …
   ∑  NATIONAL ENTERPRISES — TOTALS (all its lots)
∑∑  ALL BUYERS — GRAND TOTAL
```

The `1 2 3 4` buttons at the top left set the depth: 1 = buyer names only, 2 = buyers + their
lot numbers, 3 = + section headings, 4 = every field of every lot.

## Folding and filtering

* `−` / `+` in the grey margin on the left folds rows; `−` / `+` above the column letters
  folds columns.
* `Ledger Filter - Lots Down` — filter arrows in row 5. The buyer summary rows, the KPI band
  in row 3 and the `∑ GRAND TOTAL` row all use `SUBTOTAL`, so they only add up the rows the
  filter leaves visible.
* `Ledger Filter - Lots Across` — filter arrow on the FIELD column picks which lines to show;
  the `−` above the `TOTAL` column folds all 37 lot columns away.
* Conditional formatting on the wide sheets: outstanding in red, settled in green, lots with
  no invoice number yet in amber, plus a data bar on Mat. Value in the lots-down ledger.

## Regenerating the views

```bash
pip install openpyxl
python3 tools/build_buyer_lot_views.py 06.09.2026.xlsx   # rebuild all six sheets in place
python3 tools/verify_views.py 06.09.2026.xlsx            # needs: pip install formulas
```

`verify_views.py` recalculates the workbook with a formula engine and checks every cell of
all six views against `Final Calculation Sheet` — 8,102 checks.
