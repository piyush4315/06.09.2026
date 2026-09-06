# 06.09.2026

MSTC combined bid sheet (auctions 21977, 21978, 21979, 21980).

## Workbook contents

| Sheet | What it is |
| --- | --- |
| `Final Calculation Sheet` | The original working sheet. Untouched — it is the single source of truth. |
| `Buyer Collapsible View` | Colour coded block per buyer, its lots side by side as columns, every field down the rows. Two fold levels: a whole buyer, or a single section. |
| `Lot Ledger (Filter)` | One row per lot with an Excel AutoFilter on every column, lots grouped under a filter aware buyer summary row. |

Both new sheets are **live links** to `Final Calculation Sheet` — nothing is typed in, so
they update the moment the source sheet changes.

### Buyer Collapsible View

* `−` / `+` in the left margin: outer level folds a whole buyer, inner level folds one
  section (LOT INFORMATION, FINANCIALS, SECURITY DEPOSIT, FINAL PAYMENT, LPP, SUMMARY,
  DOCUMENT). The `1 2 3` buttons at the top left expand everything again.
* `−` / `+` above the `BUYER TOTAL` column folds all lot columns away, leaving one total
  column per buyer.
* Every buyer has its own colour; the buyer banner shows lot count, mat. value, received
  and outstanding, recalculated from the source sheet.
* A **Buyer Index** at the bottom lists all buyers with their totals and a click-to-jump
  link to each block.

### Lot Ledger (Filter)

* Filter arrows in row 5 — buyer, lot no., bid sheet, unit, payment status, and so on.
* `−` in the left margin folds a buyer's lots into its coloured summary row; those summary
  rows use `SUBTOTAL`, so they only add up the rows the filter leaves visible.
* `−` above a category's first column folds that block of columns away (Lot Name,
  Mat. Value, SD / FP / LPP Expected, Total Received and Invoice No. stay on screen).
* KPI band in row 3 and the GRAND TOTAL row follow the filter as well.
* Conditional formatting: outstanding amounts in red, settled lots in green, lots with no
  invoice number yet highlighted in amber, data bars on Mat. Value.

## Regenerating the views

The two sheets are built by a script, so they can be rebuilt after lots are added or
removed:

```bash
pip install openpyxl
python3 tools/build_buyer_lot_views.py 06.09.2026.xlsx   # rebuild both sheets in place
python3 tools/verify_views.py 06.09.2026.xlsx            # needs: pip install formulas
```

`verify_views.py` recalculates the workbook with a formula engine and checks every cell of
both views against `Final Calculation Sheet` (3,380 checks).
