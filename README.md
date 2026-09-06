# 06.09.2026

MSTC combined bid sheet (auctions 21977, 21978, 21979, 21980).

## Workbook contents

| Sheet | Orientation | What it is |
| --- | --- | --- |
| `Final Calculation Sheet` | — | The original working sheet. Untouched — the single source of truth. |
| `Collapsible - Lots Across` | fields ↓ rows, lots → columns | Colour block per buyer, its lots side by side, every field down the rows. Folds by buyer and by section. Buyer index with jump links at the bottom. |
| `Collapsible - Lots Down` | lots ↓ rows, fields → columns | Colour block per buyer, one row per lot, `∑ TOTALS` row per buyer. Folds by buyer and by column category. |
| `Ledger Filter - Lots Down` | lots ↓ rows, fields → columns | One flat table, AutoFilter on all 35 columns, lots grouped under filter aware `SUBTOTAL` buyer rows, KPI band on top. |
| `Ledger Filter - Lots Across` | fields ↓ rows, lots → columns | One flat table, all 37 lots as columns, AutoFilter on the FIELD column, colour coded buyer band across the top. |

Every figure on the four views is a **live formula** pointing at `Final Calculation Sheet` —
nothing is typed in, so they all update the moment the source sheet changes.

## Folding and filtering

* `−` / `+` in the grey margin on the left folds rows; `−` / `+` above the column letters
  folds columns. The `1 2 3` buttons at the top left expand everything again.
* `Ledger Filter - Lots Down` — filter arrows in row 5. The buyer summary rows, the KPI band
  in row 3 and the `∑ GRAND TOTAL` row all use `SUBTOTAL`, so they only add up the rows the
  filter leaves visible.
* `Ledger Filter - Lots Across` — filter arrow on the FIELD column picks which lines to show;
  the `−` above the `TOTAL` column folds all 37 lot columns away.
* Conditional formatting on all four: outstanding in red, settled in green, lots with no
  invoice number yet in amber, plus a data bar on Mat. Value in the lots-down ledger.

## Regenerating the views

```bash
pip install openpyxl
python3 tools/build_buyer_lot_views.py 06.09.2026.xlsx   # rebuild all four sheets in place
python3 tools/verify_views.py 06.09.2026.xlsx            # needs: pip install formulas
```

`verify_views.py` recalculates the workbook with a formula engine and checks every cell of
all four views against `Final Calculation Sheet` — 6,129 checks.
