# 06.09.2026

MSTC combined bid sheet (auctions 21977, 21978, 21979, 21980).

## Workbook contents

| Sheet | Layout | What it is |
| --- | --- | --- |
| `Final Calculation Sheet` | — | The original working sheet. Untouched — the single source of truth. |
| `Live Search` | lots ↓ rows, fields → columns | **One search box (B3).** Type a lot number, part of one, a buyer or a lot name and the matching lots light up green while the rest fade to grey — live, as you type, no macros, and **no row ever moves or hides**. F3 narrows what the text is looked for in, H3 counts the hits, the `∑` row totals the hits only. The workbook opens here. |
| `Final Calc (Transposed)` | **field labels ↓ rows**, values → columns | A one-for-one mirror of `Final Calculation Sheet` turned on its side: the 33 captions of row 3 down column A in the same order, one column per lot, and the source's own total row as the last column. Header colour = auction. **Filter button on the row-label column** (`A3:A36`) — tick the fields you want and their values stay, the rest hide. |
| `Transposed + Field Filter` | field labels ↓ rows, values → columns | Same as `Final Calc (Transposed)` (label-only filter, no arrows on the 37 lot headers), kept as a separate tab with its own help line. |
| `Transposed + Lot Folds` | field labels ↓ rows, values → columns | Same mirror with the lot columns **sorted by auction then buyer** and grouped: a `−` above a thin divider folds one buyer's lots, a `−` above a wide divider folds a whole auction. Auction and buyer bands above the headers. |
| `Transposed + Both Filters` | field labels ↓ rows, values → columns | The two together: the label-column filter button *and* the buyer / auction fold buttons. |
| `Transposed - Pick a Buyer` | field labels ↓ rows, values → columns | Same transposed layout, but the lot columns follow a **buyer dropdown in B3**: pick a buyer and only its lots are shown (up to 13 columns), every figure re-pointed by INDEX/MATCH. No macros. |
| `Transposed - Filter Values` | field labels ↓ rows, values → columns | Same grid, plus a **query panel in row 3**: pick a buyer, pick a ROW (any of the 33 fields), a test (`>`, `contains`, `is blank` …) and a value — only the lots that pass fill the columns. This is the sheet that filters by a row's values. |
| `Buyer Pivot` | **buyers → columns**, details ↓ rows | One column per buyer, 30 detail rows down the side, `TOTAL — ALL BUYERS` at the end. Every cell a live SUMIF/COUNTIF on the buyer name in row 3. |
| `Buyer Rows (Filter)` | **buyers ↓ rows**, details → columns | One row per buyer, 30 detail columns across, a `▼` on **every** column — including `BUYER`, so ticking buyers hides the rest. The `∑ TOTAL` row uses SUBTOTAL, so it totals only the buyers left visible. |
| `Lot-wise Vertical` | **everything downwards** | Buyer ▸ lot ▸ field, one below the other. Three fold levels: buyer → lot → field. Each buyer ends with its own `∑ TOTALS` block, the sheet ends with an all-buyers grand total. |
| `Collapsible - Lots Across` | fields ↓ rows, lots → columns | Colour block per buyer, its lots side by side, every field down the rows. Folds by buyer and by section. Buyer index with jump links at the bottom. |
| `Collapsible + Field Filter` | fields ↓ rows, lots → columns | Identical to `Collapsible - Lots Across` plus **one filter button on the label column**, so a buyer's block can be reduced to, say, only its Outstanding and Total Received rows. |
| `Collapsible - Lots Down` | lots ↓ rows, fields → columns | Colour block per buyer, one row per lot, `∑ TOTALS` row per buyer. Folds by buyer and by column category. |
| `Ledger Filter - Lots Down` | lots ↓ rows, fields → columns | One flat table, AutoFilter on all 35 columns, lots grouped under filter aware `SUBTOTAL` buyer rows, KPI band on top. |
| `Ledger Filter - Lots Across` | fields ↓ rows, lots → columns | One flat table, all 37 lots as columns, AutoFilter on the FIELD column, colour coded buyer band across the top. |

Every figure on the fifteen views is a **live formula** pointing at `Final Calculation Sheet` —
nothing is typed in, so they all update the moment the source sheet changes.

## Live Search — type, and the misses fade out

| | A3 | B3 | | | E3 | F3 | G3 | H3 | I3 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| | 🔍 SEARCH ▸ | **187** | | | SEARCH IN ▸ | **Lot No. + Buyer + Name** ▾ | MATCHES ▸ | **3 of 37** | matching 3 of 37 lots • search: 187 • in: Lot No. + Buyer + Name • the rest are greyed out, not hidden |
| | # | Lot No. | Buyer | Lot Name | … | | | | |
| | 1 | *1763* | *AL HAMD TRADE CORPORATION* | *SCRAP COPPER OF T/F WINDING …* | ← greyed: still on screen, just not matching | | | | |
| | 2 | **1874** | **NATIONAL ENTERPRISES** | **Scrap of Empty oil drum** | ← matching | | | | |
| | 3 | **1875** | **F R KHANS ENTERPRISES** | **Scrap of Copper of faulty AC compressor** | | | | | |
| | 4 | **1876** | **F R KHANS ENTERPRISES** | **Scrap of MS (extracted from equip cage…** | | | | | |
| | 5 | *1923* | *OMKAR STEELS* | *SCRAP ACSR CONDUCTORS …* | ← greyed again | | | | |
| | ∑ | | | | … totals of the 3 matching lots only | | | | |

* **All 37 rows stay where they are.** Row *n* of the table is always lot *n* of
  `Final Calculation Sheet`, so nothing jumps around while you type and the `#` column is the
  lot's own position, 1..37. A matching lot gets a light green wash; the others keep their
  place but fade — grey text on a grey fill — so the whole list is still readable and you can
  still see what did *not* match.
* Matching is **contains**, not case sensitive, so `187` finds 1874 / 1875 / 1876, `NATIONAL`
  finds all 14 lots of NATIONAL ENTERPRISES and NATIONAL SCRAP AND BUILDING MATERIAL
  SUPPLIER, and `copper` finds the two lots with COPPER in the name. Clear the box and every
  row goes back to full colour.
* **H3** counts the hits (`3 of 37`) and the hint in **I3** echoes what you typed and where it
  is being looked for.
* **F3** narrows the search to one field: `Lot No.`, `Buyer`, `Lot Name`, `Bid Sheet` or
  `Unit` (default searches lot no. + buyer + name together).
* The `∑` row at the bottom totals **only the matching lots** (`SUMIF` over the 1/0 flags), so
  with `187` typed it totals those three lots' material value and outstanding, and it reads
  blank when nothing matches. `Payment Status` / `Outstanding` keep their colours on the hits.
* No macro and nothing typed in: every visible cell is a plain link to the source row of the
  same number. Hidden columns AK..AO do the work — AK builds the text each lot is searched in
  from the field chosen in F3, AM holds the 1/0 match flag that both the formatting and the
  totals read, and AO is the drop-down list behind F3.

## Final Calc (Transposed) — the source sheet on its side

Exactly the same fields in exactly the same order, nothing added or dropped:

| FIELD (row 3 of the source) ▸ | 1763 | 1874 | … | 2091 | TOTAL (src row 41) |
| --- | --- | --- | --- | --- | --- |
| Quantity | 669 | 270 | … | 3 | |
| Lot Name | SCRAP COPPER OF T/F WINDING… | Scrap of Empty oil drum | … | | |
| Rate / Bid Sheet / Unit / Lot No. / Buyer | | | | | |
| Mat. Value … Doc./Invoice Date | | | | | 15,130,598 … |

* Every cell is `=IF('Final Calculation Sheet'!$H$4="","",'Final Calculation Sheet'!$H$4)` —
  the two sheets can never disagree, and blanks stay blank instead of turning into zeros.
* The source caption `Date of Receipt` heads two columns (SD and FP), so those two rows are
  labelled `Date of Receipt (col V)` and `Date of Receipt (col Y)`.
* Number formats are copied from the source, column A is frozen, and the `▼` on the FIELD
  column filters which rows show.

## Filter buttons on the row labels

An AutoFilter range puts a `▼` on **every** cell of its header row, so a wide range like
`A3:AM36` grows 39 arrows — one on each lot number. These sheets give the labels their own
button instead, by making the filter range a **single column**:

| Sheet | Filter range | Buttons |
| --- | --- | --- |
| `Final Calc (Transposed)` | `A3:A36` | one `▼` on `FIELD (row 3 of the source) ▸`, listing all 33 field names |
| `Transposed + Field Filter` | `A3:A36` | same, on its own tab |
| `Transposed + Lot Folds` | none | `−` / `+` fold buttons per buyer and per auction |
| `Transposed + Both Filters` | `A5:A38` | the label `▼` **and** the fold buttons |
| `Collapsible + Field Filter` | `A5:A563` | one `▼` on the `FIELD ▸ BUYER \| LOT →` header, stopping above the buyer index so filtering never hides the jump table |

Tick a few fields — say `Outstanding`, `Total Received`, `Payment Status` — and the other rows
hide; the 37 lot columns stay exactly where they are. The dropdown reads the row names down
column A, so you filter the values *against the row name*: pick `Outstanding` and every lot's
outstanding figure lines up across the sheet.

### Filtering by buyer

An Excel `▼` can only hide **rows**, never columns — so on the transposed sheets (lots as
columns) a dropdown on the `Buyer` row cannot hide other buyers' lot columns. Two sheets solve
it, one each way:

* **`Transposed - Pick a Buyer`** keeps the transposed look and replaces the column filter with
  a **dropdown in B3**. A hidden helper column (Q) numbers each lot of the chosen buyer 1..n,
  a second (S) turns that into a source row, and every lot column is an
  `INDEX(... , $S<row> - 3)`. Pick `NATIONAL ENTERPRISES` and its 13 lots fill B..N; pick
  `SAHARA ENTERPRISES` and only B is used. The `TOTAL` column then totals just that buyer, and
  the `▼` on the FIELD column still filters which rows show.
* **`Buyer Rows (Filter)`** turns the table the other way: 13 buyer rows × 30 detail columns.
  Now `BUYER` is a real column, so its `▼` filters buyers the way Excel intends; every other
  column has an arrow too (Payment Status, Collection %, …). The `∑ TOTAL — visible buyers`
  row sits below the filter range and uses `SUBTOTAL(109,…)`, so it follows the filter.

### Filtering by a ROW's values

The arrow on the FIELD column picks *which rows show*. To pick which **lots** show on the
strength of one row's values — "the lots whose Outstanding is greater than 0", "the lots whose
Lot Name contains COPPER", "the lots with no invoice number yet" — use
**`Transposed - Filter Values`**. Its row 3 is a query panel, not a filter arrow:

| | A3 | B3 | C3 | D3 | E3 | F3 | G3 | H3 | I3 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| | FILTER ▸ | **1 buyer** ▾ | 2 ROW ▸ | **Outstanding** ▾ | 3 TEST ▸ | **>** ▾ | 4 VALUE ▸ | **0** | showing 9 of 37 lots • buyer: ALL BUYERS • Outstanding > 0 |

Only the lots that pass fill the columns, left to right; the rest stay empty. Verified
combinations: `Outstanding > 0` → 9 lots · `Total Received >= 1000000` → 4 · `Buyer =
STERLING ENTERPRISES` → 7 · `Invoice No. is blank` → 5 · buyer *and* `Outstanding > 0` → 2.
Set cell 2 back to `— no row filter —` to see all 37 again. The `TOTAL` column then totals
just the lots on screen, and the `▼` on the FIELD column still hides rows.

Hidden helper columns AO..BA do the work: AO/AQ/AS hold the three dropdown lists, AW flags
each source lot 1/0 against the panel, AY numbers the ones that pass, and BA turns that number
back into a source row — so every visible cell is still a plain `INDEX` into
`Final Calculation Sheet`.

### Lot folds (the two grouped sheets)

The lot columns are sorted **auction → buyer → lot no.** and outlined in two levels, with a
thin divider column between every two groups so the buttons never merge:

```
        1   2                                     ← outline buttons: 1 = auctions only, 2 = every buyer
AUCTION 21977            │ AUCTION 21978 …
F R KHANS │HIND│NATIONAL│SHAR JAHAN│STERLING │ …
  −    −  │ −  │   −    │    −     │    −    │
 1763 1874 │2011│1875 1876│ 1990 2059│ 1991 …│
```

* `−` above a **thin** divider folds that buyer's lots away.
* `−` above a **wide** divider folds a whole auction (all 5–7 buyers of it).
* `1` / `2` above the column letters collapse every auction / every buyer in one click.
* 37 lot columns at outline level 2, 17 buyer dividers at level 1, 3 auction dividers at
  level 0 — `Transposed + Both Filters` is the same with the label filter on top.

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
python3 tools/build_buyer_lot_views.py 06.09.2026.xlsx   # rebuild all fifteen sheets in place
python3 tools/verify_views.py 06.09.2026.xlsx            # needs: pip install formulas
```

`verify_views.py` recalculates the workbook with a formula engine and checks every cell of
all fifteen views against `Final Calculation Sheet` — 18,856 checks, including the filter
ranges, the freeze panes, the outline levels, the 21 buyer / 4 auction bands, the buyer picker
recalculated three times, the row-value filter four times and the live search four times
(empty box, `187`, `NATIONAL`, `OMKAR` restricted to Buyer). Ten recalculations, so it takes
about ten minutes.
