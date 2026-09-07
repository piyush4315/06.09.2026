#!/usr/bin/env python3
"""
Build two extra views on top of `Final Calculation Sheet` inside the MSTC bid
workbook:

  1. "Buyer Collapsible View"  - every buyer gets a colour coded block; the lots
     of that buyer run left -> right as columns and every detail field runs down
     as a row.  Two row outline levels: fold a whole buyer, or fold a single
     section (LOT INFORMATION / FINANCIALS / ...).  The lot columns are grouped
     as well, so they can be folded away leaving only the buyer totals.

  2. "Lot Ledger (Filter)"     - one row per lot with an Excel AutoFilter on
     every column.  Lots are grouped under a coloured, filter aware buyer
     summary row (SUBTOTAL), so folding a buyer keeps its totals on screen, and
     the KPI band on top always follows whatever the filter leaves visible.

Every figure written by this script is a live link back to
'Final Calculation Sheet' - nothing is hard coded, so both views follow the
source sheet.  Re-run this script after adding or removing lots.

Usage:  python3 tools/build_buyer_lot_views.py [workbook.xlsx]
"""

from __future__ import annotations

import sys
from collections import OrderedDict

from openpyxl import load_workbook
from openpyxl.formatting.rule import (CellIsRule, ColorScaleRule, DataBarRule,
                                      FormulaRule)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

SRC = "Final Calculation Sheet"
# four views, two per layout: fields-down/lots-across and lots-down/fields-across
COLLAPSIBLE = "Collapsible - Lots Across"        # fields as rows, lots as columns
COLLAPSIBLE_ROWS = "Collapsible - Lots Down"     # lots as rows, fields as columns
LEDGER = "Ledger Filter - Lots Down"             # lots as rows, fields as columns
LEDGER_COLS = "Ledger Filter - Lots Across"      # fields as rows, lots as columns
LOT_VERTICAL = "Lot-wise Vertical"               # buyer > lot > fields, all downwards
BUYER_PIVOT = "Buyer Pivot"                      # buyers across, every detail down
TRANSPOSED = "Final Calc (Transposed)"           # the source sheet, turned on its side
T_FIELD_FILTER = "Transposed + Field Filter"     # one filter button on the label column
T_LOT_FOLDS = "Transposed + Lot Folds"           # per buyer / per auction fold buttons
T_BOTH = "Transposed + Both Filters"             # the two together
C_FIELD_FILTER = "Collapsible + Field Filter"    # collapsible blocks + label filter
T_PICK_BUYER = "Transposed - Pick a Buyer"       # choose a buyer, its lots fill the columns
BUYER_ROWS = "Buyer Rows (Filter)"               # one row per buyer, filter the Buyer column
T_VALUE_FILTER = "Transposed - Filter Values"    # filter the lot columns by any row's values
ALL_BUYERS = "ALL BUYERS"
NO_FIELD = "\u2014 no row filter \u2014"
TESTS = ("=", "<>", ">", ">=", "<", "<=", "contains", "is blank", "is not blank")
# in workbook order: the sheet built LAST is the one that lands at index 1
ALL_VIEWS = (TRANSPOSED, T_FIELD_FILTER, T_LOT_FOLDS, T_BOTH, T_PICK_BUYER,
             T_VALUE_FILTER, BUYER_PIVOT, BUYER_ROWS, LOT_VERTICAL, COLLAPSIBLE,
             C_FIELD_FILTER, COLLAPSIBLE_ROWS, LEDGER, LEDGER_COLS)
# sheets from the first revision, renamed since - dropped so they do not linger
LEGACY_VIEWS = ("Buyer Collapsible View", "Lot Ledger (Filter)")
FIRST_DATA_ROW = 4          # first lot row on the source sheet
LAST_SRC_ROW = 1000         # generous tail so new lots are picked up by SUMIF

# --------------------------------------------------------------------------- #
# field catalogue: (label, source column, kind, aggregation for buyer totals)
# --------------------------------------------------------------------------- #
KIND_FMT = {
    "num0": "#,##0",
    "dec": "#,##0.00",
    "qty": "#,##0.###",
    "pct": "0%",
    "text": "@",
}

SECTIONS = [
    ("LOT INFORMATION", [
        ("Lot No.",                       "F",  "num0", None),
        ("Lot Name",                      "B",  "text", None),
        ("Bid Sheet / Auction",           "D",  "num0", None),
        ("Quantity",                      "A",  "qty",  "sum"),
        ("Unit",                          "E",  "text", None),
        ("Rate (\u20b9)",                 "C",  "num0", None),
    ]),
    ("FINANCIALS", [
        ("Mat. Value (\u20b9)",           "H",  "num0", "sum"),
        ("GST @ 18% (\u20b9)",            "I",  "num0", "sum"),
        ("Mat. Value + GST (\u20b9)",     "J",  "num0", "sum"),
        ("TCS @ 2% (\u20b9)",             "K",  "num0", "sum"),
        ("TDS u/s 194(O) (\u20b9)",       "L",  "dec",  "sum"),
        ("Service Charge gross (\u20b9)", "M",  "dec",  "sum"),
        ("TDS u/s 194(H) (\u20b9)",       "N",  "dec",  "sum"),
        ("Net Service Charge (\u20b9)",   "O",  "dec",  "sum"),
        ("Svc Charge to MSTC (\u20b9)",   "P",  "num0", "sum"),
        ("GST TDS Rate",                  "Q",  "pct",  None),
        ("GST TDS (\u20b9)",              "R",  "num0", "sum"),
        ("Total Receivables (\u20b9)",    "S",  "num0", "sum"),
    ]),
    ("SECURITY DEPOSIT", [
        ("SD Expected (\u20b9)",          "T",  "num0", "sum"),
        ("SD Received (\u20b9)",          "U",  "num0", "sum"),
        ("SD Receipt Date",               "V",  "text", None),
    ]),
    ("FINAL PAYMENT", [
        ("FP Expected (\u20b9)",          "W",  "num0", "sum"),
        ("FP Received (\u20b9)",          "X",  "num0", "sum"),
        ("FP Receipt Date",               "Y",  "text", None),
    ]),
    ("LPP", [
        ("LPP Expected (\u20b9)",         "Z",  "num0", "sum"),
        ("LPP Received (\u20b9)",         "AA", "num0", "sum"),
        ("LPP Receipt Date",              "AB", "text", None),
    ]),
    ("SUMMARY", [
        ("Total Received (\u20b9)",       "AC", "num0", "sum"),
        ("Outstanding (\u20b9)",          "AD", "num0", "sum"),
        ("Payment Status",                None, "text", None),   # computed
    ]),
    ("DOCUMENT", [
        ("Invoice No.",                   "AE", "text", None),
        ("SAP Document",                  "AF", "num0", None),
        ("Doc./Invoice Date",             "AG", "text", None),
    ]),
]

# ledger column widths, keyed by field label (defaults per kind)
WIDTHS = {
    "Buyer": 30, "Lot No.": 8.5, "Lot Name": 50, "Bid Sheet / Auction": 10,
    "Quantity": 10, "Unit": 6.5, "Payment Status": 14, "Invoice No.": 15,
    "SAP Document": 13, "Src Row": 8,
}

# ledger column categories: (name, [field labels], colour)
CATEGORIES = [
    ("IDENTITY",         ["Buyer", "Lot No."], "404040"),
    ("LOT INFO",         ["Lot Name", "Bid Sheet / Auction", "Quantity", "Unit", "Rate (\u20b9)"], "1F4E79"),
    ("FINANCIALS",       [f[0] for f in SECTIONS[1][1]], "C55A11"),
    ("SECURITY DEPOSIT", [f[0] for f in SECTIONS[2][1]], "548235"),
    ("FINAL PAYMENT",    [f[0] for f in SECTIONS[3][1]], "7030A0"),
    ("LPP",              [f[0] for f in SECTIONS[4][1]], "00838F"),
    ("SUMMARY",          [f[0] for f in SECTIONS[5][1]], "C00000"),
    ("DOCUMENT",         [f[0] for f in SECTIONS[6][1]], "BF8F00"),
    ("TRACE",            ["Src Row"], "808080"),
]

# eight buyer themes: (accent, tint, pale)
THEMES = [
    ("1F4E79", "DDEBF7", "F4F9FD"),
    ("C55A11", "FCE4D6", "FEF4EC"),
    ("548235", "E2EFDA", "F4FAF0"),
    ("7030A0", "E6D9F2", "F6F1FB"),
    ("C00000", "FBE0E0", "FEF3F3"),
    ("00838F", "D7F0F2", "EFF9FA"),
    ("BF8F00", "FFF2CC", "FFFAE8"),
    ("A51E4D", "F8DCE6", "FDF2F6"),
]

THIN = Side(style="thin", color="D0D0D0")
MED = Side(style="medium", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
LEFTW = Alignment(horizontal="left", vertical="center", wrap_text=True)


def fill(rgb: str) -> PatternFill:
    return PatternFill("solid", fgColor=rgb)


def src(col: str, row: int) -> str:
    """Live link to one cell of the source sheet; a blank source stays blank."""
    ref = f"'{SRC}'!${col}${row}"
    return f'=IF({ref}="","",{ref})'


def status_formula(row: int) -> str:
    ref = f"'{SRC}'!$AD${row}"
    return f'=IF({ref}="","",IF({ref}<=0,"SETTLED","OUTSTANDING"))'


def criteria(buyer: str) -> str:
    return f"'{SRC}'!$G${FIRST_DATA_ROW}:$G${LAST_SRC_ROW},\"{buyer}\""


def read_lots(ws) -> "OrderedDict[str, list[int]]":
    """buyer -> [source row numbers], buyers alphabetical, lots by lot number."""
    buyers: "OrderedDict[str, list[int]]" = OrderedDict()
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        buyer, lot = ws[f"G{r}"].value, ws[f"F{r}"].value
        if buyer in (None, "") or lot in (None, ""):
            continue
        buyers.setdefault(str(buyer).strip(), []).append(r)
    for rows in buyers.values():
        rows.sort(key=lambda r: ws[f"F{r}"].value)
    return OrderedDict(sorted(buyers.items()))


# --------------------------------------------------------------------------- #
# sheet 1 - transposed collapsible view
# --------------------------------------------------------------------------- #
def build_collapsible(wb, buyers, sheet_name=None, field_filter=False, tab="ED7D31") -> None:
    n_lots = max(len(v) for v in buyers.values())
    my_name = sheet_name or COLLAPSIBLE
    ws = wb.create_sheet(my_name)
    last_idx = 1 + n_lots + 1                       # A + lot columns + total
    last_col = get_column_letter(last_idx)
    total_col = last_col

    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False    # banner heads the group
    ws.sheet_format.outlineLevelRow = 2
    ws.sheet_format.outlineLevelCol = 1

    ws.column_dimensions["A"].width = 34
    for i in range(n_lots):
        cd = ws.column_dimensions[get_column_letter(2 + i)]
        cd.width = 15.5
        cd.outlineLevel = 1                           # foldable lot columns
    ws.column_dimensions[total_col].width = 17

    # ---- title block ----------------------------------------------------- #
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER \u00d7 LOT COLLAPSIBLE VIEW   "
               "(fields \u2193 rows  |  lots \u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("Auctions 21977 \u2022 21978 \u2022 21979 \u2022 21980   \u2014   "
               "every figure below is a live link to 'Final Calculation Sheet'")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("2E75B6")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    ws.merge_cells(f"A3:{last_col}3")
    c = ws["A3"]
    c.value = ("HOW TO USE  \u25b6  the \u2212 / + buttons in the grey margin on the LEFT fold a whole buyer "
               "(outer level) or a single section such as FINANCIALS (inner level)   \u2022   the \u2212 / + above "
               f"column {total_col} folds all lot columns away and keeps only the buyer totals   \u2022   the "
               "1 2 3 buttons at the top left expand everything again   \u2022   a colour coded buyer index with "
               "totals and jump links sits at the bottom of this sheet")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[3].height = 32
    ws.row_dimensions[4].height = 6

    # ---- column header --------------------------------------------------- #
    hdr = 5
    ws.cell(row=hdr, column=1, value="FIELD  \u25b8  BUYER  |  LOT \u2192")
    for i in range(n_lots):
        ws.cell(row=hdr, column=2 + i, value=f"LOT {i + 1}")
    ws.cell(row=hdr, column=last_idx, value="BUYER TOTAL")
    for col in range(1, last_idx + 1):
        c = ws.cell(row=hdr, column=col)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = fill("404040")
        c.alignment = CENTER
        c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 26
    ws.freeze_panes = f"B{hdr + 1}"

    # ---- buyer blocks ---------------------------------------------------- #
    r = hdr + 1
    anchors = {}
    for bi, (buyer, rows) in enumerate(buyers.items()):
        accent, tint, pale = THEMES[bi % len(THEMES)]
        anchors[buyer] = r

        # buyer banner ------------------------------------------------------ #
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_idx)
        c = ws.cell(row=r, column=1)
        c.value = (f'="\u25bc  {buyer}   \u2022   "&COUNTIF({criteria(buyer)})&" LOT(S)   \u2022   '
                   f'MAT. VALUE \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$H${FIRST_DATA_ROW}:$H${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   RECEIVED \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AC${FIRST_DATA_ROW}:$AC${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   OUTSTANDING \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AD${FIRST_DATA_ROW}:$AD${LAST_SRC_ROW}),"#,##0")')
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = fill(accent)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        for col in range(1, last_idx + 1):
            ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        r += 1

        # lot number map ---------------------------------------------------- #
        c = ws.cell(row=r, column=1, value="   Lot No. \u2192")
        c.font = Font(bold=True, size=10, color=accent)
        c.alignment = LEFT
        for i, srow in enumerate(rows):
            cc = ws.cell(row=r, column=2 + i, value=src("F", srow))
            cc.font = Font(bold=True, size=10, color=accent)
            cc.alignment = CENTER
            cc.number_format = KIND_FMT["num0"]
        cc = ws.cell(row=r, column=last_idx, value="TOTAL")
        cc.font = Font(bold=True, size=10, color=accent)
        cc.alignment = CENTER
        for col in range(1, last_idx + 1):
            ws.cell(row=r, column=col).fill = fill(tint)
            ws.cell(row=r, column=col).border = BOX
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].height = 18
        r += 1

        # sections ---------------------------------------------------------- #
        for sname, fields in SECTIONS:
            for col in range(1, last_idx + 1):
                cc = ws.cell(row=r, column=col)
                cc.fill = fill(tint)
                cc.border = BOX
            c = ws.cell(row=r, column=1, value=f"  \u25b8 {sname}")
            c.font = Font(bold=True, size=10, color=accent)
            c.alignment = LEFT
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].height = 16
            r += 1

            for label, scol, kind, agg in fields:
                c = ws.cell(row=r, column=1, value=f"      {label}")
                c.font = Font(size=10, color="333333")
                c.alignment = LEFT
                for i, srow in enumerate(rows):
                    cc = ws.cell(row=r, column=2 + i)
                    cc.value = status_formula(srow) if scol is None else src(scol, srow)
                    cc.number_format = KIND_FMT[kind]
                    cc.font = Font(size=9 if label == "Lot Name" else 10, bold=(scol is None))
                    cc.alignment = LEFTW if label == "Lot Name" else (
                        LEFT if kind == "text" else RIGHT)
                    if len(rows) > 1 and i % 2 == 1:
                        cc.fill = fill(pale)
                tc = ws.cell(row=r, column=last_idx)
                if agg == "sum":
                    tc.value = f"=SUM(B{r}:{get_column_letter(last_idx - 1)}{r})"
                    tc.number_format = KIND_FMT[kind]
                    tc.font = Font(bold=True, size=10, color=accent)
                    tc.alignment = RIGHT
                else:
                    tc.value = "\u2013"
                    tc.font = Font(size=10, color="808080")
                    tc.alignment = CENTER
                tc.fill = fill(tint)
                for col in range(1, last_idx + 1):
                    ws.cell(row=r, column=col).border = BOX
                ws.cell(row=r, column=last_idx).border = Border(
                    left=Side(style="thin", color=accent), right=THIN, top=THIN, bottom=THIN)
                ws.row_dimensions[r].outlineLevel = 2
                ws.row_dimensions[r].height = 60 if label == "Lot Name" else 15
                r += 1

        # spacer ------------------------------------------------------------ #
        for col in range(1, last_idx + 1):
            ws.cell(row=r, column=col).fill = fill("EDEDED")
        ws.row_dimensions[r].height = 7
        r += 1

    last_block_row = r - 2                    # last field row of the last block
    # ---- buyer index (bottom of the sheet) -------------------------------- #
    idx_title = r + 1
    ws.merge_cells(start_row=idx_title, start_column=1, end_row=idx_title, end_column=last_idx)
    c = ws.cell(row=idx_title, column=1,
                value="\u25c8  BUYER INDEX  \u2014  click a buyer name to jump to its block")
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[idx_title].height = 22

    heads = ["BUYER (click to jump)", "LOTS", "MAT. VALUE \u20b9", "TOTAL RECEIVABLES \u20b9",
             "TOTAL RECEIVED \u20b9", "OUTSTANDING \u20b9", "STATUS"]
    hr = idx_title + 1
    for i, h in enumerate(heads):
        c = ws.cell(row=hr, column=1 + i, value=h)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = fill("595959")
        c.alignment = CENTER
        c.border = BOX
    ws.row_dimensions[hr].height = 26

    first_idx = hr + 1
    for bi, buyer in enumerate(buyers):
        accent, tint, pale = THEMES[bi % len(THEMES)]
        rr = first_idx + bi
        c = ws.cell(row=rr, column=1, value=buyer)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{my_name}'!A{anchors[buyer]}",
                                display=buyer, tooltip=f"Jump to {buyer}")
        c.font = Font(bold=True, size=10, color="0563C1", underline="single")
        c.alignment = LEFT
        ws.cell(row=rr, column=2, value=f"=COUNTIF({criteria(buyer)})").number_format = "0"
        ws.cell(row=rr, column=3, value=f"=SUMIF({criteria(buyer)},"
                                        f"'{SRC}'!$H${FIRST_DATA_ROW}:$H${LAST_SRC_ROW})")
        ws.cell(row=rr, column=4, value=f"=SUMIF({criteria(buyer)},"
                                        f"'{SRC}'!$S${FIRST_DATA_ROW}:$S${LAST_SRC_ROW})")
        ws.cell(row=rr, column=5, value=f"=SUMIF({criteria(buyer)},"
                                        f"'{SRC}'!$AC${FIRST_DATA_ROW}:$AC${LAST_SRC_ROW})")
        ws.cell(row=rr, column=6, value=f"=SUMIF({criteria(buyer)},"
                                        f"'{SRC}'!$AD${FIRST_DATA_ROW}:$AD${LAST_SRC_ROW})")
        ws.cell(row=rr, column=7, value=f'=IF(F{rr}<=0,"SETTLED","OUTSTANDING")')
        for col in range(2, 8):
            cc = ws.cell(row=rr, column=col)
            cc.number_format = "0" if col == 2 else KIND_FMT["num0"] if col != 7 else "General"
            cc.font = Font(bold=True, size=10, color=accent)
            cc.alignment = CENTER if col in (2, 7) else RIGHT
        for col in range(1, 8):
            cc = ws.cell(row=rr, column=col)
            cc.fill = fill(pale if bi % 2 else tint)
            cc.border = BOX
        ws.row_dimensions[rr].height = 16

    tr = first_idx + len(buyers)
    ws.cell(row=tr, column=1, value="\u2211  ALL BUYERS")
    ws.cell(row=tr, column=2, value=f"=SUM(B{first_idx}:B{tr - 1})").number_format = "0"
    for col in range(3, 7):
        letter = get_column_letter(col)
        ws.cell(row=tr, column=col,
                value=f"=SUM({letter}{first_idx}:{letter}{tr - 1})").number_format = KIND_FMT["num0"]
    ws.cell(row=tr, column=7, value=f'=IF(F{tr}<=0,"SETTLED","OUTSTANDING")')
    for col in range(1, 8):
        cc = ws.cell(row=tr, column=col)
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.fill = fill("1F3864")
        cc.alignment = LEFT if col == 1 else (CENTER if col in (2, 7) else RIGHT)
        cc.border = BOX
    ws.row_dimensions[tr].height = 18

    ws.conditional_formatting.add(
        f"G{first_idx}:G{tr}",
        CellIsRule(operator="equal", formula=['"SETTLED"'], font=Font(bold=True, color="006100"),
                   fill=fill("C6EFCE")))
    ws.conditional_formatting.add(
        f"G{first_idx}:G{tr}",
        CellIsRule(operator="equal", formula=['"OUTSTANDING"'], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))

    if field_filter:
        # a one-column range = a single filter button on the label column, so the
        # 13 lot headers do not sprout an arrow each; it stops above the index so
        # filtering never hides the buyer jump table
        ws.auto_filter.ref = f"A{hdr}:A{last_block_row}"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# --------------------------------------------------------------------------- #
# sheet 2 - flat filterable ledger
# --------------------------------------------------------------------------- #
def build_ledger(wb, buyers) -> None:
    ws = wb.create_sheet(LEDGER)
    ws.sheet_properties.tabColor = "2E75B6"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False   # category fold buttons on the left
    ws.sheet_format.outlineLevelRow = 1
    ws.sheet_format.outlineLevelCol = 1

    # ---- column plan ----------------------------------------------------- #
    plan = [("Buyer", "G", "text", None)]
    for _sname, fields in SECTIONS:
        for label, scol, kind, agg in fields:
            plan.append((label, scol, kind, agg))
    plan.append(("Src Row", None, "num0", None))

    letters, cat_of = {}, {}
    for i, (label, _s, kind, _a) in enumerate(plan):
        letter = get_column_letter(i + 1)
        letters[label] = letter
        ws.column_dimensions[letter].width = WIDTHS.get(
            label, 13.5 if kind in ("num0", "dec") else 12)
    for name, labels, _colour in CATEGORIES:
        for label in labels:
            cat_of[label] = name
    last_col = letters["Src Row"]

    def col_of(label: str) -> str:
        return letters[label]

    # ---- title + instructions -------------------------------------------- #
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER & LOT LEDGER  (filter + fold)   "
               "(lots \u2193 rows  |  fields \u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  filter with the \u25bc arrows in row 5 (buyer, lot no., bid sheet, unit, "
               "payment status \u2026)   \u2022   fold a buyer with the \u2212 button in the left margin \u2014 its "
               "coloured summary row stays on screen   \u2022   fold a column category (FINANCIALS, SECURITY "
               "DEPOSIT \u2026) with the \u2212 above its first column; that key column stays on screen   "
               "\u2022   the KPI band in row 3 and the GRAND TOTAL row always follow the filter")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 28

    # ---- category band (row 4) ------------------------------------------- #
    cat_colour = {name: colour for name, _l, colour in CATEGORIES}
    i = 0
    while i < len(plan):
        name = cat_of[plan[i][0]]
        j = i
        while j + 1 < len(plan) and cat_of[plan[j + 1][0]] == name:
            j += 1
        c1, c2 = get_column_letter(i + 1), get_column_letter(j + 1)
        if c1 != c2:
            ws.merge_cells(f"{c1}4:{c2}4")
        if name not in ("IDENTITY", "TRACE"):
            # one fold button per category: the first (key) column of the
            # category stays visible, the rest fold away.  summaryRight=False
            # puts the button above that first column.
            for k in range(i + 1, j + 1):
                ws.column_dimensions[get_column_letter(k + 1)].outlineLevel = 1
        c = ws[f"{c1}4"]
        c.value = name
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = fill(cat_colour[name])
        c.alignment = CENTER
        for k in range(i, j + 1):
            ws.cell(row=4, column=k + 1).border = BOX
        i = j + 1
    ws.row_dimensions[4].height = 17

    # ---- header row (row 5) ---------------------------------------------- #
    hdr = 5
    for label, _s, _k, _a in plan:
        c = ws[f"{col_of(label)}{hdr}"]
        c.value = label
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = fill(cat_colour[cat_of[label]])
        c.alignment = CENTER
        c.border = BOX
    ws.row_dimensions[hdr].height = 42
    ws.freeze_panes = f"C{hdr + 1}"

    # ---- lot rows, grouped under a buyer summary row ---------------------- #
    r = hdr + 1
    blocks = []
    for bi, (buyer, rows) in enumerate(buyers.items()):
        accent, tint, pale = THEMES[bi % len(THEMES)]
        brow = r
        r += 1
        first_lot = r
        for k, srow in enumerate(rows):
            for label, scol, kind, _a in plan:
                cc = ws[f"{col_of(label)}{r}"]
                if label == "Src Row":
                    cc.value = srow
                elif label == "Payment Status":
                    cc.value = status_formula(srow)
                else:
                    cc.value = src(scol, srow)
                cc.number_format = KIND_FMT[kind]
                cc.font = Font(size=9, color="333333")
                cc.alignment = (LEFTW if label in ("Buyer", "Lot Name")
                                else LEFT if kind == "text" else RIGHT)
                cc.fill = fill(pale if k % 2 else tint)
                cc.border = BOX
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].height = 15
            r += 1
        blocks.append((buyer, brow, first_lot, r - 1, accent))
    data_first, data_last = hdr + 2, r - 1

    # ---- buyer summary rows (SUBTOTAL -> they follow the filter) ---------- #
    for buyer, brow, f_lot, l_lot, accent in blocks:
        n = l_lot - f_lot + 1
        for label, _s, kind, agg in plan:
            cc = ws[f"{col_of(label)}{brow}"]
            if label == "Buyer":
                cc.value = f"\u25bc  {buyer}"
                cc.font = Font(bold=True, size=10, color="FFFFFF")
                cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif label == "Lot Name":
                cc.value = f"{n} lot{'s' if n > 1 else ''}  \u2014  click \u2212 to fold this buyer"
                cc.font = Font(bold=True, size=9, color="FFFFFF")
                cc.alignment = LEFT
            elif agg == "sum":
                cc.value = f"=SUBTOTAL(109,{col_of(label)}{f_lot}:{col_of(label)}{l_lot})"
                cc.number_format = KIND_FMT[kind]
                cc.font = Font(bold=True, size=10, color="FFFFFF")
                cc.alignment = RIGHT
            cc.fill = fill(accent)
            cc.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        ws.row_dimensions[brow].height = 20

    # ---- grand total ------------------------------------------------------ #
    gt = r
    for label, _s, kind, agg in plan:
        cc = ws[f"{col_of(label)}{gt}"]
        if label == "Buyer":
            cc.value = "\u2211  GRAND TOTAL (follows the filter)"
            cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        elif agg == "sum":
            cc.value = f"=SUBTOTAL(109,{col_of(label)}{data_first}:{col_of(label)}{data_last})"
            cc.number_format = KIND_FMT[kind]
            cc.alignment = RIGHT
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.fill = fill("1F3864")
        cc.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[gt].height = 20

    # ---- KPI band (row 3) ------------------------------------------------- #
    kpis = [("LOTS VISIBLE", "103", col_of("Lot No."), "0"),
            ("MAT. VALUE \u20b9", "109", col_of("Mat. Value (\u20b9)"), KIND_FMT["num0"]),
            ("RECEIVABLES \u20b9", "109", col_of("Total Receivables (\u20b9)"), KIND_FMT["num0"]),
            ("SD RECEIVED \u20b9", "109", col_of("SD Received (\u20b9)"), KIND_FMT["num0"]),
            ("FP RECEIVED \u20b9", "109", col_of("FP Received (\u20b9)"), KIND_FMT["num0"]),
            ("TOTAL RECEIVED \u20b9", "109", col_of("Total Received (\u20b9)"), KIND_FMT["num0"]),
            ("OUTSTANDING \u20b9", "109", col_of("Outstanding (\u20b9)"), KIND_FMT["num0"])]
    col = 1
    for label, fn, letter, fmt in kpis:
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        lab = ws.cell(row=3, column=col, value=label + "  \u25b8")
        lab.font = Font(bold=True, size=9, color="FFFFFF")
        lab.fill = fill("7030A0")
        lab.alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=3, start_column=col + 2, end_row=3, end_column=col + 3)
        val = ws.cell(row=3, column=col + 2,
                      value=f"=SUBTOTAL({fn},${letter}${data_first}:${letter}${data_last})")
        val.font = Font(bold=True, size=11, color="FFFFFF")
        val.fill = fill("9E4EA8")
        val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val.number_format = fmt
        col += 4
    ws.row_dimensions[3].height = 22

    # ---- filter + conditional formatting ---------------------------------- #
    ws.auto_filter.ref = f"A{hdr}:{last_col}{data_last}"
    status_col = col_of("Payment Status")
    out_col = col_of("Outstanding (\u20b9)")
    inv_col = col_of("Invoice No.")
    ws.conditional_formatting.add(
        f"{out_col}{data_first}:{out_col}{data_last}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))
    ws.conditional_formatting.add(
        f"{out_col}{data_first}:{out_col}{data_last}",
        CellIsRule(operator="lessThanOrEqual", formula=["0"], font=Font(bold=True, color="006100")))
    ws.conditional_formatting.add(
        f"{status_col}{data_first}:{status_col}{data_last}",
        CellIsRule(operator="equal", formula=['"SETTLED"'], font=Font(bold=True, color="006100"),
                   fill=fill("C6EFCE")))
    ws.conditional_formatting.add(
        f"{status_col}{data_first}:{status_col}{data_last}",
        CellIsRule(operator="equal", formula=['"OUTSTANDING"'], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))
    ws.conditional_formatting.add(
        f"A{data_first}:{last_col}{data_last}",
        FormulaRule(formula=[f'AND(ISNUMBER($B{data_first}),${inv_col}{data_first}="")'],
                    fill=fill("FFE699")))
    mat_col = col_of("Mat. Value (\u20b9)")
    ws.conditional_formatting.add(
        f"{mat_col}{data_first}:{mat_col}{data_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="638EC6",
                    showValue=True))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hdr}:{hdr}"


# --------------------------------------------------------------------------- #
# sheet 3 - collapsible, lots as rows / fields as columns
# --------------------------------------------------------------------------- #
def plan_fields():
    """The field catalogue as a column plan for the lots-as-rows layouts."""
    return [(label, scol, kind, agg) for _n, fields in SECTIONS
            for (label, scol, kind, agg) in fields]


SECTION_COLOURS = ["1F4E79", "C55A11", "548235", "7030A0", "00838F", "C00000", "BF8F00"]


def colour_of(label):
    for (name, fields), colour in zip(SECTIONS, SECTION_COLOURS):
        if label in [f[0] for f in fields]:
            return colour
    return "404040"


def build_collapsible_rows(wb, buyers):
    """One colour block per buyer; inside it one row per lot."""
    ws = wb.create_sheet(COLLAPSIBLE_ROWS)
    plan = plan_fields()
    letters = {}
    for i, (label, _s, kind, _a) in enumerate(plan):
        letter = get_column_letter(i + 1)
        letters[label] = letter
        ws.column_dimensions[letter].width = WIDTHS.get(
            label, 13.5 if kind in ("num0", "dec") else 12)
    last_col = letters[plan[-1][0]]

    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_format.outlineLevelRow = 1
    ws.sheet_format.outlineLevelCol = 1

    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER \u00d7 LOT COLLAPSIBLE VIEW   "
               "(lots \u2193 rows  |  fields \u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  the \u2212 / + buttons in the left margin fold a buyer: its lots hide and the "
               "banner plus the \u2211 totals row stay on screen   \u2022   the \u2212 above a column category "
               "(LOT INFO, FINANCIALS \u2026) folds those columns away, keeping the category's first column   "
               "\u2022   every figure is a live link to 'Final Calculation Sheet'   \u2022   for filtering use "
               f"'{LEDGER}' or '{LEDGER_COLS}'")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 30

    # category band + header row ------------------------------------------- #
    band, hdr = 3, 4
    i = 0
    for (name, fields), colour in zip(SECTIONS, SECTION_COLOURS):
        n = len(fields)
        c1, c2 = get_column_letter(i + 1), get_column_letter(i + n)
        if c1 != c2:
            ws.merge_cells(f"{c1}{band}:{c2}{band}")
        for k in range(i, i + n):
            cc = ws.cell(row=band, column=k + 1)
            cc.fill = fill(colour)
            cc.border = BOX
            if k > i:                      # first column of a category stays visible
                ws.column_dimensions[get_column_letter(k + 1)].outlineLevel = 1
        cc = ws[f"{c1}{band}"]
        cc.value = name
        cc.font = Font(bold=True, size=9, color="FFFFFF")
        cc.alignment = CENTER
        i += n
    for label, _s, _k, _a in plan:
        cc = ws[f"{letters[label]}{hdr}"]
        cc.value = label
        cc.font = Font(bold=True, size=9, color="FFFFFF")
        cc.fill = fill(colour_of(label))
        cc.alignment = CENTER
        cc.border = BOX
    ws.row_dimensions[band].height = 17
    ws.row_dimensions[hdr].height = 42
    ws.freeze_panes = f"C{hdr + 1}"

    # buyer blocks ---------------------------------------------------------- #
    r = hdr + 1
    for bi, (buyer, rows) in enumerate(buyers.items()):
        accent, tint, pale = THEMES[bi % len(THEMES)]
        ws.merge_cells(start_row=r, start_column=1, end_row=r,
                       end_column=len(plan))
        c = ws.cell(row=r, column=1)
        c.value = (f'="\u25bc  {buyer}   \u2022   "&COUNTIF({criteria(buyer)})&" LOT(S)   \u2022   '
                   f'MAT. VALUE \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$H${FIRST_DATA_ROW}:$H${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   RECEIVED \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AC${FIRST_DATA_ROW}:$AC${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   OUTSTANDING \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AD${FIRST_DATA_ROW}:$AD${LAST_SRC_ROW}),"#,##0")')
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = fill(accent)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        for col in range(1, len(plan) + 1):
            ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        r += 1

        first_lot = r
        for k, srow in enumerate(rows):
            for label, scol, kind, _a in plan:
                cc = ws[f"{letters[label]}{r}"]
                cc.value = status_formula(srow) if scol is None else src(scol, srow)
                cc.number_format = KIND_FMT[kind]
                cc.font = Font(size=9, color="333333")
                cc.alignment = (LEFTW if label == "Lot Name"
                                else LEFT if kind == "text" else RIGHT)
                cc.fill = fill(pale if k % 2 else tint)
                cc.border = BOX
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].height = 15
            r += 1
        last_lot = r - 1

        c = ws.cell(row=r, column=1, value=f"\u2211  {buyer} \u2014 TOTALS")
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for label, _s, kind, agg in plan:
            cc = ws[f"{letters[label]}{r}"]
            if agg == "sum":
                cc.value = f"=SUM({letters[label]}{first_lot}:{letters[label]}{last_lot})"
                cc.number_format = KIND_FMT[kind]
                cc.alignment = RIGHT
            elif label != "Lot No.":
                cc.value = "\u2013"
                cc.alignment = CENTER
            cc.font = Font(bold=True, size=10, color="FFFFFF")
            cc.fill = fill(accent)
            cc.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        ws.row_dimensions[r].height = 20
        r += 1

        for col in range(1, len(plan) + 1):
            ws.cell(row=r, column=col).fill = fill("EDEDED")
        ws.row_dimensions[r].height = 7
        r += 1

    # conditional formatting ------------------------------------------------ #
    out_col = letters["Outstanding (\u20b9)"]
    status_col = letters["Payment Status"]
    inv_col = letters["Invoice No."]
    first_lot, last_row = hdr + 2, r - 1
    ws.conditional_formatting.add(
        f"{out_col}{first_lot}:{out_col}{last_row}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))
    ws.conditional_formatting.add(
        f"{status_col}{first_lot}:{status_col}{last_row}",
        CellIsRule(operator="equal", formula=['"SETTLED"'], font=Font(bold=True, color="006100"),
                   fill=fill("C6EFCE")))
    ws.conditional_formatting.add(
        f"{status_col}{first_lot}:{status_col}{last_row}",
        CellIsRule(operator="equal", formula=['"OUTSTANDING"'], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))
    # only lot rows (a numeric lot no. in column A) - banners and totals stay clean
    ws.conditional_formatting.add(
        f"{inv_col}{first_lot}:{inv_col}{last_row}",
        FormulaRule(formula=[f'AND(ISNUMBER($A{first_lot}),${inv_col}{first_lot}="")'],
                    fill=fill("FFE699")))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hdr}:{hdr}"


# --------------------------------------------------------------------------- #
# sheet 4 - ledger, fields as rows / lots as columns
# --------------------------------------------------------------------------- #
def build_ledger_cols(wb, buyers):
    """One flat table: every lot is a column, every field a row, with a filter."""
    ws = wb.create_sheet(LEDGER_COLS)
    all_lots = [(buyer, srow) for buyer, rows in buyers.items() for srow in rows]
    n = len(all_lots)
    total_col = get_column_letter(1 + n + 1)

    ws.sheet_properties.tabColor = "00B0F0"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.sheet_format.outlineLevelRow = 1
    ws.sheet_format.outlineLevelCol = 1

    ws.column_dimensions["A"].width = 34
    for i in range(n):
        cd = ws.column_dimensions[get_column_letter(2 + i)]
        cd.width = 14
        cd.outlineLevel = 1                    # fold every lot column at once
    ws.column_dimensions[total_col].width = 17

    ws.merge_cells(f"A1:{total_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER & LOT LEDGER  (filter + fold)   "
               "(fields \u2193 rows  |  lots \u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{total_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  filter with the \u25bc arrow on the FIELD column in row 5 to keep only the "
               "lines you need   \u2022   the \u2212 / + in the left margin folds a whole field category   "
               f"\u2022   the \u2212 / + above column {total_col} folds all {n} lot columns away and keeps the "
               "totals   \u2022   every figure is a live link to 'Final Calculation Sheet'")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 30

    # buyer band (row 3) - one merged, colour coded strip per buyer ---------- #
    col = 2
    for bi, (buyer, rows) in enumerate(buyers.items()):
        accent = THEMES[bi % len(THEMES)][0]
        span = len(rows)
        if span > 1:
            ws.merge_cells(start_row=3, start_column=col, end_row=3,
                           end_column=col + span - 1)
        cc = ws.cell(row=3, column=col, value=f"{buyer}  ({span})")
        cc.font = Font(bold=True, size=9, color="FFFFFF")
        cc.alignment = CENTER
        for k in range(span):
            x = ws.cell(row=3, column=col + k)
            x.fill = fill(accent)
            x.border = BOX
        col += span
    c = ws.cell(row=3, column=1, value="BUYER  \u2192")
    c.font = Font(bold=True, size=9, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=3, column=2 + n).fill = fill("404040")
    ws.row_dimensions[3].height = 18

    # lot number band (row 4) ------------------------------------------------ #
    c = ws.cell(row=4, column=1, value="LOT NO.  \u2192")
    c.font = Font(bold=True, size=9, color="1F3864")
    c.fill = fill("DDEBF7")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i, (_b, srow) in enumerate(all_lots):
        cc = ws.cell(row=4, column=2 + i, value=src("F", srow))
        cc.number_format = KIND_FMT["num0"]
        cc.font = Font(bold=True, size=9, color="1F3864")
        cc.fill = fill("DDEBF7")
        cc.alignment = CENTER
        cc.border = BOX
    cc = ws.cell(row=4, column=2 + n, value="\u2211 ALL")
    cc.font = Font(bold=True, size=9, color="1F3864")
    cc.fill = fill("DDEBF7")
    cc.alignment = CENTER
    ws.row_dimensions[4].height = 18

    # header row (row 5) - carries the AutoFilter ---------------------------- #
    hdr = 5
    c = ws.cell(row=hdr, column=1, value="FIELD  \u25b8")
    for i in range(n):
        ws.cell(row=hdr, column=2 + i, value=f"LOT {i + 1}")
    ws.cell(row=hdr, column=2 + n, value="TOTAL")
    for k in range(1, 3 + n):
        cc = ws.cell(row=hdr, column=k)
        cc.font = Font(bold=True, size=9, color="FFFFFF")
        cc.fill = fill("404040")
        cc.alignment = CENTER
        cc.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 26
    ws.freeze_panes = f"B{hdr + 1}"

    # category header rows + field rows ------------------------------------- #
    r = hdr + 1
    field_rows = {}
    for (name, fields), colour in zip(SECTIONS, SECTION_COLOURS):
        for k in range(1, 3 + n):
            cc = ws.cell(row=r, column=k)
            cc.fill = fill(colour)
            cc.border = BOX
        cc = ws.cell(row=r, column=1, value=f"  \u25b8 {name}")
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.alignment = LEFT
        ws.row_dimensions[r].height = 17
        r += 1
        for label, scol, kind, agg in fields:
            cc = ws.cell(row=r, column=1, value=f"      {label}")
            cc.font = Font(size=10, color="333333")
            cc.alignment = LEFT
            field_rows[label] = r
            for i, (_b, srow) in enumerate(all_lots):
                x = ws.cell(row=r, column=2 + i)
                x.value = status_formula(srow) if scol is None else src(scol, srow)
                x.number_format = KIND_FMT[kind]
                x.font = Font(size=9 if label == "Lot Name" else 10,
                              bold=(scol is None))
                x.alignment = LEFTW if label == "Lot Name" else (
                    LEFT if kind == "text" else RIGHT)
                x.border = BOX
                if i % 2:
                    x.fill = fill("F7F7F7")
            t = ws.cell(row=r, column=2 + n)
            if agg == "sum":
                t.value = f"=SUM(B{r}:{get_column_letter(1 + n)}{r})"
                t.number_format = KIND_FMT[kind]
                t.font = Font(bold=True, size=10, color="1F3864")
                t.alignment = RIGHT
            else:
                t.value = "\u2013"
                t.font = Font(size=10, color="808080")
                t.alignment = CENTER
            t.fill = fill("DDEBF7")
            for k in range(1, 3 + n):
                ws.cell(row=r, column=k).border = BOX
            ws.cell(row=r, column=2 + n).border = Border(
                left=Side(style="thin", color="1F4E79"), right=THIN, top=THIN, bottom=THIN)
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].height = 44 if label == "Lot Name" else 15
            r += 1

    last_row = r - 1
    ws.auto_filter.ref = f"A{hdr}:{total_col}{last_row}"

    # conditional formatting ------------------------------------------------ #
    for label, rules in (
            ("Outstanding (\u20b9)", [
                CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                           fill=fill("FFC7CE")),
                CellIsRule(operator="lessThanOrEqual", formula=["0"], font=Font(bold=True, color="006100"))]),
            ("Payment Status", [
                CellIsRule(operator="equal", formula=['"SETTLED"'], font=Font(bold=True, color="006100"),
                           fill=fill("C6EFCE")),
                CellIsRule(operator="equal", formula=['"OUTSTANDING"'], font=Font(bold=True, color="9C0006"),
                           fill=fill("FFC7CE"))]),
            ("Invoice No.", [CellIsRule(operator="equal", formula=['""'], fill=fill("FFE699"))])):
        rr = field_rows[label]
        for rule in rules:
            ws.conditional_formatting.add(f"B{rr}:{get_column_letter(1 + n)}{rr}", rule)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True



# --------------------------------------------------------------------------- #
# sheet 5 - buyer > lot > fields, everything running downwards
# --------------------------------------------------------------------------- #
def build_lot_vertical(wb, buyers):
    """Three fold levels: buyer, then lot, then the fields of that lot."""
    ws = wb.create_sheet(LOT_VERTICAL, 1)
    ws.sheet_properties.tabColor = "00B050"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.sheet_format.outlineLevelRow = 3
    ws.sheet_format.outlineLevelCol = 0
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 30

    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER \u25b8 LOT \u25b8 FIELD  \u2014  VERTICAL DETAIL "
               "(everything reads downwards)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:B2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  the \u2212 / + buttons in the left margin fold three levels: "
               "1 = buyer names only, 2 = buyer + their lot numbers, 3 = buyer + lot + section "
               "headings, 4 = every field of every lot   \u2022   each lot is its own block with the "
               "fields listed one below the other   \u2022   every buyer ends with its own "
               "\u2211 TOTALS block, and the sheet ends with an all-buyers grand total   "
               "\u2022   every figure is a live link to 'Final Calculation Sheet'")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 44
    ws.row_dimensions[3].height = 6

    hdr = 4
    for col, text, width in ((1, "FIELD  \u25b8  BUYER  /  LOT  /  DETAIL", None), (2, "VALUE", None)):
        cc = ws.cell(row=hdr, column=col, value=text)
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.fill = fill("404040")
        cc.alignment = CENTER if col == 2 else Alignment(horizontal="left", vertical="center", indent=1)
        cc.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 24
    ws.freeze_panes = f"A{hdr + 1}"

    r = hdr + 1
    total_rows = {}          # field label -> [row of each buyer totals block]
    for bi, (buyer, rows) in enumerate(buyers.items()):
        accent, tint, pale = THEMES[bi % len(THEMES)]

        # buyer banner ------------------------------------------------------ #
        c = ws.cell(row=r, column=1)
        c.value = (f'="\u25bc  {buyer}   \u2022   "&COUNTIF({criteria(buyer)})&" LOT(S)   \u2022   '
                   f'MAT. VALUE \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$H${FIRST_DATA_ROW}:$H${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   RECEIVED \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AC${FIRST_DATA_ROW}:$AC${LAST_SRC_ROW}),"#,##0")'
                   f'&"   \u2022   OUTSTANDING \u20b9"&TEXT(SUMIF({criteria(buyer)},\'{SRC}\'!$AD${FIRST_DATA_ROW}:$AD${LAST_SRC_ROW}),"#,##0")')
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = fill(accent)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in (1, 2):
            ws.cell(row=r, column=col).fill = fill(accent)
            ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        ws.row_dimensions[r].height = 24
        r += 1

        # one block per lot -------------------------------------------------- #
        for srow in rows:
            lot = wb[SRC][f"F{srow}"].value
            c = ws.cell(row=r, column=1)
            c.value = (f'="   \u25b8  LOT "&\'{SRC}\'!$F${srow}&"   \u2022   "&\'{SRC}\'!$B${srow}&'
                       f'"   \u2022   MAT. VALUE \u20b9"&TEXT(\'{SRC}\'!$H${srow},"#,##0")&'
                       f'"   \u2022   OUTSTANDING \u20b9"&TEXT(\'{SRC}\'!$AD${srow},"#,##0")')
            c.font = Font(bold=True, size=11, color=accent)
            for col in (1, 2):
                ws.cell(row=r, column=col).fill = fill(tint)
                ws.cell(row=r, column=col).border = BOX
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].height = 20
            r += 1

            for sname, fields in SECTIONS:
                c = ws.cell(row=r, column=1, value=f"      \u25b8 {sname}")
                c.font = Font(bold=True, size=10, color=accent)
                c.alignment = LEFT
                for col in (1, 2):
                    ws.cell(row=r, column=col).fill = fill(pale)
                    ws.cell(row=r, column=col).border = BOX
                ws.row_dimensions[r].outlineLevel = 2
                ws.row_dimensions[r].height = 16
                r += 1

                for label, scol, kind, _agg in fields:
                    a = ws.cell(row=r, column=1, value=f"            {label}")
                    a.font = Font(size=10, color="333333")
                    a.alignment = LEFT
                    b = ws.cell(row=r, column=2)
                    b.value = status_formula(srow) if scol is None else src(scol, srow)
                    b.number_format = KIND_FMT[kind]
                    b.font = Font(size=10, bold=(scol is None))
                    b.alignment = LEFT if kind == "text" else RIGHT
                    for col in (1, 2):
                        ws.cell(row=r, column=col).border = BOX
                    ws.row_dimensions[r].outlineLevel = 3
                    ws.row_dimensions[r].height = 15
                    r += 1

        # buyer totals block -------------------------------------------------- #
        c = ws.cell(row=r, column=1, value=f"   \u2211  {buyer} \u2014 TOTALS (all its lots)")
        c.font = Font(bold=True, size=11, color="FFFFFF")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in (1, 2):
            ws.cell(row=r, column=col).fill = fill(accent)
            ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].height = 20
        r += 1
        for _sname, fields in SECTIONS:
            for label, scol, kind, agg in fields:
                if agg != "sum":
                    continue
                a = ws.cell(row=r, column=1, value=f"         {label}")
                a.font = Font(size=10, color="333333")
                a.alignment = LEFT
                b = ws.cell(row=r, column=2,
                            value=f"=SUMIF({criteria(buyer)},"
                                  f"'{SRC}'!${scol}${FIRST_DATA_ROW}:${scol}${LAST_SRC_ROW})")
                b.number_format = KIND_FMT[kind]
                b.font = Font(bold=True, size=10, color=accent)
                b.alignment = RIGHT
                for col in (1, 2):
                    ws.cell(row=r, column=col).fill = fill(tint)
                    ws.cell(row=r, column=col).border = BOX
                ws.row_dimensions[r].outlineLevel = 2
                ws.row_dimensions[r].height = 15
                total_rows.setdefault(label, []).append(r)
                r += 1

        for col in (1, 2):
            ws.cell(row=r, column=col).fill = fill("EDEDED")
        ws.row_dimensions[r].height = 7
        r += 1

    # grand total block ------------------------------------------------------- #
    c = ws.cell(row=r, column=1, value="\u2211\u2211  ALL BUYERS \u2014 GRAND TOTAL")
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in (1, 2):
        ws.cell(row=r, column=col).fill = fill("1F3864")
        ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[r].height = 24
    r += 1
    for label, refs in total_rows.items():
        kind = next(k for _n, f in SECTIONS for (l, _s, k, a) in f if l == label)
        a = ws.cell(row=r, column=1, value=f"      {label}")
        a.font = Font(size=10, color="333333")
        a.alignment = LEFT
        b = ws.cell(row=r, column=2, value="=" + "+".join(f"B{x}" for x in refs))
        b.number_format = KIND_FMT[kind]
        b.font = Font(bold=True, size=10, color="1F3864")
        b.alignment = RIGHT
        for col in (1, 2):
            ws.cell(row=r, column=col).fill = fill("DDEBF7")
            ws.cell(row=r, column=col).border = BOX
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].height = 15
        r += 1

    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True



# --------------------------------------------------------------------------- #
# sheet 6 - buyer pivot: buyers across the top, every detail down the side
# --------------------------------------------------------------------------- #
def _rng(col):
    return f"'{SRC}'!${col}${FIRST_DATA_ROW}:${col}${LAST_SRC_ROW}"


def _sumif(b, col):
    return f"SUMIF('{SRC}'!$G${FIRST_DATA_ROW}:$G${LAST_SRC_ROW},{b},{_rng(col)})"


def _countif(b):
    return f"COUNTIF('{SRC}'!$G${FIRST_DATA_ROW}:$G${LAST_SRC_ROW},{b})"


def _countblank(b, col):
    """how many of that buyer's lots have an empty cell in `col`.

    Deliberately uses the "" criterion rather than "<>": the non-blank form is
    not evaluated consistently across spreadsheet engines, the blank form is.
    """
    return (f"COUNTIFS('{SRC}'!$G${FIRST_DATA_ROW}:$G${LAST_SRC_ROW},{b},"
            f"{_rng(col)},\"\")")


# labels reused by the derived rows (kept in constants: an f-string expression
# may not contain a backslash escape)
L_OUT = "Outstanding (\u20b9)"
L_MAT = "Mat. Value (\u20b9)"
L_QTY = "Quantity (total)"
L_GST_TDS = "GST TDS (\u20b9)"
L_SD_EXP = "SD Expected (\u20b9)"
L_SD_REC = "SD Received (\u20b9)"
L_FP_EXP = "FP Expected (\u20b9)"
L_FP_REC = "FP Received (\u20b9)"
L_RECV = "Total Received (\u20b9)"
L_RECEIV = "Total Receivables (\u20b9)"
L_STATUS = "Payment Status"
L_COLL = "Collection %"

# (label, number format, per-buyer spec, total-column spec)
PIVOT_SECTIONS = [
    ("BUYER OVERVIEW", "1F4E79", [
        ("Lots",                     "0",        "lots",       "sum"),
        ("Mat. Value (\u20b9)",      "#,##0",    "H",          "sum"),
        ("Total Received (\u20b9)", "#,##0",    "AC",         "sum"),
        ("Outstanding (\u20b9)",    "#,##0",    "AD",         "sum"),
        ("Payment Status",           "@",        "status",     "status"),
    ]),
    ("LOT INFORMATION", "C55A11", [
        ("Quantity (total)",         "#,##0.###", "A",         "sum"),
        ("Avg Rate (\u20b9)",       "#,##0",    "avg_rate",   "avg_rate"),
        ("Lots pending invoice",     "0",        "inv_pending", "sum"),
    ]),
    ("FINANCIALS", "548235", [
        ("GST @ 18% (\u20b9)",      "#,##0",    "I",          "sum"),
        ("Mat. Value + GST (\u20b9)", "#,##0",  "J",          "sum"),
        ("TCS @ 2% (\u20b9)",       "#,##0",    "K",          "sum"),
        ("TDS u/s 194(O) (\u20b9)", "#,##0.00", "L",          "sum"),
        ("Service Charge gross (\u20b9)", "#,##0.00", "M",    "sum"),
        ("TDS u/s 194(H) (\u20b9)", "#,##0.00", "N",          "sum"),
        ("Net Service Charge (\u20b9)", "#,##0.00", "O",      "sum"),
        ("Svc Charge to MSTC (\u20b9)", "#,##0", "P",         "sum"),
        ("GST TDS (\u20b9)",        "#,##0",    "R",          "sum"),
        ("GST TDS % (effective)",    "0.00%",    "gst_tds_pct", "gst_tds_pct"),
        ("Total Receivables (\u20b9)", "#,##0", "S",          "sum"),
    ]),
    ("SECURITY DEPOSIT", "7030A0", [
        ("SD Expected (\u20b9)",    "#,##0",    "T",          "sum"),
        ("SD Received (\u20b9)",    "#,##0",    "U",          "sum"),
        ("SD Outstanding (\u20b9)", "#,##0",    "sd_out",     "sd_out"),
    ]),
    ("FINAL PAYMENT", "00838F", [
        ("FP Expected (\u20b9)",    "#,##0",    "W",          "sum"),
        ("FP Received (\u20b9)",    "#,##0",    "X",          "sum"),
        ("FP Outstanding (\u20b9)", "#,##0",    "fp_out",     "fp_out"),
    ]),
    ("LPP", "C00000", [
        ("LPP Expected (\u20b9)",   "#,##0",    "Z",          "sum"),
        ("LPP Received (\u20b9)",   "#,##0",    "AA",         "sum"),
    ]),
    ("DOCUMENT", "BF8F00", [
        ("Invoices raised",          "0",        "inv_raised", "sum"),
        ("SAP Docs posted",          "0",        "sap_posted", "sum"),
    ]),
    ("RECOVERY", "A51E4D", [
        ("Collection %",             "0.0%",     "collection", "collection"),
    ]),
]


def buyer_formula(b, spec):
    """The formula body (no leading '=') for one buyer-driven figure.

    `b` is the address of the cell holding the buyer name, so the same body works
    whether buyers run across the columns or down the rows.
    """
    if spec == "lots":
        return _countif(b)
    if spec == "status":
        return f'IF({_sumif(b, "AD")}<=0,"SETTLED","OUTSTANDING")'
    if spec == "avg_rate":
        return f'IFERROR({_sumif(b, "H")}/{_sumif(b, "A")},"")'
    if spec == "inv_pending":
        return _countblank(b, "AE")
    if spec == "gst_tds_pct":
        return f'IFERROR({_sumif(b, "R")}/{_sumif(b, "H")},"")'
    if spec == "sd_out":
        return f"{_sumif(b, 'T')}-{_sumif(b, 'U')}"
    if spec == "fp_out":
        return f"{_sumif(b, 'W')}-{_sumif(b, 'X')}"
    if spec == "collection":
        return f'IFERROR({_sumif(b, "AC")}/{_sumif(b, "S")},"")'
    if spec == "inv_raised":
        return f"{_countif(b)}-{_countblank(b, 'AE')}"
    if spec == "sap_posted":
        return f"{_countif(b)}-{_countblank(b, 'AF')}"
    return _sumif(b, spec)


def build_buyer_pivot(wb, buyers):
    """Buyer names across the top, every detail as a row, one column per buyer."""
    ws = wb.create_sheet(BUYER_PIVOT, 1)
    names = list(buyers)
    nb = len(names)
    total_col = get_column_letter(2 + nb)

    ws.sheet_properties.tabColor = "7030A0"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.sheet_format.outlineLevelRow = 1
    ws.sheet_format.outlineLevelCol = 0
    ws.column_dimensions["A"].width = 32
    for i in range(nb):
        ws.column_dimensions[get_column_letter(2 + i)].width = 16
    ws.column_dimensions[total_col].width = 18

    ws.merge_cells(f"A1:{total_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER PIVOT   (buyers \u2192 columns  |  every detail "
               "\u2193 rows)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{total_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  one column per buyer, every figure a live SUMIF / COUNTIF against "
               "'Final Calculation Sheet' - change a buyer name in row 3 and its whole column follows   "
               "\u2022   filter with the \u25bc arrow on the FIELD column in row 3   \u2022   the \u2212 / + "
               "in the left margin folds a whole detail band   \u2022   the last column totals every buyer")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 32

    # header row ------------------------------------------------------------- #
    hdr = 3
    c = ws.cell(row=hdr, column=1, value="FIELD  \u25b8   |   BUYER \u2192")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i, buyer in enumerate(names):
        accent, tint, pale = THEMES[i % len(THEMES)]
        cc = ws.cell(row=hdr, column=2 + i, value=buyer)
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.fill = fill(accent)
        cc.alignment = CENTER
    cc = ws.cell(row=hdr, column=2 + nb, value="TOTAL \u2014 ALL BUYERS")
    cc.font = Font(bold=True, size=10, color="FFFFFF")
    cc.fill = fill("1F3864")
    cc.alignment = CENTER
    for col in range(1, 3 + nb):
        ws.cell(row=hdr, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 66
    ws.freeze_panes = f"B{hdr + 1}"

    # row map (needed for the derived rows that point at other rows) --------- #
    row_of = {}
    r = hdr + 1
    layout = []
    for sname, colour, fields in PIVOT_SECTIONS:
        layout.append(("section", sname, colour, r))
        r += 1
        for label, fmt, spec, tspec in fields:
            layout.append(("field", label, (fmt, spec, tspec), r))
            row_of[label] = r
            r += 1
    last_row = r - 1

    def cell_of(buyer_col, label):
        return f"{buyer_col}{row_of[label]}"

    for kind, name, extra, rr in layout:
        if kind == "section":
            for col in range(1, 3 + nb):
                cc = ws.cell(row=rr, column=col)
                cc.fill = fill(extra)
                cc.border = BOX
            cc = ws.cell(row=rr, column=1, value=f"  \u25b8 {name}")
            cc.font = Font(bold=True, size=10, color="FFFFFF")
            cc.alignment = LEFT
            ws.row_dimensions[rr].height = 17
            continue

        label, (fmt, spec, tspec) = name, extra
        a = ws.cell(row=rr, column=1, value=f"      {label}")
        a.font = Font(size=10, color="333333")
        a.alignment = LEFT
        for i in range(nb):
            col = get_column_letter(2 + i)
            b = f"{col}${hdr}"
            accent, tint, pale = THEMES[i % len(THEMES)]
            cc = ws.cell(row=rr, column=2 + i)
            if spec == "lots":
                cc.value = f"={_countif(b)}"
            elif spec == "status":
                cc.value = f'=IF({_sumif(b, "AD")}<=0,"SETTLED","OUTSTANDING")'
            elif spec == "avg_rate":
                cc.value = f'=IFERROR({_sumif(b, "H")}/{_sumif(b, "A")},"")'
            elif spec == "inv_pending":
                cc.value = f"={_countblank(b, 'AE')}"
            elif spec == "gst_tds_pct":
                cc.value = f'=IFERROR({_sumif(b, "R")}/{_sumif(b, "H")},"")'
            elif spec == "sd_out":
                cc.value = f"={_sumif(b, 'T')}-{_sumif(b, 'U')}"
            elif spec == "fp_out":
                cc.value = f"={_sumif(b, 'W')}-{_sumif(b, 'X')}"
            elif spec == "collection":
                cc.value = f'=IFERROR({_sumif(b, "AC")}/{_sumif(b, "S")},"")'
            elif spec == "inv_raised":
                cc.value = f"={_countif(b)}-{_countblank(b, 'AE')}"
            elif spec == "sap_posted":
                cc.value = f"={_countif(b)}-{_countblank(b, 'AF')}"
            else:
                cc.value = f"={_sumif(b, spec)}"
            cc.number_format = fmt
            cc.font = Font(size=10, bold=(spec == "status"))
            cc.alignment = LEFT if fmt == "@" else RIGHT
            cc.fill = fill(pale if i % 2 else tint)
            cc.border = BOX

        t = ws.cell(row=rr, column=2 + nb)
        first, lastm = get_column_letter(2), get_column_letter(1 + nb)
        if tspec == "sum":
            t.value = f"=SUM({first}{rr}:{lastm}{rr})"
        elif tspec == "status":
            t.value = f'=IF({total_col}{row_of[L_OUT]}<=0,"SETTLED","OUTSTANDING")'
        elif tspec == "avg_rate":
            t.value = (f'=IFERROR({total_col}{row_of[L_MAT]}/'
                       f'{total_col}{row_of[L_QTY]},"")')
        elif tspec == "gst_tds_pct":
            t.value = (f'=IFERROR({total_col}{row_of[L_GST_TDS]}/'
                       f'{total_col}{row_of[L_MAT]},"")')
        elif tspec == "sd_out":
            t.value = (f'={total_col}{row_of[L_SD_EXP]}-'
                       f'{total_col}{row_of[L_SD_REC]}')
        elif tspec == "fp_out":
            t.value = (f'={total_col}{row_of[L_FP_EXP]}-'
                       f'{total_col}{row_of[L_FP_REC]}')
        elif tspec == "collection":
            t.value = (f'=IFERROR({total_col}{row_of[L_RECV]}/'
                       f'{total_col}{row_of[L_RECEIV]},"")')
        t.number_format = fmt
        t.font = Font(bold=True, size=10, color="1F3864")
        t.alignment = LEFT if fmt == "@" else RIGHT
        t.fill = fill("DDEBF7")
        t.border = Border(left=Side(style="thin", color="1F4E79"), right=THIN, top=THIN, bottom=THIN)
        for col in range(1, 2 + nb):
            ws.cell(row=rr, column=col).border = BOX
        ws.row_dimensions[rr].outlineLevel = 1
        ws.row_dimensions[rr].height = 15

    ws.auto_filter.ref = f"A{hdr}:{total_col}{last_row}"

    def row_range(label, first_col="B"):
        rr = row_of[label]
        return f"{first_col}{rr}:{total_col}{rr}"

    for rule in (CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                            fill=fill("FFC7CE")),
                 CellIsRule(operator="lessThanOrEqual", formula=["0"], font=Font(bold=True, color="006100"))):
        ws.conditional_formatting.add(row_range("Outstanding (\u20b9)"), rule)
    for rule in (CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FFE699")),):
        ws.conditional_formatting.add(row_range("Lots pending invoice"), rule)
    for text, font, bg in (("SETTLED", Font(bold=True, color="006100"), "C6EFCE"),
                           ("OUTSTANDING", Font(bold=True, color="9C0006"), "FFC7CE")):
        ws.conditional_formatting.add(
            row_range("Payment Status"),
            CellIsRule(operator="equal", formula=[f'"{text}"'], font=font, fill=fill(bg)))
    ws.conditional_formatting.add(
        row_range("Collection %"),
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="num", mid_value=0.9, mid_color="FFEB84",
                       end_type="num", end_value=1, end_color="63BE7B"))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_cols = "A:A"



# --------------------------------------------------------------------------- #
# sheet 7 - 'Final Calculation Sheet' turned on its side: labels down the rows,
#           one column per lot, exactly the same fields in the same order
# --------------------------------------------------------------------------- #
def transpose_labels(src_ws):
    """(source column letter, row label) for every column of the source header.

    The source repeats some captions ('Date of Receipt' heads both the SD and the
    FP date column), so repeated captions get their source column appended.
    """
    raw = []
    for c in src_ws[3]:
        if c.value in (None, ""):
            continue
        raw.append((c.column_letter, " ".join(str(c.value).split()).strip()))
    seen = {}
    for _letter, label in raw:
        seen[label] = seen.get(label, 0) + 1
    return [(letter, label if seen[label] == 1 else f"{label} (col {letter})")
            for letter, label in raw]


def lot_rows(src_ws):
    """source rows that hold a lot (a lot number in column F)."""
    return [r for r in range(FIRST_DATA_ROW, src_ws.max_row + 1)
            if src_ws[f"F{r}"].value not in (None, "")]


def build_transposed(wb, sheet_name=None, *, tab="808080", lot_groups=False,
                     filter_mode="wide", note=""):
    """A mirror image of 'Final Calculation Sheet': labels down, values across.

    lot_groups   order the lot columns by auction then buyer and give every buyer
                 and every auction its own fold button (column outline)
    filter_mode  "wide"  = AutoFilter over the whole block (an arrow per column)
                 "label" = AutoFilter on column A only, so there is ONE button
                 None    = no filter buttons
    """
    src_ws = wb[SRC]
    fields = transpose_labels(src_ws)
    rows = lot_rows(src_ws)
    total_src = rows[-1] + 1                     # the source's own SUM row
    ws = wb.create_sheet(sheet_name or TRANSPOSED, 1)

    # ---- column plan (sheet order) ---------------------------------------- #
    plan, buyer_spans, auction_spans = [], [], []
    if lot_groups:
        ordered = sorted(rows, key=lambda r: (src_ws[f"D{r}"].value,
                                              str(src_ws[f"G{r}"].value).strip(),
                                              src_ws[f"F{r}"].value))
        runs = []
        for r in ordered:
            key = (src_ws[f"D{r}"].value, str(src_ws[f"G{r}"].value).strip())
            if runs and runs[-1][0] == key:
                runs[-1][1].append(r)
            else:
                runs.append((key, [r]))
        last_of_bid = {}
        for i, ((bid, _b), _rs) in enumerate(runs):
            last_of_bid[bid] = i
        a_start = None
        for i, ((bid, buyer), rs) in enumerate(runs):
            if a_start is None:
                a_start = len(plan)
            first = len(plan)
            for r in rs:
                plan.append({"t": "lot", "srow": r, "bid": bid, "buyer": buyer})
            buyer_spans.append((first, len(plan) - 1,
                                f"{buyer}  ({len(rs)})" if len(rs) > 1 else buyer, buyer))
            if i == len(runs) - 1:
                auction_spans.append((a_start, len(plan) - 1, f"AUCTION {bid}", bid))
            elif last_of_bid[bid] == i:            # last buyer of this auction
                plan.append({"t": "sep", "level": 0})
                auction_spans.append((a_start, len(plan) - 2, f"AUCTION {bid}", bid))
                a_start = None
            else:
                plan.append({"t": "sep", "level": 1})
    else:
        for r in rows:
            plan.append({"t": "lot", "srow": r, "bid": src_ws[f"D{r}"].value,
                         "buyer": str(src_ws[f"G{r}"].value).strip()})
    plan.append({"t": "total"})
    for i, c in enumerate(plan):
        c["letter"] = get_column_letter(2 + i)
    last_col = plan[-1]["letter"]
    lots = [c for c in plan if c["t"] == "lot"]

    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    if lot_groups:
        ws.sheet_properties.outlinePr.summaryBelow = False
        ws.sheet_properties.outlinePr.summaryRight = True
    ws.sheet_format.outlineLevelRow = 0
    ws.sheet_format.outlineLevelCol = 2 if lot_groups else 0
    ws.column_dimensions["A"].width = 34
    for c in plan:
        cd = ws.column_dimensions[c["letter"]]
        cd.width = 2.6 if c["t"] == "sep" else 14.5
        if c["t"] == "lot" and lot_groups:
            cd.outlineLevel = 2
        elif c["t"] == "sep":
            cd.outlineLevel = c["level"]

    # ---- colour maps ------------------------------------------------------- #
    buyers_seen, auctions_seen = [], []
    for c in lots:
        if c["buyer"] not in buyers_seen:
            buyers_seen.append(c["buyer"])
        if c["bid"] not in auctions_seen:
            auctions_seen.append(c["bid"])
    buyer_accent = {b: THEMES[i % len(THEMES)][0] for i, b in enumerate(buyers_seen)}
    buyer_tint = {b: THEMES[i % len(THEMES)][1] for i, b in enumerate(buyers_seen)}
    auction_colour = {b: THEMES[i % len(THEMES)][0] for i, b in enumerate(auctions_seen)}

    # ---- title + note ------------------------------------------------------ #
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  FINAL CALCULATION SHEET, TRANSPOSED   "
               "(field labels \u2193 rows  |  values \u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = note or (
        f"HOW TO USE  \u25b6  a one-for-one copy of '{SRC}' turned on its side: the "
        f"{len(fields)} field captions of row 3 run down column A in the same order, one column per "
        f"lot, and the last column is the source's own total row (row {total_src})   \u2022   every cell "
        "is a live link, so the two sheets can never disagree   \u2022   the header colour of a lot "
        "column is its auction: " + "  \u2022  ".join(
            f"{colour} = bid sheet {bid}" for bid, colour in auction_colour.items()) +
        "   \u2022   filter with the \u25bc arrow on the FIELD column")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 32

    # ---- bands (buyer / auction) + header row ------------------------------ #
    r = 3
    if lot_groups:
        for band_row, spans, colour_map, height, size in (
                (r, auction_spans, auction_colour, 16, 9),
                (r + 1, buyer_spans, buyer_accent, 46, 8)):
            for first, last, text, key in spans:
                c1, c2 = plan[first]["letter"], plan[last]["letter"]
                if c1 != c2:
                    ws.merge_cells(f"{c1}{band_row}:{c2}{band_row}")
                cc = ws[f"{c1}{band_row}"]
                cc.value = text
                cc.font = Font(bold=True, size=size, color="FFFFFF")
                # wrap: a single-lot buyer has only one 14.5-wide column to fit in
                cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                for k in range(first, last + 1):
                    x = ws[f"{plan[k]['letter']}{band_row}"]
                    x.fill = fill(colour_map[key])
                    x.border = BOX
            lab = ws.cell(row=band_row, column=1,
                          value="AUCTION  \u2192" if band_row == r else "BUYER  \u2192")
            lab.font = Font(bold=True, size=9, color="FFFFFF")
            lab.fill = fill("404040")
            lab.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[band_row].height = height
        r += 2
    hdr = r

    c = ws.cell(row=hdr, column=1, value="FIELD (row 3 of the source)  \u25b8  |  LOT \u2192")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    for cc in plan:
        x = ws[f"{cc['letter']}{hdr}"]
        if cc["t"] == "sep":
            x.fill = fill("FFFFFF")
            continue
        if cc["t"] == "total":
            x.value = f"TOTAL (src row {total_src})"
            x.font = Font(bold=True, size=10, color="FFFFFF")
            x.fill = fill("1F3864")
            x.alignment = CENTER
        else:
            x.value = src("F", cc["srow"])
            x.number_format = KIND_FMT["num0"]
            x.font = Font(bold=True, size=10, color="FFFFFF")
            x.fill = fill(auction_colour[cc["bid"]])
            x.alignment = CENTER
        x.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 30
    ws.freeze_panes = f"B{hdr + 1}"

    # ---- one row per source field ------------------------------------------ #
    r = hdr + 1
    for k, (scol, label) in enumerate(fields):
        a = ws.cell(row=r, column=1, value=label)
        a.font = Font(bold=True, size=10, color="1F3864")
        a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        a.fill = fill("DDEBF7" if k % 2 else "F4F9FD")
        a.border = BOX
        fmt = src_ws[f"{scol}{rows[0]}"].number_format or "General"
        for i, cc in enumerate(lots):
            x = ws[f"{cc['letter']}{r}"]
            x.value = src(scol, cc["srow"])
            x.number_format = fmt
            x.font = Font(size=10, color="333333")
            x.alignment = LEFT if fmt in ("General", "@") else RIGHT
            if lot_groups:
                x.fill = fill(buyer_tint[cc["buyer"]]) if i % 2 == 0 else fill("F7F7F7")
            elif i % 2:
                x.fill = fill("F7F7F7")
            x.border = BOX
        t = ws[f"{plan[-1]['letter']}{r}"]
        t.value = src(scol, total_src)
        t.number_format = fmt
        t.font = Font(bold=True, size=10, color="1F3864")
        t.alignment = LEFT if fmt in ("General", "@") else RIGHT
        t.fill = fill("DDEBF7")
        t.border = Border(left=Side(style="thin", color="1F4E79"), right=THIN,
                          top=THIN, bottom=THIN)
        ws.row_dimensions[r].height = 15
        r += 1
    last_row = r - 1

    if filter_mode == "wide":
        ws.auto_filter.ref = f"A{hdr}:{last_col}{last_row}"
    elif filter_mode == "label":
        # a one-column range = a single filter button on the label column, so the
        # lot headers do not sprout an arrow each
        ws.auto_filter.ref = f"A{hdr}:A{last_row}"

    out_row = hdr + 1 + [lbl for _l, lbl in fields].index("Outstanding")
    ws.conditional_formatting.add(
        f"{lots[0]['letter']}{out_row}:{lots[-1]['letter']}{out_row}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_cols = "A:A"


# --------------------------------------------------------------------------- #
# transposed sheet whose lot columns follow a buyer picked from a dropdown
# --------------------------------------------------------------------------- #
ADDITIVE_SRC_COLS = {"A", "H", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S",
                     "T", "U", "W", "X", "Z", "AA", "AC", "AD"}


def build_pick_buyer(wb, buyers) -> None:
    """Labels down the rows, one column per lot *of the buyer picked in B3*.

    Excel filters can only hide rows, so "show me one buyer's lots" is done with
    a dropdown plus INDEX/MATCH instead: a hidden helper column numbers each lot
    of the chosen buyer 1..n, and every lot column picks up the n-th one.
    """
    src_ws = wb[SRC]
    fields = transpose_labels(src_ws)
    rows = lot_rows(src_ws)
    first, last = rows[0], rows[-1]
    names = list(buyers)
    nb, width = len(names), max(len(v) for v in buyers.values())
    ws = wb.create_sheet(T_PICK_BUYER)
    total_idx = 2 + width                       # A + width lot columns -> total
    last_col = get_column_letter(total_idx)
    helper = total_idx + 2                      # hidden block: rank / row / list
    h_rank = get_column_letter(helper)
    h_row = get_column_letter(helper + 2)
    h_list = get_column_letter(helper + 4)

    ws.sheet_properties.tabColor = "203864"
    ws.sheet_view.showGridLines = False
    ws.sheet_format.outlineLevelRow = 0
    ws.sheet_format.outlineLevelCol = 0
    ws.column_dimensions["A"].width = 34
    for i in range(width):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14.5
    ws.column_dimensions[last_col].width = 17
    for L in (h_rank, get_column_letter(helper + 1), h_row,
              get_column_letter(helper + 3), h_list):
        ws.column_dimensions[L].width = 12
        ws.column_dimensions[L].hidden = True

    # ---- title ------------------------------------------------------------ #
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  TRANSPOSED, FILTERED BY BUYER   (pick a buyer "
               "\u2192 only its lots are shown)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  click the blue cell B3 and pick a buyer from the list - the lot "
               f"columns immediately re-point at that buyer's lots (up to {width}, the widest buyer here) "
               "and every figure follows   \u2022   the \u25bc on the FIELD column still filters which "
               "rows show   \u2022   nothing is typed in: each cell is an INDEX / MATCH into "
               f"'{SRC}', so the two can never disagree   \u2022   columns "
               f"{h_rank}, {h_row} and {h_list} are hidden helpers - the buyer list in {h_list} is "
               "rewritten whenever this sheet is rebuilt")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 44

    # ---- the buyer picker (row 3) ------------------------------------------ #
    c = ws["A3"]
    c.value = "PICK A BUYER  \u25b8"
    c.font = Font(bold=True, size=11, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("B3:D3")
    sel = ws["B3"]
    # open on the widest buyer so the sheet does not look half empty
    sel.value = max(buyers, key=lambda b: len(buyers[b]))
    sel.font = Font(bold=True, size=12, color="1F3864")
    sel.fill = fill("FFE699")
    sel.alignment = CENTER
    for L in ("B", "C", "D"):
        ws[f"{L}3"].border = Border(left=MED, right=MED, top=MED, bottom=MED)
    dv = DataValidation(type="list", formula1=f"${h_list}$4:${h_list}${3 + nb}",
                        allow_blank=False, showDropDown=False)
    dv.prompt = "Pick a buyer"
    dv.promptTitle = "Buyer"
    ws.add_data_validation(dv)
    dv.add(sel)
    ws.merge_cells(f"E3:{last_col}3")
    st = ws["E3"]
    grng = f"'{SRC}'!$G${first}:$G${last}"
    st.value = ('=IF($B$3="","\u25c0 pick a buyer in B3","showing "&COUNT($B$4:'
                f'${get_column_letter(1 + width)}$4)&" lot(s) of "&'
                f'COUNTIF({grng},$B$3)&" for "&$B$3)')
    st.font = Font(bold=True, size=10, color="1F3864")
    st.fill = fill("DDEBF7")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 26

    # ---- hidden helpers ---------------------------------------------------- #
    blank = '""'                      # an empty string, for inside a formula
    for sr in range(first, last + 1):
        ws[f"{h_rank}{sr}"] = (f"=IF('{SRC}'!$G{sr}=$B$3,"
                               f"COUNTIF('{SRC}'!$G${first}:$G{sr},$B$3),{blank})")
    ws[f"{h_rank}3"] = "rank of each source lot row inside the chosen buyer"
    for i in range(1, width + 1):
        ws[f"{h_row}{3 + i}"] = (f"=IFERROR(MATCH({i},${h_rank}${first}:${h_rank}${last},0)"
                                 f"+{first - 1},{blank})")
    ws[f"{h_row}2"] = "source row of each lot column"
    for i, name in enumerate(names):
        ws[f"{h_list}{4 + i}"] = name
    ws[f"{h_list}3"] = "buyer list for the B3 dropdown"
    for L in (h_rank, h_row, h_list):
        ws[f"{L}3"].font = Font(size=8, italic=True, color="808080")
    ws[f"{h_row}2"].font = Font(size=8, italic=True, color="808080")

    # ---- header row -------------------------------------------------------- #
    hdr = 4
    c = ws.cell(row=hdr, column=1, value="FIELD (row 3 of the source)  \u25b8  |  LOT \u2192")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i in range(width):
        L = get_column_letter(2 + i)
        x = ws.cell(row=hdr, column=2 + i,
                    value=(f"=IFERROR(INDEX('{SRC}'!$F${first}:$F${last},"
                           f"${h_row}{4 + i}-{first - 1}),{blank})"))
        x.number_format = KIND_FMT["num0"]
        x.font = Font(bold=True, size=10, color="FFFFFF")
        x.fill = fill(THEMES[i % len(THEMES)][0])
        x.alignment = CENTER
    t = ws.cell(row=hdr, column=total_idx, value='="TOTAL  "&$B$3')
    t.font = Font(bold=True, size=10, color="FFFFFF")
    t.fill = fill("1F3864")
    t.alignment = CENTER
    for col in range(1, total_idx + 1):
        ws.cell(row=hdr, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 30
    ws.freeze_panes = f"B{hdr + 1}"

    # ---- one row per source field ------------------------------------------ #
    r = hdr + 1
    for k, (scol, label) in enumerate(fields):
        a = ws.cell(row=r, column=1, value=label)
        a.font = Font(bold=True, size=10, color="1F3864")
        a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        a.fill = fill("DDEBF7" if k % 2 else "F4F9FD")
        fmt = src_ws[f"{scol}{first}"].number_format or "General"
        for i in range(width):
            L = get_column_letter(2 + i)
            idx = f"${h_row}{4 + i}-{first - 1}"
            rng = f"'{SRC}'!${scol}${first}:${scol}${last}"
            x = ws[f"{L}{r}"]
            x.value = (f'=IFERROR(IF(INDEX({rng},{idx})="","",INDEX({rng},{idx})),"")')
            x.number_format = fmt
            x.font = Font(size=10, color="333333")
            x.alignment = LEFT if fmt in ("General", "@") else RIGHT
            if i % 2:
                x.fill = fill("F7F7F7")
        tt = ws.cell(row=r, column=total_idx)
        if scol in ADDITIVE_SRC_COLS:
            tt.value = (f"=IF(COUNT({get_column_letter(2)}{r}:"
                        f"{get_column_letter(1 + width)}{r})=0,{blank},"
                        f"SUM({get_column_letter(2)}{r}:{get_column_letter(1 + width)}{r}))")
            tt.number_format = fmt
        else:
            tt.value = "\u2013"
            tt.number_format = "General"
        tt.font = Font(bold=True, size=10, color="1F3864")
        tt.alignment = LEFT if fmt in ("General", "@") else RIGHT
        tt.fill = fill("DDEBF7")
        for col in range(1, total_idx + 1):
            ws.cell(row=r, column=col).border = BOX
        ws.cell(row=r, column=total_idx).border = Border(
            left=Side(style="thin", color="1F4E79"), right=THIN, top=THIN, bottom=THIN)
        ws.row_dimensions[r].height = 15
        r += 1
    last_row = r - 1

    ws.auto_filter.ref = f"A{hdr}:A{last_row}"
    out_row = hdr + 1 + [lbl for _l, lbl in fields].index("Outstanding")
    ws.conditional_formatting.add(
        f"{get_column_letter(2)}{out_row}:{get_column_letter(1 + width)}{out_row}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))
    ws.conditional_formatting.add(
        f"{get_column_letter(2)}{out_row}:{get_column_letter(1 + width)}{out_row}",
        CellIsRule(operator="lessThanOrEqual", formula=["0"], font=Font(color="006100")))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_cols = "A:A"


# --------------------------------------------------------------------------- #
# one row per buyer, every detail across - filter the Buyer column
# --------------------------------------------------------------------------- #
def build_buyer_rows(wb, buyers) -> None:
    names = list(buyers)
    nb = len(names)
    fields = [(label, fmt, spec, tspec)
              for _s, _c, fl in PIVOT_SECTIONS for (label, fmt, spec, tspec) in fl]
    nf = len(fields)
    ws = wb.create_sheet(BUYER_ROWS)
    last_col = get_column_letter(1 + nf)
    col_of = {label: get_column_letter(2 + i) for i, (label, _f, _s, _t) in enumerate(fields)}

    ws.sheet_properties.tabColor = "9E480E"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_format.outlineLevelRow = 0
    ws.sheet_format.outlineLevelCol = 0
    ws.column_dimensions["A"].width = 34
    for i in range(nf):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13.5

    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  BUYER LIST   (one row per buyer  |  every detail "
               "\u2192 columns)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  one row per buyer - click the \u25bc on the BUYER column and tick "
               "the buyers you want; every other row hides   \u2022   every other column has its own "
               "\u25bc too, so you can also filter on Payment Status, Outstanding, Collection % ...   "
               "\u2022   the \u2211 TOTAL row uses SUBTOTAL, so it adds up only the buyers the filter "
               f"leaves visible   \u2022   every figure is a live SUMIF / COUNTIF against '{SRC}' "
               "driven by the buyer name in column A - rename a cell and its whole row follows")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 40

    # ---- section band + header --------------------------------------------- #
    band, hdr = 3, 4
    col = 2
    for sname, colour, fl in PIVOT_SECTIONS:
        c1 = get_column_letter(col)
        c2 = get_column_letter(col + len(fl) - 1)
        if c1 != c2:
            ws.merge_cells(f"{c1}{band}:{c2}{band}")
        x = ws[f"{c1}{band}"]
        x.value = sname
        x.font = Font(bold=True, size=9, color="FFFFFF")
        x.alignment = CENTER
        for k in range(len(fl)):
            cc = ws.cell(row=band, column=col + k)
            cc.fill = fill(colour)
            cc.border = BOX
        col += len(fl)
    lab = ws.cell(row=band, column=1, value="DETAIL BAND  \u2192")
    lab.font = Font(bold=True, size=9, color="FFFFFF")
    lab.fill = fill("404040")
    lab.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[band].height = 16

    c = ws.cell(row=hdr, column=1, value="BUYER  \u25b8  |  FIELD \u2192")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i, (label, _f, _s, _t) in enumerate(fields):
        cc = ws.cell(row=hdr, column=2 + i, value=label)
        cc.font = Font(bold=True, size=9, color="FFFFFF")
        cc.fill = fill("595959")
        cc.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True,
                                 textRotation=90)
    for col_i in range(1, 2 + nf):
        ws.cell(row=hdr, column=col_i).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 108
    ws.freeze_panes = f"B{hdr + 1}"

    # ---- one row per buyer -------------------------------------------------- #
    first_row = hdr + 1
    for bi, buyer in enumerate(names):
        rr = first_row + bi
        accent, tint, pale = THEMES[bi % len(THEMES)]
        a = ws.cell(row=rr, column=1, value=buyer)
        a.font = Font(bold=True, size=10, color=accent)
        a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for label, fmt, spec, _t in fields:
            cc = ws[f"{col_of[label]}{rr}"]
            cc.value = f"={buyer_formula(f'$A{rr}', spec)}"
            cc.number_format = fmt
            cc.font = Font(size=10, bold=(spec == "status"))
            cc.alignment = LEFT if fmt == "@" else RIGHT
        for col_i in range(1, 2 + nf):
            cc = ws.cell(row=rr, column=col_i)
            cc.fill = fill(pale if bi % 2 else tint)
            cc.border = BOX
        ws.row_dimensions[rr].height = 16
    last_buyer_row = first_row + nb - 1

    # ---- filter-aware total row -------------------------------------------- #
    tr = last_buyer_row + 1
    a = ws.cell(row=tr, column=1, value="\u2211  TOTAL \u2014 visible buyers")
    a.font = Font(bold=True, size=10, color="FFFFFF")
    a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for label, fmt, _spec, tspec in fields:
        L = col_of[label]
        cc = ws[f"{L}{tr}"]
        if tspec == "sum":
            cc.value = f"=SUBTOTAL(109,{L}{first_row}:{L}{last_buyer_row})"
        elif tspec == "status":
            cc.value = f'=IF({col_of[L_OUT]}{tr}<=0,"SETTLED","OUTSTANDING")'
        elif tspec == "avg_rate":
            cc.value = f'=IFERROR({col_of[L_MAT]}{tr}/{col_of[L_QTY]}{tr},"")'
        elif tspec == "gst_tds_pct":
            cc.value = f'=IFERROR({col_of[L_GST_TDS]}{tr}/{col_of[L_MAT]}{tr},"")'
        elif tspec == "sd_out":
            cc.value = f"={col_of[L_SD_EXP]}{tr}-{col_of[L_SD_REC]}{tr}"
        elif tspec == "fp_out":
            cc.value = f"={col_of[L_FP_EXP]}{tr}-{col_of[L_FP_REC]}{tr}"
        elif tspec == "collection":
            cc.value = f'=IFERROR({col_of[L_RECV]}{tr}/{col_of[L_RECEIV]}{tr},"")'
        cc.number_format = fmt
        cc.font = Font(bold=True, size=10, color="FFFFFF")
        cc.alignment = LEFT if fmt == "@" else RIGHT
    for col_i in range(1, 2 + nf):
        cc = ws.cell(row=tr, column=col_i)
        cc.fill = fill("1F3864")
        cc.border = BOX
    ws.row_dimensions[tr].height = 18

    # the filter covers the buyer rows only, so the total row is never hidden
    ws.auto_filter.ref = f"A{hdr}:{last_col}{last_buyer_row}"

    for rule in (CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                            fill=fill("FFC7CE")),
                 CellIsRule(operator="lessThanOrEqual", formula=["0"], font=Font(color="006100"))):
        ws.conditional_formatting.add(
            f"{col_of[L_OUT]}{first_row}:{col_of[L_OUT]}{tr}", rule)
    for text, fnt, bg in (("SETTLED", Font(bold=True, color="006100"), "C6EFCE"),
                          ("OUTSTANDING", Font(bold=True, color="9C0006"), "FFC7CE")):
        ws.conditional_formatting.add(
            f"{col_of[L_STATUS]}{first_row}:{col_of[L_STATUS]}{tr}",
            CellIsRule(operator="equal", formula=[f'"{text}"'], font=fnt, fill=fill(bg)))
    ws.conditional_formatting.add(
        f"{col_of[L_COLL]}{first_row}:{col_of[L_COLL]}{tr}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="num", mid_value=0.9, mid_color="FFEB84",
                       end_type="num", end_value=1, end_color="63BE7B"))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_cols = "A:A"


# --------------------------------------------------------------------------- #
# transposed sheet that filters the lot columns by the values of any one row
# --------------------------------------------------------------------------- #
def build_value_filter(wb, buyers) -> None:
    """Same transposed grid, but the lot columns are filtered by a row's values.

    Excel's own filter arrow can only ever hide rows, so "show me the lots whose
    Outstanding is > 0" cannot be done with an arrow on the transposed grid. This
    sheet does it with a query panel instead: pick a buyer, pick a ROW (field),
    a test and a value, and only the lots that pass fill the columns.
    """
    src_ws = wb[SRC]
    fields = transpose_labels(src_ws)
    rows = lot_rows(src_ws)
    first, last = rows[0], rows[-1]
    names = list(buyers)
    ncol = len(rows)
    last_scol = fields[-1][0]
    ws = wb.create_sheet(T_VALUE_FILTER)
    total_idx = 2 + ncol
    last_col = get_column_letter(total_idx)
    blank = '""'
    src2d = f"'{SRC}'!$A${first}:${last_scol}${last}"

    h = total_idx + 2
    h_labels = get_column_letter(h)
    h_buyers = get_column_letter(h + 2)
    h_tests = get_column_letter(h + 4)
    h_pos = get_column_letter(h + 6)
    h_pass = get_column_letter(h + 8)
    h_run = get_column_letter(h + 10)
    h_row = get_column_letter(h + 12)

    ws.sheet_properties.tabColor = "833C00"
    ws.sheet_view.showGridLines = False
    ws.sheet_format.outlineLevelRow = 0
    ws.sheet_format.outlineLevelCol = 0
    ws.column_dimensions["A"].width = 34
    for i in range(ncol):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14.5
    ws.column_dimensions[last_col].width = 17
    for k in range(h, h + 13):
        L = get_column_letter(k)
        ws.column_dimensions[L].width = 12
        ws.column_dimensions[L].hidden = True

    # ---- title + how-to --------------------------------------------------- #
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = ("MSTC LIMITED  \u2022  TRANSPOSED, FILTERED BY A ROW'S VALUES   "
               "(pick a row, a test and a value \u2192 only the matching lots show)")
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill("1F3864")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("HOW TO USE  \u25b6  an Excel filter arrow can only hide ROWS, so on this grid (lots "
               "across the columns) it cannot drop the lots that fail a test.  These four cells do it "
               "instead:  1 pick a buyer (or ALL),  2 pick the ROW whose values you want to test - "
               "Outstanding, Total Received, Buyer, Lot Name, anything -  3 pick the test,  4 type the "
               "value.  Only the lots that pass fill the columns, left to right; clear cell 2 back to "
               f"'{NO_FIELD}' to see them all.   \u2022   the \u25bc on the FIELD column still hides "
               "rows, and every figure is a live INDEX / MATCH into "
               f"'{SRC}', so nothing here can drift from the source.   \u2022   columns "
               f"{h_labels}..{h_row} are hidden helpers.")
    c.font = Font(size=9, italic=True, color="1F3864")
    c.fill = fill("FFF2CC")
    c.alignment = LEFTW
    ws.row_dimensions[2].height = 56

    # ---- the query panel (row 3) ------------------------------------------ #
    panel = (("1  \u25b8  BUYER", "B", f"${h_buyers}$4:${h_buyers}${4 + len(names)}"),
             ("2  \u25b8  ROW (field)", "D", f"${h_labels}$4:${h_labels}${3 + len(fields)}"),
             ("3  \u25b8  TEST", "F", f"${h_tests}$4:${h_tests}${3 + len(TESTS)}"))
    for text, L, listref in panel:
        lab = ws[f"{chr(ord(L) - 1)}3"]
        lab.value = text
        lab.font = Font(bold=True, size=10, color="FFFFFF")
        lab.fill = fill("404040")
        lab.alignment = Alignment(horizontal="right", vertical="center")
        cell = ws[f"{L}3"]
        cell.font = Font(bold=True, size=11, color="1F3864")
        cell.fill = fill("FFE699")
        cell.alignment = CENTER
        cell.border = Border(left=MED, right=MED, top=MED, bottom=MED)
        dv = DataValidation(type="list", formula1=listref, allow_blank=True,
                            showDropDown=False)
        ws.add_data_validation(dv)
        dv.add(cell)
    ws.column_dimensions["A"].width = 34
    ws["A3"].value = "FILTER  \u25b8"
    ws["A3"].font = Font(bold=True, size=11, color="FFFFFF")
    ws["A3"].fill = fill("1F3864")
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["B3"] = ALL_BUYERS
    ws["D3"] = NO_FIELD
    ws["F3"] = TESTS[0]
    lab = ws["G3"]
    lab.value = "4  \u25b8  VALUE"
    lab.font = Font(bold=True, size=10, color="FFFFFF")
    lab.fill = fill("404040")
    lab.alignment = Alignment(horizontal="right", vertical="center")
    val = ws["H3"]
    val.value = 0
    val.font = Font(bold=True, size=11, color="1F3864")
    val.fill = fill("FFE699")
    val.alignment = CENTER
    val.border = Border(left=MED, right=MED, top=MED, bottom=MED)
    ws.merge_cells(f"I3:{last_col}3")
    st = ws["I3"]
    grng = f"'{SRC}'!$F${first}:$F${last}"
    st.value = ('="showing "&COUNT($B$4:$' + last_col + '$4)&" of "&COUNT(' + grng +
                ')&" lots   \u2022   buyer: "&$B$3'
                f'&IF($D$3="{NO_FIELD}","","   \u2022   "&$D$3&" "&$F$3'
                f'&IF(OR($F$3="is blank",$F$3="is not blank"),""," "&$H$3))')
    st.font = Font(bold=True, size=10, color="1F3864")
    st.fill = fill("DDEBF7")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 26

    # ---- hidden helpers ---------------------------------------------------- #
    ws[f"{h_pos}2"] = f"=IFERROR(MATCH($D$3,${h_labels}$4:${h_labels}${3 + len(fields)},0),1)"
    ws[f"{h_pos}1"] = "source column of the chosen row"
    for i, (_scol, label) in enumerate(fields):
        ws[f"{h_labels}{4 + i}"] = label
    ws[f"{h_labels}3"] = "row (field) list for cell 2"
    ws[f"{h_buyers}4"] = ALL_BUYERS
    for i, name in enumerate(names):
        ws[f"{h_buyers}{5 + i}"] = name
    ws[f"{h_buyers}3"] = "buyer list for cell 1"
    for i, t in enumerate(TESTS):
        ws[f"{h_tests}{4 + i}"] = t
    ws[f"{h_tests}3"] = "test list for cell 3"
    for L in (h_pos, h_labels, h_buyers, h_tests):
        for rr in (1, 3):
            ws[f"{L}{rr}"].font = Font(size=8, italic=True, color="808080")

    def test_expr(r):
        v = f"INDEX({src2d},{r}-3,${h_pos}$2)"
        return (f'IF($F$3="is blank",IF({v}="",1,0),'
                f'IF($F$3="is not blank",IF({v}="",0,1),'
                f'IF($F$3="contains",IF(ISNUMBER(SEARCH($H$3,{v})),1,0),'
                f'IF($F$3="=",IF({v}=$H$3,1,0),'
                f'IF($F$3="<>",IF({v}=$H$3,0,1),'
                f'IF($F$3=">",IF(N({v})>N($H$3),1,0),'
                f'IF($F$3=">=",IF(N({v})>=N($H$3),1,0),'
                f'IF($F$3="<",IF(N({v})<N($H$3),1,0),'
                f'IF($F$3="<=",IF(N({v})<=N($H$3),1,0),0)))))))))')

    for sr in range(first, last + 1):
        ws[f"{h_pass}{sr}"] = (
            f'=IF(OR($B$3="{ALL_BUYERS}",\'{SRC}\'!$G{sr}=$B$3),'
            f'IF($D$3="{NO_FIELD}",1,{test_expr(sr)}),0)')
        ws[f"{h_run}{sr}"] = f'=IF({h_pass}{sr}=1,SUM(${h_pass}${first}:{h_pass}{sr}),{blank})'
    ws[f"{h_pass}3"] = "1 = this lot passes the filter"
    ws[f"{h_run}3"] = "running count of the lots that pass"
    ws[f"{h_row}3"] = "source row behind each lot column"
    for L in (h_pass, h_run, h_row):
        ws[f"{L}3"].font = Font(size=8, italic=True, color="808080")
    for i in range(1, ncol + 1):
        ws[f"{h_row}{3 + i}"] = (f"=IFERROR(MATCH({i},${h_run}${first}:${h_run}${last},0)"
                                 f"+{first - 1},{blank})")

    # ---- header row -------------------------------------------------------- #
    hdr = 4
    c = ws.cell(row=hdr, column=1, value="FIELD (row 3 of the source)  \u25b8  |  MATCHING LOT \u2192")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = fill("404040")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i in range(ncol):
        x = ws.cell(row=hdr, column=2 + i,
                    value=(f"=IFERROR(INDEX('{SRC}'!$F${first}:$F${last},"
                           f"${h_row}{4 + i}-{first - 1}),{blank})"))
        x.number_format = KIND_FMT["num0"]
        x.font = Font(bold=True, size=10, color="FFFFFF")
        x.fill = fill("833C00")
        x.alignment = CENTER
    t = ws.cell(row=hdr, column=total_idx, value='="TOTAL  (matching lots)"')
    t.font = Font(bold=True, size=10, color="FFFFFF")
    t.fill = fill("1F3864")
    t.alignment = CENTER
    for col in range(1, total_idx + 1):
        ws.cell(row=hdr, column=col).border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[hdr].height = 30
    ws.freeze_panes = f"B{hdr + 1}"

    # ---- one row per source field ------------------------------------------ #
    r = hdr + 1
    for k, (scol, label) in enumerate(fields):
        a = ws.cell(row=r, column=1, value=label)
        a.font = Font(bold=True, size=10, color="1F3864")
        a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        a.fill = fill("DDEBF7" if k % 2 else "F4F9FD")
        fmt = src_ws[f"{scol}{first}"].number_format or "General"
        for i in range(ncol):
            L = get_column_letter(2 + i)
            rng = f"'{SRC}'!${scol}${first}:${scol}${last}"
            idx = f"${h_row}{4 + i}-{first - 1}"
            x = ws[f"{L}{r}"]
            x.value = (f'=IFERROR(IF(INDEX({rng},{idx})="","",INDEX({rng},{idx})),{blank})')
            x.number_format = fmt
            x.font = Font(size=10, color="333333")
            x.alignment = LEFT if fmt in ("General", "@") else RIGHT
            if i % 2:
                x.fill = fill("F7F7F7")
        tt = ws.cell(row=r, column=total_idx)
        if scol in ADDITIVE_SRC_COLS:
            tt.value = (f"=IF(COUNT({get_column_letter(2)}{r}:"
                        f"{get_column_letter(1 + ncol)}{r})=0,{blank},"
                        f"SUM({get_column_letter(2)}{r}:{get_column_letter(1 + ncol)}{r}))")
            tt.number_format = fmt
        else:
            tt.value = "\u2013"
            tt.number_format = "General"
        tt.font = Font(bold=True, size=10, color="1F3864")
        tt.alignment = LEFT if fmt in ("General", "@") else RIGHT
        tt.fill = fill("DDEBF7")
        for col in range(1, total_idx + 1):
            ws.cell(row=r, column=col).border = BOX
        ws.cell(row=r, column=total_idx).border = Border(
            left=Side(style="thin", color="1F4E79"), right=THIN, top=THIN, bottom=THIN)
        ws.row_dimensions[r].height = 15
        r += 1
    last_row = r - 1

    ws.auto_filter.ref = f"A{hdr}:A{last_row}"
    out_row = hdr + 1 + [lbl for _l, lbl in fields].index("Outstanding")
    ws.conditional_formatting.add(
        f"{get_column_letter(2)}{out_row}:{get_column_letter(1 + ncol)}{out_row}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color="9C0006"),
                   fill=fill("FFC7CE")))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_cols = "A:A"


def main(path: str) -> None:
    wb = load_workbook(path)
    if SRC not in wb.sheetnames:
        raise SystemExit(f"source sheet '{SRC}' not found in {path}")
    for name in ALL_VIEWS + LEGACY_VIEWS:
        if name in wb.sheetnames:
            del wb[name]
    buyers = read_lots(wb[SRC])
    build_pick_buyer(wb, buyers)
    build_value_filter(wb, buyers)
    build_buyer_rows(wb, buyers)
    build_ledger_cols(wb, buyers)
    build_ledger(wb, buyers)
    build_collapsible_rows(wb, buyers)
    build_collapsible(wb, buyers, sheet_name=C_FIELD_FILTER, field_filter=True,
                      tab="F4B183")
    build_collapsible(wb, buyers)
    build_lot_vertical(wb, buyers)
    build_buyer_pivot(wb, buyers)
    # --- the four filter layouts, one sheet each (nothing existing is changed) -- #
    build_transposed(wb, T_BOTH, tab="255E91", lot_groups=True, filter_mode="label",
                     note="HOW TO USE  \u25b6  both at once: the \u25bc on the FIELD label cell (A5) filters "
                          "which rows show, and the \u2212 above the thin / wide dividers fold a buyer's or a "
                          "whole auction's lot columns away.")
    build_transposed(wb, T_LOT_FOLDS, tab="4472C4", lot_groups=True, filter_mode=None,
                     note="HOW TO USE  \u25b6  the lot columns are sorted by auction then buyer and grouped: "
                          "the \u2212 above a thin divider folds that BUYER's lots away, the \u2212 above a wide "
                          "divider folds a whole AUCTION.  The 1 / 2 buttons in the outline area collapse every "
                          "buyer or every auction at once.  No filter arrows here - this sheet is for folding.")
    build_transposed(wb, T_FIELD_FILTER, tab="7F7F7F", filter_mode="label",
                     note="HOW TO USE  \u25b6  ONE filter button, on the FIELD label cell at the top of column "
                          "A: click it and tick the fields you want to see - every other row hides.  The lot "
                          "headers carry no arrows.  Otherwise identical to 'Final Calc (Transposed)'.")
    build_transposed(wb, filter_mode="label")   # the filter button sits on the row labels
    # the builders insert at different places, so put the tabs in ALL_VIEWS order
    for i, name in enumerate([SRC] + list(ALL_VIEWS)):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))
    wb.active = wb.sheetnames.index(T_VALUE_FILTER)  # opens on the row-value filter
    # openpyxl serialises sheetFormatPr before the column outline levels, so it
    # never records outlineLevelCol; prime it so Excel draws the column group
    # buttons in the outline symbol area.
    for name in ALL_VIEWS:
        wb[name].column_dimensions.to_tree()
    wb.save(path)
    lots = sum(len(v) for v in buyers.values())
    print(f"{path}: rebuilt {len(ALL_VIEWS)} views for {len(buyers)} buyers / {lots} lots "
          f"(widest buyer {max(len(v) for v in buyers.values())} lots): " + ", ".join(ALL_VIEWS))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "06.09.2026.xlsx")
