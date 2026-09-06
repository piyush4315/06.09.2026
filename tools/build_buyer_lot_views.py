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
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

SRC = "Final Calculation Sheet"
COLLAPSIBLE = "Buyer Collapsible View"
LEDGER = "Lot Ledger (Filter)"
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
def build_collapsible(wb, buyers) -> None:
    n_lots = max(len(v) for v in buyers.values())
    ws = wb.create_sheet(COLLAPSIBLE)
    last_idx = 1 + n_lots + 1                       # A + lot columns + total
    last_col = get_column_letter(last_idx)
    total_col = last_col

    ws.sheet_properties.tabColor = "ED7D31"
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
    c.value = "MSTC LIMITED  \u2022  BUYER \u00d7 LOT COLLAPSIBLE VIEW"
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
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{COLLAPSIBLE}'!A{anchors[buyer]}",
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
    c.value = "MSTC LIMITED  \u2022  BUYER & LOT LEDGER  (filter + fold)"
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


def main(path: str) -> None:
    wb = load_workbook(path)
    if SRC not in wb.sheetnames:
        raise SystemExit(f"source sheet '{SRC}' not found in {path}")
    for name in (COLLAPSIBLE, LEDGER):
        if name in wb.sheetnames:
            del wb[name]
    buyers = read_lots(wb[SRC])
    build_collapsible(wb, buyers)
    build_ledger(wb, buyers)
    wb.active = wb.sheetnames.index(COLLAPSIBLE)
    # openpyxl serialises sheetFormatPr before the column outline levels, so it
    # never records outlineLevelCol; prime it so Excel draws the column group
    # buttons in the outline symbol area.
    for name in (COLLAPSIBLE, LEDGER):
        wb[name].column_dimensions.to_tree()
    wb.save(path)
    lots = sum(len(v) for v in buyers.values())
    print(f"{path}: rebuilt '{COLLAPSIBLE}' + '{LEDGER}' "
          f"({len(buyers)} buyers, {lots} lots, widest buyer {max(len(v) for v in buyers.values())} lots)")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "06.09.2026.xlsx")
