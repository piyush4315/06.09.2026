#!/usr/bin/env python3
"""Verify the generated views against the source sheet, using a real
spreadsheet formula engine (pip install formulas).

Ground truth = the engine's own calculation of 'Final Calculation Sheet'.
SUBTOTAL() is not implemented by the engine, so a throw-away copy of the
workbook is made where SUBTOTAL(109,..)->SUM(..) and SUBTOTAL(103,..)->COUNT(..):
  * a buyer summary row then sums exactly its own lot rows
  * a range that also contains those buyer rows (KPI band, grand total) comes
    out at exactly 2x the true total, because the buyer rows partition the lots
"""
import re
import shutil
import sys
from collections import OrderedDict

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, "tools")
from build_buyer_lot_views import (ADDITIVE_SRC_COLS, ALL_BUYERS, BUYER_PIVOT,  # noqa: E402
                                   BUYER_ROWS, C_FIELD_FILTER, COLLAPSIBLE, COLLAPSIBLE_ROWS,
                                   LEDGER, LEDGER_COLS, LIVE_SEARCH, LOT_VERTICAL, MAX_TERMS,
                                   NO_FIELD,
                                   PIVOT_SECTIONS, SEARCH_ALL, SEARCH_COLS, SECTIONS, SRC,
                                   TESTS, T_BOTH,
                                   T_FIELD_FILTER, T_LOT_FOLDS, T_PICK_BUYER, T_VALUE_FILTER,
                                   TRANSPOSED, lot_rows, plan_fields, read_lots,
                                   transpose_labels)

import formulas  # noqa: E402


def flat(v):
    try:
        v = v[0][0]
    except Exception:
        pass
    # the engine reports an empty cell as its own `empty` singleton
    return None if v is None or str(v) == "empty" else v


def calculate(path):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    out = {}
    for k, v in sol.items():
        m = re.match(r"^'\[[^\]]+\]([^']+)'!([A-Z]{1,3}\d+)$", k)
        if m:
            out[(m.group(1).upper(), m.group(2))] = flat(v.value)
    return out


def num(x, default=0.0):
    return default if x in (None, "", "#N/A") else float(x)


def main(path):
    wb = openpyxl.load_workbook(path)
    buyers = read_lots(wb[SRC])
    lots = sum(len(v) for v in buyers.values())

    # ---- throw-away copy with SUBTOTAL rewritten -------------------------- #
    tmp = "/tmp/verify_subtotal.xlsx"
    wb2 = openpyxl.load_workbook(path)
    n_sub = 0
    for name in (COLLAPSIBLE, LEDGER, BUYER_ROWS):
        for row in wb2[name].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=SUBTOTAL("):
                    c.value = c.value.replace("=SUBTOTAL(109,", "=SUM(").replace(
                        "=SUBTOTAL(103,", "=COUNT(")
                    n_sub += 1
    wb2.save(tmp)
    print(f"SUBTOTAL formulas rewritten for the check: {n_sub}")

    vals = calculate(tmp)
    fails, checks = [], 0

    def eq(where, got, want, tol=0.005):
        nonlocal checks
        checks += 1
        if isinstance(want, str) or isinstance(got, str):
            ok = str(got).strip() == str(want).strip()
        elif got in (None, "") and want in (None, "", 0):
            ok = want in (None, "") or abs(num(want)) < tol
        else:
            ok = abs(num(got) - num(want)) < tol
        if not ok:
            fails.append(f"{where}: got {got!r} want {want!r}")

    def V(sheet, addr):
        return vals.get((sheet.upper(), addr))

    def truth(col, srow):
        return V(SRC, f"{col}{srow}")

    # ================= LEDGER ============================================ #
    ws = wb[LEDGER]
    header = [c.value for c in ws[5]]
    letter_of = {}
    for i, h in enumerate(header):
        letter_of[h] = openpyxl.utils.get_column_letter(i + 1)
    src_of = {"Buyer": "G"}
    for _s, fields in SECTIONS:
        for label, scol, _k, _a in fields:
            src_of[label] = scol

    row_of_lot, buyer_rows = {}, {}
    for r in range(6, ws.max_row + 1):
        lvl = ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        a = ws[f"A{r}"].value
        if isinstance(a, str) and a.startswith("\u25bc"):
            buyer_rows[a[3:].strip()] = r
        elif lvl == 1:
            row_of_lot[int(ws[f"{letter_of['Src Row']}{r}"].value)] = r

    print(f"ledger: {len(buyer_rows)} buyer rows, {len(row_of_lot)} lot rows")
    assert len(row_of_lot) == lots, (len(row_of_lot), lots)

    for buyer, rows in buyers.items():
        for srow in rows:
            r = row_of_lot[srow]
            for label, letter in letter_of.items():
                if label == "Src Row":
                    eq(f"{LEDGER}!{letter}{r}", ws[f"{letter}{r}"].value, srow)
                    continue
                if label == "Payment Status":
                    out = num(truth("AD", srow), None)
                    want = "" if out is None else ("SETTLED" if out <= 0 else "OUTSTANDING")
                    eq(f"{LEDGER}!{letter}{r}", V(LEDGER, f"{letter}{r}"), want)
                    continue
                t = truth(src_of[label], srow)
                got = V(LEDGER, f"{letter}{r}")
                want = "" if t is None else t
                eq(f"{LEDGER}!{letter}{r}", got, want)

    # buyer summary rows + grand total + KPI band
    for label, scol, _kind, agg in [f for _s, fields in SECTIONS for f in fields]:
        if agg != "sum":
            continue
        letter = letter_of[label]
        total = 0.0
        for buyer, rows in buyers.items():
            want = sum(num(truth(scol, r)) for r in rows)
            total += want
            eq(f"{LEDGER} buyer row {buyer} {letter}", V(LEDGER, f"{letter}{buyer_rows[buyer]}"), want)
        # the grand-total range starts at the first LOT row, so the first buyer
        # summary row is outside it: rewritten SUM = 2*total - first buyer
        first = list(buyers)[0]
        t_first = sum(num(truth(scol, r)) for r in buyers[first])
        gt = ws.max_row
        eq(f"{LEDGER} grand total {letter}", V(LEDGER, f"{letter}{gt}"), 2 * total - t_first)

    eq(f"{LEDGER} KPI lots visible", V(LEDGER, "C3"), lots)
    kpi = {"G3": "H", "K3": "S", "O3": "U", "S3": "X", "W3": "AC", "AA3": "AD"}
    first = list(buyers)[0]
    for cell, col in kpi.items():
        want = sum(num(truth(col, r)) for rs in buyers.values() for r in rs)
        f1 = sum(num(truth(col, r)) for r in buyers[first])
        eq(f"{LEDGER} KPI {cell}", V(LEDGER, cell), 2 * want - f1)

    # ================= COLLAPSIBLE ======================================= #
    ws = wb[COLLAPSIBLE]
    n_lots = max(len(v) for v in buyers.values())
    total_letter = openpyxl.utils.get_column_letter(1 + n_lots + 1)
    # walk the blocks: banner rows are level 0 and hold a formula in A
    blocks = OrderedDict()
    cur = None
    for r in range(6, ws.max_row + 1):
        a = ws[f"A{r}"].value
        if isinstance(a, str) and a.startswith('="\u25bc'):
            cur = re.search(r'\u25bc  (.+?)   \u2022', a).group(1)
            blocks[cur] = {"banner": r, "fields": {}}
        elif cur and isinstance(a, str) and a.startswith("      "):
            blocks[cur]["fields"][a.strip()] = r
        elif isinstance(a, str) and a.startswith("\u25c8"):
            cur = None
    print(f"collapsible: {len(blocks)} blocks, "
          f"{sum(len(b['fields']) for b in blocks.values())} field rows")
    assert len(blocks) == len(buyers)

    for buyer, rows in buyers.items():
        b = blocks[buyer]
        for label, scol, kind, agg in [f for _s, fields in SECTIONS for f in fields]:
            r = b["fields"][label]
            for i, srow in enumerate(rows):
                letter = openpyxl.utils.get_column_letter(2 + i)
                if label == "Payment Status":
                    out = num(truth("AD", srow), None)
                    want = "" if out is None else ("SETTLED" if out <= 0 else "OUTSTANDING")
                else:
                    t = truth(scol, srow)
                    want = "" if t is None else t
                eq(f"{COLLAPSIBLE}!{letter}{r} [{buyer}/{label}]", V(COLLAPSIBLE, f"{letter}{r}"), want)
            if agg == "sum":
                want = sum(num(truth(scol, r2)) for r2 in rows)
                eq(f"{COLLAPSIBLE}!{total_letter}{r} [{buyer}/{label} total]",
                   V(COLLAPSIBLE, f"{total_letter}{r}"), want)
            else:
                eq(f"{COLLAPSIBLE}!{total_letter}{r} [{buyer}/{label} dash]",
                   V(COLLAPSIBLE, f"{total_letter}{r}"), "\u2013")
        banner = V(COLLAPSIBLE, f"A{b['banner']}")
        mv = sum(num(truth("H", r)) for r in rows)
        rc = sum(num(truth("AC", r)) for r in rows)
        od = sum(num(truth("AD", r)) for r in rows)
        eq(f"{COLLAPSIBLE} banner {buyer} lot count", f"{len(rows)} LOT(S)" in banner, True)
        eq(f"{COLLAPSIBLE} banner {buyer} mat value", f"\u20b9{mv:,.0f}" in banner, True)
        eq(f"{COLLAPSIBLE} banner {buyer} received", f"\u20b9{rc:,.0f}" in banner, True)
        eq(f"{COLLAPSIBLE} banner {buyer} outstanding", f"\u20b9{od:,.0f}" in banner, True)

    # buyer index at the bottom
    idx = {}
    started = False
    for r in range(6, ws.max_row + 1):
        a = ws[f"A{r}"].value
        if isinstance(a, str) and a.startswith("\u25c8"):
            started = True
            continue
        if started and isinstance(a, str) and a in buyers:
            idx[a] = r
    print(f"collapsible: index rows for {len(idx)} buyers")
    assert len(idx) == len(buyers)
    for buyer, rows in buyers.items():
        r = idx[buyer]
        eq(f"index {buyer} lots", V(COLLAPSIBLE, f"B{r}"), len(rows))
        for cell, col in (("C", "H"), ("D", "S"), ("E", "AC"), ("F", "AD")):
            eq(f"index {buyer} {cell}", V(COLLAPSIBLE, f"{cell}{r}"),
               sum(num(truth(col, x)) for x in rows))
        od = sum(num(truth("AD", x)) for x in rows)
        eq(f"index {buyer} status", V(COLLAPSIBLE, f"G{r}"),
           "SETTLED" if od <= 0 else "OUTSTANDING")
        link = ws[f"A{r}"].hyperlink
        checks += 1
        if link is None or link.location != f"'{COLLAPSIBLE}'!A{blocks[buyer]['banner']}":
            fails.append(f"index {buyer}: hyperlink {link} does not point at the block")

    tr = max(idx.values()) + 1
    eq("index ALL BUYERS lots", V(COLLAPSIBLE, f"B{tr}"), lots)
    for cell, col in (("C", "H"), ("D", "S"), ("E", "AC"), ("F", "AD")):
        eq(f"index ALL BUYERS {cell}", V(COLLAPSIBLE, f"{cell}{tr}"),
           sum(num(truth(col, r)) for rs in buyers.values() for r in rs))


    # ================= COLLAPSIBLE - LOTS DOWN ============================ #
    plan = plan_fields()
    ws = wb[COLLAPSIBLE_ROWS]
    letters2 = {str(c.value): c.column_letter for c in ws[4] if c.value}
    blocks2 = OrderedDict()
    cur = None
    for r in range(5, ws.max_row + 1):
        a = ws[f"A{r}"].value
        lvl = ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        if isinstance(a, str) and a.startswith('="\u25bc'):
            cur = re.search(r'\u25bc  (.+?)   \u2022', a).group(1)
            blocks2[cur] = {"banner": r, "lots": [], "totals": None}
        elif isinstance(a, str) and a.startswith("\u2211") and cur:
            blocks2[cur]["totals"] = r
        elif lvl == 1 and cur and isinstance(a, str):
            blocks2[cur]["lots"].append((r, int(re.search(r"\$[A-Z]{1,2}\$(\d+)", a).group(1))))
    print(f"collapsible (lots down): {len(blocks2)} blocks, "
          f"{sum(len(b['lots']) for b in blocks2.values())} lot rows")
    assert len(blocks2) == len(buyers)
    for buyer, rows in buyers.items():
        b = blocks2[buyer]
        checks += 1
        if [x[1] for x in b["lots"]] != rows:
            fails.append(f"{COLLAPSIBLE_ROWS}: {buyer} lots {b['lots']} != source rows {rows}")
        for r, srow in b["lots"]:
            for label, scol, _kind, _a in plan:
                letter = letters2[label]
                if label == "Payment Status":
                    o = num(truth("AD", srow), None)
                    want = "" if o is None else ("SETTLED" if o <= 0 else "OUTSTANDING")
                else:
                    t = truth(scol, srow)
                    want = "" if t is None else t
                eq(f"{COLLAPSIBLE_ROWS}!{letter}{r} [{buyer}/{label}]",
                   V(COLLAPSIBLE_ROWS, f"{letter}{r}"), want)
        for label, scol, _kind, agg in plan:
            if agg != "sum":
                continue
            letter = letters2[label]
            eq(f"{COLLAPSIBLE_ROWS} totals {buyer} {label}",
               V(COLLAPSIBLE_ROWS, f"{letter}{b['totals']}"),
               sum(num(truth(scol, x)) for x in rows))

    # ================= LEDGER - LOTS ACROSS =============================== #
    ws = wb[LEDGER_COLS]
    total_letter = openpyxl.utils.get_column_letter(1 + lots + 1)
    lot_cols = {}
    for c in ws[4]:
        if c.column > 1 and isinstance(c.value, str) and c.value.startswith("=IF("):
            lot_cols[c.column_letter] = int(re.search(r"\$[A-Z]{1,2}\$(\d+)", c.value).group(1))
    rows4 = {}
    for r in range(6, ws.max_row + 1):
        a = ws[f"A{r}"].value
        if isinstance(a, str) and a.startswith("      "):
            rows4[a.strip()] = r
    print(f"ledger (lots across): {len(lot_cols)} lot columns, {len(rows4)} field rows, "
          f"total column {total_letter}")
    checks += 1
    if sorted(lot_cols.values()) != sorted(x for rs in buyers.values() for x in rs):
        fails.append(f"{LEDGER_COLS}: lot columns {sorted(lot_cols.values())} do not match the source lots")
    assert len(rows4) == len(plan), (len(rows4), len(plan))
    for label, scol, _kind, agg in plan:
        r = rows4[label]
        for letter, srow in lot_cols.items():
            if label == "Payment Status":
                o = num(truth("AD", srow), None)
                want = "" if o is None else ("SETTLED" if o <= 0 else "OUTSTANDING")
            else:
                t = truth(scol, srow)
                want = "" if t is None else t
            eq(f"{LEDGER_COLS}!{letter}{r} [{label}]", V(LEDGER_COLS, f"{letter}{r}"), want)
        if agg == "sum":
            eq(f"{LEDGER_COLS} TOTAL {label}", V(LEDGER_COLS, f"{total_letter}{r}"),
               sum(num(truth(scol, x)) for rs in buyers.values() for x in rs))
        else:
            eq(f"{LEDGER_COLS} TOTAL dash {label}", V(LEDGER_COLS, f"{total_letter}{r}"), "\u2013")


    # ================= LOT-WISE VERTICAL ================================== #
    ws = wb[LOT_VERTICAL]
    blocks5 = OrderedDict()
    grand5 = {}
    cur_b = cur_lot = None
    mode = None
    for r in range(5, ws.max_row + 1):
        a = ws[f"A{r}"].value
        lvl = ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        if isinstance(a, str) and a.startswith('="\u25bc'):
            cur_b = re.search(r'\u25bc  (.+?)   \u2022', a).group(1)
            blocks5[cur_b] = {"lots": [], "totals": {}}
            mode = "buyer"
        elif isinstance(a, str) and a.startswith('="   \u25b8'):
            srow = int(re.search(r"\$F\$(\d+)", a).group(1))
            cur_lot = {}
            blocks5[cur_b]["lots"].append((srow, cur_lot))
            mode = "lot"
        elif isinstance(a, str) and a.startswith("   \u2211"):
            mode = "totals"
        elif isinstance(a, str) and a.startswith("\u2211\u2211"):
            mode = "grand"
        elif isinstance(a, str) and a.startswith("      \u25b8"):
            pass                                        # section heading
        elif lvl == 3 and cur_lot is not None:
            cur_lot[a.strip()] = r
        elif lvl == 2 and mode == "totals":
            blocks5[cur_b]["totals"][a.strip()] = r
        elif lvl == 1 and mode == "grand":
            grand5[a.strip()] = r
    print(f"lot-wise vertical: {len(blocks5)} buyers, "
          f"{sum(len(b['lots']) for b in blocks5.values())} lot blocks, "
          f"{len(grand5)} grand total lines")
    assert len(blocks5) == len(buyers)
    for buyer, rows in buyers.items():
        b = blocks5[buyer]
        checks += 1
        if [x[0] for x in b["lots"]] != rows:
            fails.append(f"{LOT_VERTICAL}: {buyer} lots {[x[0] for x in b['lots']]} != {rows}")
        for srow, fields in b["lots"]:
            checks += 1
            if len(fields) != len(plan):
                fails.append(f"{LOT_VERTICAL}: lot {srow} has {len(fields)} field rows, expected {len(plan)}")
            for label, scol, _kind, _a in plan:
                rr = fields.get(label)
                if rr is None:
                    fails.append(f"{LOT_VERTICAL}: lot {srow} missing row '{label}'")
                    continue
                if label == "Payment Status":
                    o = num(truth("AD", srow), None)
                    want = "" if o is None else ("SETTLED" if o <= 0 else "OUTSTANDING")
                else:
                    t = truth(scol, srow)
                    want = "" if t is None else t
                eq(f"{LOT_VERTICAL}!B{rr} [lot {srow}/{label}]", V(LOT_VERTICAL, f"B{rr}"), want)
        for label, scol, _kind, agg in plan:
            if agg != "sum":
                continue
            rr = b["totals"].get(label)
            if rr is None:
                fails.append(f"{LOT_VERTICAL}: {buyer} totals block missing '{label}'")
                continue
            eq(f"{LOT_VERTICAL} {buyer} totals {label}", V(LOT_VERTICAL, f"B{rr}"),
               sum(num(truth(scol, x)) for x in rows))
    for label, scol, _kind, agg in plan:
        if agg != "sum":
            continue
        rr = grand5.get(label)
        if rr is None:
            fails.append(f"{LOT_VERTICAL}: grand total missing '{label}'")
            continue
        eq(f"{LOT_VERTICAL} GRAND TOTAL {label}", V(LOT_VERTICAL, f"B{rr}"),
           sum(num(truth(scol, x)) for rs in buyers.values() for x in rs))


    # ================= BUYER PIVOT ======================================== #
    ws = wb[BUYER_PIVOT]
    total_letter = openpyxl.utils.get_column_letter(2 + len(buyers))
    col_of_buyer = {}
    for c in ws[3]:
        if c.column >= 2 and isinstance(c.value, str) and c.value in buyers:
            col_of_buyer[c.column_letter] = c.value
    print(f"buyer pivot: {len(col_of_buyer)} buyer columns, total column {total_letter}")
    checks += 1
    if len(col_of_buyer) != len(buyers):
        fails.append(f"{BUYER_PIVOT}: {len(col_of_buyer)} buyer columns, expected {len(buyers)}")

    def agg(spec, rows):
        """expected value of one pivot cell, straight off the source data"""
        if spec == "lots":
            return len(rows)
        if spec == "status":
            return "SETTLED" if sum(num(truth("AD", x)) for x in rows) <= 0 else "OUTSTANDING"
        if spec == "avg_rate":
            q = sum(num(truth("A", x)) for x in rows)
            return sum(num(truth("H", x)) for x in rows) / q if q else ""
        if spec == "inv_pending":
            return sum(1 for x in rows if truth("AE", x) is None)
        if spec == "inv_raised":
            return sum(1 for x in rows if truth("AE", x) is not None)
        if spec == "sap_posted":
            return sum(1 for x in rows if truth("AF", x) is not None)
        if spec == "gst_tds_pct":
            h = sum(num(truth("H", x)) for x in rows)
            return sum(num(truth("R", x)) for x in rows) / h if h else ""
        if spec == "sd_out":
            return sum(num(truth("T", x)) - num(truth("U", x)) for x in rows)
        if spec == "fp_out":
            return sum(num(truth("W", x)) - num(truth("X", x)) for x in rows)
        if spec == "collection":
            s_ = sum(num(truth("S", x)) for x in rows)
            return sum(num(truth("AC", x)) for x in rows) / s_ if s_ else ""
        return sum(num(truth(spec, x)) for x in rows)

    all_rows = [x for rs in buyers.values() for x in rs]
    seen = 0
    for sname, _colour, fields in PIVOT_SECTIONS:
        for label, _fmt, spec, tspec in fields:
            rr = None
            for r in range(4, ws.max_row + 1):
                if ws[f"A{r}"].value == f"      {label}":
                    rr = r
                    break
            if rr is None:
                fails.append(f"{BUYER_PIVOT}: row '{label}' not found")
                continue
            seen += 1
            for letter, buyer in col_of_buyer.items():
                eq(f"{BUYER_PIVOT}!{letter}{rr} [{buyer}/{label}]",
                   V(BUYER_PIVOT, f"{letter}{rr}"), agg(spec, buyers[buyer]))
            eq(f"{BUYER_PIVOT}!{total_letter}{rr} [TOTAL/{label}]",
               V(BUYER_PIVOT, f"{total_letter}{rr}"), agg(spec, all_rows))
    print(f"buyer pivot: {seen} detail rows checked against the source")
    checks += 1
    if seen != sum(len(f) for _n, _c, f in PIVOT_SECTIONS):
        fails.append(f"{BUYER_PIVOT}: only {seen} of "
                     f"{sum(len(f) for _n, _c, f in PIVOT_SECTIONS)} detail rows found")


    # ================= BUYER ROWS (FILTER) ================================= #
    ws = wb[BUYER_ROWS]
    bhdr = 4
    col_of = {}
    for c in ws[bhdr]:
        if c.column >= 2 and isinstance(c.value, str) and c.value.strip():
            col_of[c.value.strip()] = c.column_letter
    row_of_buyer = {}
    for r in range(bhdr + 1, ws.max_row + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str) and v in buyers:
            row_of_buyer[v] = r
    total_row = max(row_of_buyer.values()) + 1 if row_of_buyer else None
    print(f"buyer rows: {len(row_of_buyer)} buyers, {len(col_of)} detail columns, "
          f"total row {total_row}")
    checks += 3
    if len(row_of_buyer) != len(buyers):
        fails.append(f"{BUYER_ROWS}: {len(row_of_buyer)} buyer rows, expected {len(buyers)}")
    if len(col_of) != sum(len(f) for _n, _c, f in PIVOT_SECTIONS):
        fails.append(f"{BUYER_ROWS}: {len(col_of)} detail columns")
    if ws.auto_filter.ref != f"A{bhdr}:{get_column_letter(1 + len(col_of))}{total_row - 1}":
        fails.append(f"{BUYER_ROWS}: autofilter {ws.auto_filter.ref}, expected the buyer "
                     f"rows only (A{bhdr}:..{total_row - 1}) so the total row is never hidden")
    for _sname, _c, fl in PIVOT_SECTIONS:
        for label, _fmt, spec, _t in fl:
            if label not in col_of:
                fails.append(f"{BUYER_ROWS}: column '{label}' missing")
                continue
            L = col_of[label]
            for buyer, rr in row_of_buyer.items():
                eq(f"{BUYER_ROWS}!{L}{rr} [{buyer}/{label}]",
                   V(BUYER_ROWS, f"{L}{rr}"), agg(spec, buyers[buyer]))
            eq(f"{BUYER_ROWS}!{L}{total_row} [TOTAL/{label}]",
               V(BUYER_ROWS, f"{L}{total_row}"), agg(spec, all_rows))

    # ================= TRANSPOSED - PICK A BUYER =========================== #
    # the lot columns follow B3, so each buyer needs its own recalculation
    ws = wb[T_PICK_BUYER]
    src_ws = wb[SRC]
    fields8 = transpose_labels(src_ws)
    srows8 = lot_rows(src_ws)
    width = max(len(v) for v in buyers.values())
    prow_of = {}
    for r in range(5, ws.max_row + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str) and v.strip():
            prow_of[v.strip()] = r
    one_lot = min((b for b, v in buyers.items() if len(v) == 1), key=str)
    two_lot = min((b for b, v in buyers.items() if len(v) == 2), key=str)
    widest = max(buyers, key=lambda b: len(buyers[b]))
    for buyer in (widest, two_lot, one_lot):
        tmp = f"/tmp/verify_pick_{len(buyer)}.xlsx"
        shutil.copy(path, tmp)
        wb3 = openpyxl.load_workbook(tmp)
        wb3[T_PICK_BUYER]["B3"] = buyer
        wb3.save(tmp)
        vals3 = calculate(tmp)
        rows_b = buyers[buyer]

        def V3(addr, _v=vals3):
            return _v.get((T_PICK_BUYER.upper(), addr))
        print(f"pick a buyer = {buyer!r}: {len(rows_b)} lot(s) expected")
        for i in range(width):
            L = get_column_letter(2 + i)
            want = truth("F", rows_b[i]) if i < len(rows_b) else ""
            eq(f"{T_PICK_BUYER}!{L}4 [lot {i + 1} of {buyer}]", V3(f"{L}4"),
               "" if want is None else want)
        for scol, label in fields8:
            rr = prow_of.get(label)
            if rr is None:
                fails.append(f"{T_PICK_BUYER}: field row '{label}' missing")
                continue
            for i in range(width):
                L = get_column_letter(2 + i)
                t = truth(scol, rows_b[i]) if i < len(rows_b) else ""
                eq(f"{T_PICK_BUYER}!{L}{rr} [{buyer}/{label}/lot {i + 1}]",
                   V3(f"{L}{rr}"), "" if t is None else t)
            tot = get_column_letter(2 + width)
            if scol in ADDITIVE_SRC_COLS:
                # the total stays blank when every lot of this buyer is blank
                vals_row = [truth(scol, x) for x in rows_b]
                want = "" if all(v is None for v in vals_row) else \
                    sum(num(v) for v in vals_row)
                eq(f"{T_PICK_BUYER}!{tot}{rr} [{buyer}/{label}/TOTAL]", V3(f"{tot}{rr}"), want)
            else:
                eq(f"{T_PICK_BUYER}!{tot}{rr} [{buyer}/{label}/TOTAL]", V3(f"{tot}{rr}"),
                   "\u2013")
        eq(f"{T_PICK_BUYER}!E3 [{buyer}/status line]", V3("E3"),
           f"showing {len(rows_b)} lot(s) of {len(rows_b)} for {buyer}")

    # ================= TRANSPOSED - FILTER VALUES ========================== #
    # the lot columns follow the query panel in row 3, so the default panel is
    # checked from the main recalculation and each filter needs its own run
    ws = wb[T_VALUE_FILTER]
    fields9 = transpose_labels(wb[SRC])
    srows9 = lot_rows(wb[SRC])
    ncol9 = len(srows9)
    vrow_of = {}
    for r in range(5, ws.max_row + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str) and v.strip():
            vrow_of[v.strip()] = r
    row_of_lotno = {}
    for srow in srows9:
        row_of_lotno[num(truth("F", srow))] = srow
    checks += 4
    if len(row_of_lotno) != len(srows9):
        fails.append(f"{T_VALUE_FILTER}: lot numbers are not unique, cannot map columns")
    if ws.auto_filter.ref != f"A4:A{4 + len(fields9)}":
        fails.append(f"{T_VALUE_FILTER}: autofilter {ws.auto_filter.ref}, expected the "
                     f"label column only")
    dvs = {str(dv.sqref): dv.formula1 for dv in ws.data_validations.dataValidation}
    if sorted(dvs) != ["B3", "D3", "F3"]:
        fails.append(f"{T_VALUE_FILTER}: dropdowns on {sorted(dvs)}, expected B3, D3, F3")
    if ws["D3"].value != NO_FIELD or ws["B3"].value != ALL_BUYERS:
        fails.append(f"{T_VALUE_FILTER}: the panel does not start unfiltered "
                     f"({ws['B3'].value!r} / {ws['D3'].value!r})")

    def vcols(vals):
        """{source row: column letter} for the lots the panel currently lets through"""
        out = {}
        for i in range(ncol9):
            L = get_column_letter(2 + i)
            no = vals.get((T_VALUE_FILTER.upper(), f"{L}4"))
            if no not in (None, "", "empty"):
                out[row_of_lotno[num(no)]] = L
        return out

    default_cols = vcols(vals)
    print(f"filter values: default panel lets {len(default_cols)} of {ncol9} lots through")
    checks += 1
    if sorted(default_cols) != sorted(srows9):
        fails.append(f"{T_VALUE_FILTER}: the unfiltered panel shows "
                     f"{len(default_cols)} lots, expected all {ncol9}")
    for scol, label in fields9:
        rr = vrow_of.get(label)
        if rr is None:
            fails.append(f"{T_VALUE_FILTER}: field row '{label}' missing")
            continue
        for srow, L in default_cols.items():
            t = truth(scol, srow)
            eq(f"{T_VALUE_FILTER}!{L}{rr} [{label}/lot row {srow}]",
               V(T_VALUE_FILTER, f"{L}{rr}"), "" if t is None else t)
        tot = get_column_letter(2 + ncol9)
        if scol in ADDITIVE_SRC_COLS:
            eq(f"{T_VALUE_FILTER}!{tot}{rr} [{label}/TOTAL all lots]",
               V(T_VALUE_FILTER, f"{tot}{rr}"), sum(num(truth(scol, x)) for x in srows9))

    # three filtered runs: a numeric row test, a text row test, and both panels
    widest9 = max(buyers, key=lambda b: len(buyers[b]))
    cases = [("Outstanding", ">", 0, None),
             ("Lot Name", "contains", "COPPER", None),
             ("Outstanding", ">", 0, widest9)]
    for field9, test9, value9, buyer9 in cases:
        tmp = f"/tmp/verify_vf_{len(field9)}{len(str(value9))}{len(buyer9 or '')}.xlsx"
        shutil.copy(path, tmp)
        wb4 = openpyxl.load_workbook(tmp)
        p4 = wb4[T_VALUE_FILTER]
        p4["D3"], p4["F3"], p4["H3"] = field9, test9, value9
        if buyer9:
            p4["B3"] = buyer9
        wb4.save(tmp)
        vals4 = calculate(tmp)

        def V4(addr, _v=vals4):
            return _v.get((T_VALUE_FILTER.upper(), addr))
        src_col = dict((lbl, sc) for sc, lbl in fields9)[field9]
        want = []
        for srow in srows9:
            if buyer9 and str(truth("G", srow)).strip() != buyer9:
                continue
            v9 = truth(src_col, srow)
            if test9 == ">":
                ok = num(v9) > num(value9)
            elif test9 == "contains":
                ok = str(value9).upper() in str(v9 or "").upper()
            else:
                ok = False
            if ok:
                want.append(srow)
        got = vcols(vals4)
        print(f"filter values: {buyer9 or ALL_BUYERS} / {field9} {test9} {value9} "
              f"-> {len(got)} lot(s), expected {len(want)}")
        checks += 1
        if sorted(got) != sorted(want):
            fails.append(f"{T_VALUE_FILTER}: '{field9} {test9} {value9}'"
                         f"{' for ' + buyer9 if buyer9 else ''} let through "
                         f"{sorted(got)}, expected {sorted(want)}")
        for scol, label in fields9:
            rr = vrow_of[label]
            for srow, L in got.items():
                t = truth(scol, srow)
                eq(f"{T_VALUE_FILTER}!{L}{rr} [{label}/lot row {srow} filtered]",
                   V4(f"{L}{rr}"), "" if t is None else t)
        tot = get_column_letter(2 + ncol9)
        for scol, label in fields9:
            if scol not in ADDITIVE_SRC_COLS:
                continue
            rr = vrow_of[label]
            v9s = [truth(scol, x) for x in want]
            eq(f"{T_VALUE_FILTER}!{tot}{rr} [{label}/TOTAL filtered]", V4(f"{tot}{rr}"),
               "" if all(x is None for x in v9s) else sum(num(x) for x in v9s))
        eq(f"{T_VALUE_FILTER}!I3 [status line]", V4("I3"),
           f"showing {len(want)} of {ncol9} lots   \u2022   buyer: {buyer9 or ALL_BUYERS}"
           f"   \u2022   {field9} {test9}"
           + ("" if test9 in ("is blank", "is not blank") else f" {value9}"))

    # ================= LIVE SEARCH ========================================= #
    # The hits are grouped by the value typed in the box: one block per value
    # with its own total row, then a REMAINING block for the misses, then the
    # grand total. The rows therefore move as you type, so what gets checked is
    # the whole drawn table - which lot is on which row, every field on it, and
    # each group / remaining / grand total - recomputed here from the source.
    ws = wb[LIVE_SEARCH]
    src_ws10 = wb[SRC]
    label_of10 = {scol: lbl for scol, lbl in transpose_labels(src_ws10)}
    srows10 = lot_rows(src_ws10)
    n10 = len(srows10)
    heads10 = ["#"] + [label_of10[c] for c in SEARCH_COLS] + ["Payment Status"]
    # the builder puts the hidden helpers at ncol + 2; mirror that arithmetic
    hh10 = 1 + len(SEARCH_COLS) + 1 + 2
    L_HAY = get_column_letter(hh10)
    L_HIT = get_column_letter(hh10 + 2)
    L_GRP, L_RNK = get_column_letter(hh10 + 15), get_column_letter(hh10 + 16)
    L_OUT, L_SRC = get_column_letter(hh10 + 17), get_column_letter(hh10 + 18)
    L_CNT, L_ST, L_EN = (get_column_letter(hh10 + 19), get_column_letter(hh10 + 20),
                         get_column_letter(hh10 + 21))
    L_GD, L_ROL, L_BIG = (get_column_letter(hh10 + 22), get_column_letter(hh10 + 23),
                          get_column_letter(hh10 + 24))
    L_ROW, L_REM = get_column_letter(hh10 + 25), get_column_letter(hh10 + 26)
    L_MTF = get_column_letter(hh10 + 27)     # 1 on the TOTAL MATCHED row
    G_ALL10, G_REM10 = MAX_TERMS + 1, MAX_TERMS + 2
    first10, last_lot10 = 5, 4 + n10
    last10 = 4 + n10 + MAX_TERMS + 4
    checks += 6

    if [ws.cell(row=4, column=i).value for i in range(1, len(heads10) + 1)] != heads10:
        fails.append(f"{LIVE_SEARCH}: header row is not {heads10[:4]}...")
    if ws["B3"].value not in (None, ""):
        fails.append(f"{LIVE_SEARCH}: the search box does not start empty ({ws['B3'].value!r})")
    if ws["F3"].value != SEARCH_ALL:
        fails.append(f"{LIVE_SEARCH}: 'search in' starts at {ws['F3'].value!r}")
    if [str(dv.sqref) for dv in ws.data_validations.dataValidation] != ["F3"]:
        fails.append(f"{LIVE_SEARCH}: dropdowns on "
                     f"{[str(dv.sqref) for dv in ws.data_validations.dataValidation]}, expected F3")
    if ws.freeze_panes != "C5":
        fails.append(f"{LIVE_SEARCH}: freeze {ws.freeze_panes}, expected C5")
    if not str(ws[f"{L_HIT}5"].value or "").startswith('=IF($B$3="",1,IF(AND('):
        fails.append(f"{LIVE_SEARCH}: {L_HIT}5 is not the facet flag "
                     f"({str(ws[f'{L_HIT}5'].value)[:50]!r})")

    def parse10(text):
        """'a, b + c' -> [(a,1), (b,1), (c,2)]: ',' and '/' OR, '+' and '&' AND"""
        out, facet, cur = [], 1, ""
        for ch in text.replace("&", "+").replace("/", ",") + "+":
            if ch in ",+":
                if cur.strip():
                    out.append((cur.strip(), facet))
                if ch == "+":
                    facet += 1
                cur = ""
            else:
                cur += ch
        return out

    def slots10(text):
        """the box as the sheet splits it: one entry per slot, empties kept"""
        out, cur = [], ""
        for ch in text.replace("&", "+").replace("/", ","):
            if ch in ",+":
                out.append(cur.strip())
                cur = ""
            else:
                cur += ch
        out.append(cur.strip())
        return out[:MAX_TERMS]

    def layout10(text, hays):
        """which lots fall under which typed value, and the blocks they draw"""
        pairs = parse10(text)
        facets = sorted({f for _t, f in pairs})
        terms = slots10(text)
        memb = {g: [] for g in range(1, G_REM10 + 1)}
        for i, hay in enumerate(hays):
            hu = hay.upper()
            hit = all(any(t.upper() in hu for t, f in pairs if f == g) for g in facets)
            if text == "":
                memb[G_ALL10].append(i)
            elif not hit:
                memb[G_REM10].append(i)
            else:
                g = next((k + 1 for k, t in enumerate(terms) if t and t.upper() in hu), G_REM10)
                memb[g].append(i)
        cnt = {g: len(v) for g, v in memb.items()}
        st, en = {}, {}
        for g in range(1, G_ALL10 + 1):        # the blocks the search hit, end to end
            st[g] = first10 if g == 1 else (st[g - 1] if cnt[g - 1] == 0 else en[g - 1] + 1)
            en[g] = 0 if cnt[g] == 0 else st[g] + cnt[g]
        mt = st[G_ALL10] if cnt[G_ALL10] == 0 else en[G_ALL10] + 1
        st[G_REM10] = mt + 1                   # REMAINING starts under the matched total
        en[G_REM10] = 0 if cnt[G_REM10] == 0 else st[G_REM10] + cnt[G_REM10]
        grand = st[G_REM10] if cnt[G_REM10] == 0 else en[G_REM10] + 1
        return memb, cnt, st, en, mt, grand, terms

    def hay10(srow, sin):
        if sin == "Lot No.":
            cols = ("F",)
        elif sin == "Buyer":
            cols = ("G",)
        elif sin == "Lot Name":
            cols = ("B",)
        elif sin == "Bid Sheet":
            cols = ("D",)
        elif sin == "Unit":
            cols = ("E",)
        else:
            cols = ("F", "G", "B")
        return " ".join(str(truth(c, srow) or "") for c in cols)

    def check_live10(vals, tag, text, sin):
        S = LIVE_SEARCH.upper()
        hays = [hay10(r, sin) for r in srows10]
        memb, cnt, st, en, mt, grand, terms = layout10(text, hays)
        grp_of = {i: g for g, w in memb.items() for i in w}
        role_of = {en[g]: g for g in range(1, G_REM10 + 1) if cnt[g]}
        row_of = {}
        for g, who in memb.items():
            for k, i in enumerate(who):
                row_of[st[g] + k] = (i, k + 1, g)
        hits = sum(len(v) for g, v in memb.items() if g != G_REM10)
        eq(f"{LIVE_SEARCH}!H3 [matches] {tag}", vals.get((S, "H3")), f"{hits} of {n10}")
        eq(f"{LIVE_SEARCH}!{L_GD}4 [matched-total row] {tag}", vals.get((S, f"{L_GD}4")), mt)
        eq(f"{LIVE_SEARCH}!{L_GD}5 [grand-total row] {tag}", vals.get((S, f"{L_GD}5")), grand)
        for g in range(1, G_REM10 + 1):
            eq(f"{LIVE_SEARCH}!{L_CNT}{3 + g} [group {g} count] {tag}",
               vals.get((S, f"{L_CNT}{3 + g}")), cnt[g])
            eq(f"{LIVE_SEARCH}!{L_EN}{3 + g} [group {g} total row] {tag}",
               vals.get((S, f"{L_EN}{3 + g}")), en[g])
        for rr in range(first10, last10 + 1):
            want_role = role_of.get(rr, 0)
            want_big = 1 if rr == grand else 0
            want_mtf = 1 if rr == mt else 0
            here = row_of.get(rr)
            eq(f"{LIVE_SEARCH}!{L_ROL}{rr} [row role] {tag}", vals.get((S, f"{L_ROL}{rr}")),
               want_role)
            eq(f"{LIVE_SEARCH}!{L_BIG}{rr} [grand row] {tag}", vals.get((S, f"{L_BIG}{rr}")),
               want_big)
            eq(f"{LIVE_SEARCH}!{L_MTF}{rr} [matched row] {tag}", vals.get((S, f"{L_MTF}{rr}")),
               want_mtf)
            eq(f"{LIVE_SEARCH}!{L_ROW}{rr} [source row shown] {tag}",
               vals.get((S, f"{L_ROW}{rr}")), srows10[here[0]] if here else 0)
            eq(f"{LIVE_SEARCH}!{L_REM}{rr} [greyed] {tag}", vals.get((S, f"{L_REM}{rr}")),
               1 if (here and here[2] == G_REM10) or want_role == G_REM10 else 0)
            if here:
                i, rank, g = here
                srow = srows10[i]
                eq(f"{LIVE_SEARCH}!A{rr} [place in group] {tag}", vals.get((S, f"A{rr}")), rank)
                for scol in SEARCH_COLS:
                    L = get_column_letter(2 + SEARCH_COLS.index(scol))
                    t = truth(scol, srow)
                    eq(f"{LIVE_SEARCH}!{L}{rr} [{label_of10[scol]}/{srow}] {tag}",
                       vals.get((S, f"{L}{rr}")), "" if t is None else t)
                ad = truth("AD", srow)
                eq(f"{LIVE_SEARCH}!AI{rr} [status/{srow}] {tag}", vals.get((S, f"AI{rr}")),
                   "" if ad is None else ("SETTLED" if num(ad) <= 0 else "OUTSTANDING"))
            else:
                # a group / remaining / grand total row, or a row the table does
                # not reach - which the sheet leaves completely empty
                used = bool(want_role or want_big or want_mtf)
                if want_big:
                    who = list(range(n10))
                elif want_mtf:
                    who = [i for i in range(n10) if grp_of[i] != G_REM10]
                else:
                    who = memb.get(want_role, [])
                many = len(who) if (want_big or want_mtf) else cnt.get(want_role, 0)
                if not used:
                    label10 = ""
                elif want_big:
                    label10 = "GRAND TOTAL  \u2014  matching + remaining"
                elif want_mtf:
                    label10 = "TOTAL MATCHED  \u2014  every group together"
                elif want_role == G_REM10:
                    label10 = "REMAINING  \u2014  did not match the search"
                elif want_role == G_ALL10:
                    label10 = "ALL LOTS  \u2014  nothing typed in the box"
                else:
                    label10 = (f"GROUP {want_role} TOTAL  \u2014  lots matching  "
                               f"{terms[want_role - 1]}")
                eq(f"{LIVE_SEARCH}!A{rr} [total marker] {tag}", vals.get((S, f"A{rr}")),
                   "\u2211" if used else "")
                for scol in SEARCH_COLS:
                    L = get_column_letter(2 + SEARCH_COLS.index(scol))
                    if scol == "G":
                        want = f"{many} lot(s)" if used else ""
                    elif scol == "B":
                        want = label10
                    elif scol == "F":
                        want = ""
                    elif scol in ADDITIVE_SRC_COLS:
                        want = sum(num(truth(scol, srows10[i])) for i in who) if used else ""
                    else:
                        want = "\u2013" if used else ""
                    eq(f"{LIVE_SEARCH}!{L}{rr} [{label_of10[scol]}/TOTAL] {tag}",
                       vals.get((S, f"{L}{rr}")), want)
                out = sum(num(truth("AD", srows10[i])) for i in who)
                eq(f"{LIVE_SEARCH}!AI{rr} [status/TOTAL] {tag}", vals.get((S, f"AI{rr}")),
                   ("ALL SETTLED" if out <= 0 else "OUTSTANDING") if used else "")
        return hits

    hits10 = check_live10(vals, "[empty search]", "", None)
    print(f"live search: empty box -> all {n10} lots in one ALL LOTS block, grand total below")

    cases10 = (("187", None), ("NATIONAL", None), ("OMKAR", "Buyer"),
               ("1874, 1923", None), ("187,,1923 ,", None), ("OMKAR, STERLING", "Buyer"),
               ("copper + 1875", None), ("STERLING + 2011 / 2006", None))
    for i10, (text10, sin10) in enumerate(cases10):
        tmp = f"/tmp/verify_ls_{i10}.xlsx"
        shutil.copy(path, tmp)
        wb5 = openpyxl.load_workbook(tmp)
        wb5[LIVE_SEARCH]["B3"] = text10
        if sin10:
            wb5[LIVE_SEARCH]["F3"] = sin10
        wb5.save(tmp)
        vals5 = calculate(tmp)
        got10 = check_live10(vals5, f"[search {text10!r}/{sin10 or SEARCH_ALL}]", text10, sin10)
        print(f"live search: {text10!r} in {sin10 or SEARCH_ALL} -> {got10} match(es) grouped, "
              f"{n10 - got10} in the remaining block")

    # ================= TRANSPOSED MIRROR OF THE SOURCE ==================== #
    # four sheets share this layout; the folded ones simply reorder the columns
    from openpyxl.utils import get_column_letter as _gcl
    src_ws = wb[SRC]
    fields7 = transpose_labels(src_ws)
    srows = lot_rows(src_ws)
    total_srow = srows[-1] + 1
    for tname, thdr in ((TRANSPOSED, 3), (T_FIELD_FILTER, 3), (T_LOT_FOLDS, 5), (T_BOTH, 5)):
        ws = wb[tname]
        row_of_label = {}
        for r in range(thdr + 1, ws.max_row + 1):
            v = ws[f"A{r}"].value
            if isinstance(v, str) and v.strip():
                row_of_label[v.strip()] = r
        lot_col_of_row, total_letter = {}, None
        for c in ws[thdr]:
            if c.column < 2 or not isinstance(c.value, str):
                continue
            m = re.search(r"\$F\$(\d+)", c.value)
            if m:
                lot_col_of_row[int(m.group(1))] = c.column_letter
            elif c.value.startswith("TOTAL (src row"):
                total_letter = c.column_letter
        print(f"{tname}: {len(row_of_label)} field rows, {len(lot_col_of_row)} lot columns, "
              f"total column {total_letter} (source row {total_srow})")
        checks += 3
        if len(row_of_label) != len(fields7):
            fails.append(f"{tname}: {len(row_of_label)} field rows, expected {len(fields7)}")
        if sorted(lot_col_of_row) != sorted(srows):
            fails.append(f"{tname}: lot columns {sorted(lot_col_of_row)} != source rows {sorted(srows)}")
        if total_letter is None:
            fails.append(f"{tname}: no TOTAL column on the header row")
        for scol, label in fields7:
            rr = row_of_label.get(label)
            if rr is None:
                fails.append(f"{tname}: field row '{label}' missing")
                continue
            for srow in srows:
                letter = lot_col_of_row[srow]
                t = truth(scol, srow)
                eq(f"{tname}!{letter}{rr} [{label}/lot row {srow}]",
                   V(tname, f"{letter}{rr}"), "" if t is None else t)
            if total_letter:
                t = truth(scol, total_srow)
                eq(f"{tname}!{total_letter}{rr} [{label}/source total row]",
                   V(tname, f"{total_letter}{rr}"), "" if t is None else t)

    # ---------- structure of the four filter sheets ------------------------- #
    # frows = how many rows the dropdown must list (33 field rows on the
    # transposed sheets, the whole buyer-block stack on the collapsible one)
    spec = {TRANSPOSED: dict(hdr=3, filt="A3:A36", freeze="B4", lvl2=0, label=True, frows=33),
            T_FIELD_FILTER: dict(hdr=3, filt="A3:A36", freeze="B4", lvl2=0, label=True, frows=33),
            T_LOT_FOLDS: dict(hdr=5, filt=None, freeze="B6", lvl2=37, label=False, frows=0),
            T_BOTH: dict(hdr=5, filt="A5:A38", freeze="B6", lvl2=37, label=True, frows=33),
            # the collapsible stack carries one grey spacer row per buyer but the
            # last, and those sit inside the filter range
            C_FIELD_FILTER: dict(hdr=5, filt="A5:A563", freeze="B6", lvl2=0, label=True,
                                 frows=558, blanks=len(buyers) - 1)}
    for name, sp in spec.items():
        ws = wb[name]
        checks += 2
        if ws.auto_filter.ref != sp["filt"]:
            fails.append(f"{name}: autofilter {ws.auto_filter.ref!r}, expected {sp['filt']!r}")
        if ws.freeze_panes != sp["freeze"]:
            fails.append(f"{name}: freeze {ws.freeze_panes!r}, expected {sp['freeze']!r}")
        # a label-only filter must stay in one column, otherwise every header
        # cell in the range grows its own dropdown arrow
        if sp["filt"] and sp["label"]:
            lo, hi = sp["filt"].split(":")
            checks += 1
            if re.sub(r"\d", "", lo) != re.sub(r"\d", "", hi) or re.sub(r"\d", "", lo) != "A":
                fails.append(f"{name}: filter {sp['filt']} is not a single label column")
            # and the dropdown must list the 33 field names, one per row
            labels = [ws[f"A{r}"].value for r in range(sp["hdr"] + 1,
                                                      int(re.sub(r"\D", "", hi)) + 1)]
            blanks = sum(1 for v in labels if not (isinstance(v, str) and v.strip()))
            checks += 1
            if len(labels) != sp["frows"] or blanks != sp.get("blanks", 0):
                fails.append(f"{name}: filter lists {len(labels)} rows ({blanks} blank), "
                             f"expected {sp['frows']} rows ({sp.get('blanks', 0)} blank)")
        n2 = sum(1 for L, d in ws.column_dimensions.items() if d.outlineLevel == 2)
        checks += 1
        if n2 != sp["lvl2"]:
            fails.append(f"{name}: {n2} level-2 columns, expected {sp['lvl2']}")
    print(f"filter sheets: autofilter / freeze / outline levels checked for {len(spec)} sheets")

    # ---------- the folded sheets: order, groups and bands ------------------- #
    for name in (T_LOT_FOLDS, T_BOTH):
        ws = wb[name]
        hdr = 5
        order = []
        for c in ws[hdr]:
            if c.column >= 2 and isinstance(c.value, str) and "$F$" in c.value:
                order.append((c.column_letter, int(re.search(r"\$F\$(\d+)", c.value).group(1))))
        keys = [(src_ws[f"D{r}"].value, str(src_ws[f"G{r}"].value).strip(),
                 float(src_ws[f"F{r}"].value)) for _l, r in order]
        checks += 1
        if keys != sorted(keys):
            fails.append(f"{name}: lot columns are not sorted by auction, buyer, lot no.")
        runs, auctions = [], []
        for (letter, srow), key in zip(order, keys):
            if runs and runs[-1][0] == key[:2]:
                runs[-1][2].append(letter)
            else:
                runs.append((key[:2], None, [letter]))
            if not auctions or auctions[-1][0] != key[0]:
                auctions.append((key[0], letter, letter))
            else:
                auctions[-1] = (key[0], auctions[-1][1], letter)
        print(f"{name}: {len(runs)} buyer runs across {len(auctions)} auctions")
        checks += 2
        if len(runs) != 21 or len(auctions) != 4:
            fails.append(f"{name}: {len(runs)} buyer runs / {len(auctions)} auctions, expected 21 / 4")
        # a band is any labelled cell on row 3 / 4; single-column runs are not merged
        bands = {}
        for rr_ in (3, 4):
            for c in ws[rr_]:
                if c.column >= 2 and isinstance(c.value, str) and c.value.strip():
                    last = c.column
                    for m in ws.merged_cells.ranges:
                        if m.min_row == rr_ and m.min_col == c.column:
                            last = m.max_col
                    bands.setdefault(rr_, []).append((c.column, last, c.value.strip()))
            bands.setdefault(rr_, []).sort()
        checks += 2
        if len(bands.get(4, [])) != len(runs):
            fails.append(f"{name}: {len(bands.get(4, []))} buyer bands, expected {len(runs)}")
        if len(bands.get(3, [])) != len(auctions):
            fails.append(f"{name}: {len(bands.get(3, []))} auction bands, expected {len(auctions)}")
        # each buyer band must cover exactly that buyer's lot columns
        from openpyxl.utils import column_index_from_string as _cix
        for (_key, _x, letters), (c1, c2, text) in zip(
                runs, sorted(bands.get(4, []), key=lambda b: b[0])):
            checks += 1
            if (c1, c2) != (_cix(letters[0]), _cix(letters[-1])) or text.split("  (")[0] != _key[1]:
                fails.append(f"{name}: buyer band {text!r} at cols {c1}-{c2} != "
                             f"{_key[1]} over {letters[0]}-{letters[-1]}")
        # a level-1 column must sit between every two buyer runs of one auction
        n1 = sum(1 for L, d in ws.column_dimensions.items() if d.outlineLevel == 1)
        checks += 1
        if n1 != len(runs) - len(auctions):
            fails.append(f"{name}: {n1} level-1 divider columns, expected {len(runs) - len(auctions)}")

    # ---------- Collapsible + Field Filter mirrors Collapsible - Lots Across - #
    a, b = wb[COLLAPSIBLE], wb[C_FIELD_FILTER]
    checks += 1
    if (a.max_row, a.max_column) != (b.max_row, b.max_column):
        fails.append(f"{C_FIELD_FILTER}: size {b.max_row}x{b.max_column} != "
                     f"{a.max_row}x{a.max_column}")
    diff = 0
    for ra, rb in zip(a.iter_rows(min_row=1, max_row=a.max_row, max_col=a.max_column),
                      b.iter_rows(min_row=1, max_row=b.max_row, max_col=b.max_column)):
        for ca, cb in zip(ra, rb):
            if ca.value != cb.value:
                diff += 1
                if diff < 4:
                    fails.append(f"{C_FIELD_FILTER}!{cb.coordinate}: {cb.value!r} != "
                                 f"{COLLAPSIBLE}!{ca.coordinate}: {ca.value!r}")
    checks += 1
    if diff:
        fails.append(f"{C_FIELD_FILTER}: {diff} cells differ from {COLLAPSIBLE}")
    lvl_diff = sum(1 for r in range(1, a.max_row + 1)
                   if a.row_dimensions[r].outlineLevel != b.row_dimensions[r].outlineLevel)
    checks += 1
    if lvl_diff:
        fails.append(f"{C_FIELD_FILTER}: {lvl_diff} rows differ in outline level")
    jump = sum(1 for r in range(1, b.max_row + 1)
               if b[f"A{r}"].hyperlink is not None
               and b[f"A{r}"].hyperlink.location.startswith(f"'{C_FIELD_FILTER}'!"))
    checks += 1
    if jump != len(buyers):
        fails.append(f"{C_FIELD_FILTER}: {jump} index links point at its own sheet, "
                     f"expected {len(buyers)}")
    print(f"{C_FIELD_FILTER}: mirrored {a.max_row}x{a.max_column} of {COLLAPSIBLE}, "
          f"{jump} jump links, {diff} cell diffs")

    print(f"\nchecks run: {checks}")
    if fails:
        print(f"FAILURES: {len(fails)}")
        for f in fails[:25]:
            print("  -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "/tmp/test.xlsx")
