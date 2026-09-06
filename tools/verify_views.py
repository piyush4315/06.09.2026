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
import sys
from collections import OrderedDict

import openpyxl

sys.path.insert(0, "tools")
from build_buyer_lot_views import (COLLAPSIBLE, LEDGER, SECTIONS, SRC,  # noqa: E402
                                   read_lots)

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
    for name in (COLLAPSIBLE, LEDGER):
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

    print(f"\nchecks run: {checks}")
    if fails:
        print(f"FAILURES: {len(fails)}")
        for f in fails[:25]:
            print("  -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "/tmp/test.xlsx")
