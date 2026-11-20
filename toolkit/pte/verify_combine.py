"""Prove each ALL DIVISIONS sheet still says exactly what its source said.

    python3 verify_combine.py <source_dir_or_files> <output_dir>

Every source cell is compared with its counterpart in the combined sheet: plain
values directly, formulas translated back to their original row and compared to
the source string. Exits non-zero if anything differs, so a failed check can't
slip past into a delivery. Also reports the cross-file anomalies worth telling
a QS about.

Ported from the 06-02 skill for the AAA Tools Hub: verify() returns the per-file rows,
totals and notes so the hub can show them; main() prints them as before.
"""
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

MAXCOL = 16


def divisions(wb):
    return [n for n in wb.sheetnames
            if re.match(r"^DIVISION\b", n.strip().upper()) and "SUMMARY" not in n.upper()]


def extent(ws, maxcol=MAXCOL):
    lr = lc = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=maxcol):
        for c in row:
            if c.value is not None and str(c.value).strip() != "":
                lr, lc = max(lr, c.row), max(lc, c.column)
    return lr, lc


def header_row(ws, limit=15):
    for r in range(1, min(limit, ws.max_row) + 1):
        vals = [str(ws.cell(row=r, column=c).value or "").strip().upper() for c in range(1, 7)]
        if "ITEM" in vals and "DESCRIPTION" in vals:
            return r
    return None


def sources(args):
    files = []
    for a in args:
        if os.path.isdir(a):
            files += [os.path.join(a, f) for f in sorted(os.listdir(a))
                      if f.lower().endswith((".xlsx", ".xlsm"))
                      and not f.startswith("~$") and "ALL DIVISIONS" not in f]
        else:
            files.append(a)
    return files


def verify(src_args, outdir):
    """Returns (rows, totals, notes); rows = one dict per source file."""
    total = defaultdict(int)
    fingerprints, present, notes, rows = {}, {}, [], []

    for src_path in sources(src_args):
        stem = os.path.splitext(os.path.basename(src_path))[0]
        out_path = os.path.join(outdir, f"{stem} - ALL DIVISIONS.xlsx")
        if not os.path.exists(out_path):
            rows.append(dict(file=stem, missing=True, bad=1))
            total["bad"] += 1
            continue

        s_wb = openpyxl.load_workbook(src_path)
        t = openpyxl.load_workbook(out_path)["ALL DIVISIONS"]
        banners = {}
        for r in range(1, t.max_row + 1):
            v = t.cell(row=r, column=1).value
            if isinstance(v, str) and re.match(r"^DIVISION \d+$", v.strip()):
                banners[v.strip()] = r

        vals = forms = bad = items = 0
        fp = []
        present[stem] = set()
        for name in divisions(s_wb):
            ws = s_wb[name]
            present[stem].add(name.strip())
            end, ncols = extent(ws)
            h = header_row(ws)
            if name.strip() not in banners:
                notes.append(f"{stem}: {name} has no block in the output")
                bad += 1
                continue
            if h and any("SAR" in str(ws.cell(row=h, column=c).value or "").upper()
                         for c in range(1, 7)):
                notes.append(f"{stem}: {name} header names SAR, not EGP")
            start = banners[name.strip()] + 1
            for sr in range((h or 0) + 1, end + 1):
                tr = start + sr - ((h or 0) + 1)
                for sc in range(1, ncols + 1):
                    sv = ws.cell(row=sr, column=sc).value
                    tv = t.cell(row=tr, column=sc).value
                    if sv is None and tv is None:
                        continue
                    fp.append(f"{name}|{sr}|{sc}|{sv}")
                    if isinstance(sv, str) and sv.startswith("="):
                        forms += 1
                        col = get_column_letter(sc)
                        back = Translator(tv, origin=f"{col}{tr}").translate_formula(f"{col}{sr}")
                        if back != sv:
                            bad += 1
                            notes.append(f"{stem}: {name} r{sr} formula {sv} -> {tv}")
                    else:
                        vals += 1
                        if sv != tv:
                            bad += 1
                            notes.append(f"{stem}: {name} r{sr}c{sc} {sv!r} -> {tv!r}")
                    if sc > 6 and sv not in (None, ""):
                        notes.append(f"{stem}: stray value {sv!r} at "
                                     f"{name}!{get_column_letter(sc)}{sr} (outside the BOQ columns)")
                a = ws.cell(row=sr, column=1).value
                if a and str(a).strip() not in ("ITEM", "TOTAL OF PAGE") \
                        and ws.cell(row=sr, column=3).value:
                    items += 1

        fingerprints[stem] = hash("\n".join(fp))
        rows.append(dict(file=stem, divs=len(divisions(s_wb)), values=vals, formulas=forms, items=items, bad=bad))
        for k, v in (("values", vals), ("formulas", forms), ("items", items), ("bad", bad)):
            total[k] += v

    # cross-file checks: identical BOQs, and divisions missing from some files
    seen = defaultdict(list)
    for stem, fp in fingerprints.items():
        seen[fp].append(stem)
    for group in seen.values():
        if len(group) > 1:
            notes.append("identical division content: " + ", ".join(sorted(group)))
    everywhere = set().union(*present.values()) if present else set()
    for stem, have in present.items():
        for missing in sorted(everywhere - have):
            notes.append(f"{stem}: no {missing} tab (other files have it)")
    return rows, total, sorted(set(notes))


def main():
    rows, total, notes = verify(sys.argv[1:-1], sys.argv[-1])
    print(f"{'file':46s} {'divs':>5s} {'values':>8s} {'formulas':>9s} {'items':>6s} {'mismatch':>9s}")
    for r in rows:
        if r.get("missing"):
            print(f"{r['file'][:46]:46s}   MISSING OUTPUT")
        else:
            print(f"{r['file'][:46]:46s} {r['divs']:5d} {r['values']:8d} {r['formulas']:9d} {r['items']:6d} {r['bad']:9d}")
    print(f"\ntotal: {total['values']} values, {total['formulas']} formulas, "
          f"{total['items']} priced items, {total['bad']} mismatches")
    if notes:
        print("\nworth reporting to the user:")
        for n in notes:
            print(f"  - {n}")

    sys.exit(1 if total["bad"] else 0)


if __name__ == "__main__":
    main()
