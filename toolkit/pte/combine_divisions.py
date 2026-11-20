"""Stack every DIVISION detail sheet of a BOQ workbook into one ALL DIVISIONS sheet.

    python3 combine_divisions.py <file_or_dir> [more...] [--outdir DIR] [--zip NAME]

Writes "<original name> - ALL DIVISIONS.xlsx" per input, recalculates it in Excel
so the formulas carry cached values, and optionally zips the results.
Source workbooks are never modified.

Ported from the 06-02 skill (Step 01: Collecting all division in one sheet) for the
AAA Tools Hub: Excel (excel.py) replaces LibreOffice for recalculation and .xls conversion.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from copy import copy

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

import excel
MAXCOL = 16          # ceiling for the scan; the real extent is detected per file
SHEET_NAME = "ALL DIVISIONS"

BANNER_FONT = Font(name="Arial", size=12, bold=True, color="FFFFFF")
BANNER_FILL = PatternFill("solid", fgColor="1F3864")
_THIN = Side(style="thin", color="1F3864")
BANNER_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def division_sheets(wb):
    """Division detail tabs, in workbook order. Summaries are excluded."""
    return [n for n in wb.sheetnames
            if re.match(r"^DIVISION\b", n.strip().upper()) and "SUMMARY" not in n.upper()]


def used_extent(ws, maxcol=MAXCOL):
    """Last row and column carrying an actual value, ignoring stray formatting."""
    last_r = last_c = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=maxcol):
        for c in row:
            if c.value is not None and str(c.value).strip() != "":
                last_r, last_c = max(last_r, c.row), max(last_c, c.column)
    return last_r, last_c


def find_header(ws, limit=15):
    """Row index of the ITEM / DESCRIPTION header, or None if the tab has none."""
    for r in range(1, min(limit, ws.max_row) + 1):
        vals = [str(ws.cell(row=r, column=c).value or "").strip().upper()
                for c in range(1, 7)]
        if "ITEM" in vals and "DESCRIPTION" in vals:
            return r
    return None


def copy_row(src, sr, tgt, tr, ncols):
    """Copy one row with its styles, re-basing any formula to its new row."""
    for sc in range(1, ncols + 1):
        s = src.cell(row=sr, column=sc)
        if s.value is None and not s.has_style:
            continue
        d = tgt.cell(row=tr, column=sc)
        v = s.value
        if isinstance(v, str) and v.startswith("="):
            col = get_column_letter(sc)
            v = Translator(v, origin=f"{col}{sr}").translate_formula(f"{col}{tr}")
        d.value = v
        d._style = copy(s._style)
    if src.row_dimensions[sr].height is not None:
        tgt.row_dimensions[tr].height = src.row_dimensions[sr].height


def to_xlsx(path, workdir):
    """Legacy .xls needs converting before openpyxl can read it."""
    if not path.lower().endswith(".xls"):
        return path
    return excel.xls_to_xlsx(path, workdir)


def combine(path, outdir, workdir=None):
    """Returns (output path, status text, per-division report, recalc result, source read)."""
    workdir = workdir or os.path.join(tempfile.gettempdir(), "_boq_convert")
    src = to_xlsx(path, workdir)
    wb = openpyxl.load_workbook(src)
    divs = division_sheets(wb)
    if not divs:
        return None, "no DIVISION sheets found", [], None, src

    ncols = max(used_extent(wb[d])[1] for d in divs)
    template = wb[divs[0]]
    hdr_row = find_header(template)
    hdr_cols = 6
    if hdr_row:
        hdr_cols = max((c for c in range(1, ncols + 1)
                        if template.cell(row=hdr_row, column=c).value not in (None, "")),
                       default=6)

    tgt = wb.create_sheet(SHEET_NAME)
    tgt.sheet_format = copy(template.sheet_format)
    for c in range(1, ncols + 1):
        letter = get_column_letter(c)
        if letter in template.column_dimensions:
            tgt.column_dimensions[letter].width = template.column_dimensions[letter].width
    tgt.sheet_view.showGridLines = template.sheet_view.showGridLines
    tgt.page_setup.orientation = template.page_setup.orientation
    tgt.page_setup.paperSize = template.page_setup.paperSize
    tgt.page_margins = copy(template.page_margins)

    r = 1
    if hdr_row:                 # one header at the top, repeated on every printed page
        copy_row(template, hdr_row, tgt, 1, ncols)
        r = 2

    report = []
    for i, name in enumerate(divs):
        ws = wb[name]
        end, _ = used_extent(ws, ncols)
        h = find_header(ws)
        start_src = (h + 1) if (h and hdr_row) else 1
        if end < start_src:
            report.append((name, 0, "empty"))
            continue

        if i:
            tgt.row_breaks.append(Break(id=r - 1))      # new printed page per division

        tgt.cell(row=r, column=1, value=name.strip())
        tgt.merge_cells(start_row=r, start_column=1, end_row=r, end_column=hdr_cols)
        b = tgt.cell(row=r, column=1)
        b.font, b.fill, b.border = BANNER_FONT, BANNER_FILL, BANNER_BORDER
        b.alignment = Alignment(horizontal="center", vertical="center")
        tgt.row_dimensions[r].height = 20
        r += 1

        block_start, offset = r, r - start_src
        for sr in range(start_src, end + 1):
            copy_row(ws, sr, tgt, r, ncols)
            r += 1
        for rng in ws.merged_cells.ranges:               # carry merges, shifted
            if start_src <= rng.min_row and rng.max_row <= end:
                tgt.merge_cells(start_row=rng.min_row + offset, start_column=rng.min_col,
                                end_row=rng.max_row + offset, end_column=rng.max_col)
        report.append((name, r - block_start, f"{block_start}-{r - 1}"))

    last = r - 1
    tgt.print_area = f"'{SHEET_NAME}'!$A$1:${get_column_letter(hdr_cols)}${last}"
    if hdr_row:
        tgt.print_title_rows = "$1:$1"
        tgt.freeze_panes = "A2"

    for s in list(wb.sheetnames):                        # keep only the combined sheet
        if s != SHEET_NAME:
            del wb[s]
    for n in list(wb.defined_names):                     # drop names pointing at deleted tabs
        del wb.defined_names[n]

    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(outdir, f"{stem} - ALL DIVISIONS.xlsx")
    wb.save(out)

    rc = excel.recalc(out)
    status = "recalculated" if rc["errors"] == 0 else f"RECALC PROBLEM ({rc['errors']} formula errors)"
    return out, f"{len(divs)} divisions, {last} rows, {status}", report, rc, src


def collect(args):
    files = []
    for a in args:
        if os.path.isdir(a):
            files += [os.path.join(a, f) for f in sorted(os.listdir(a))
                      if f.lower().endswith((".xlsx", ".xlsm", ".xls"))
                      and not f.startswith("~$") and "ALL DIVISIONS" not in f]
        else:
            files.append(a)
    return files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--zip", dest="zipname", default=None,
                    help="archive the outputs under this file name")
    a = ap.parse_args()

    produced, failed = [], []
    for f in collect(a.inputs):
        print(f"\n{os.path.basename(f)}")
        try:
            out, status, report, _, _ = combine(f, a.outdir)
        except Exception as e:      # one bad file shouldn't sink the whole batch
            failed.append((os.path.basename(f), f"{type(e).__name__}: {e}"))
            print(f"  -> SKIPPED, {type(e).__name__}: {e}")
            continue
        print(f"  -> {status}")
        for name, n, where in report:
            print(f"     {name:14s} {n:5d} rows  @ {where}")
        if out:
            produced.append(out)
            print(f"  -> {os.path.basename(out)}")
        else:
            failed.append((os.path.basename(f), status))

    if a.zipname and produced:
        zpath = os.path.join(a.outdir, a.zipname)
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for p in produced:
                z.write(p, os.path.join("ALL DIVISIONS", os.path.basename(p)))
        print(f"\narchive: {os.path.basename(zpath)} ({len(produced)} files)")

    if failed:
        print(f"\n{len(failed)} file(s) NOT combined — say so rather than "
              f"letting them go missing from the archive:")
        for name, why in failed:
            print(f"  - {name}: {why}")

    shutil.rmtree(os.path.join(tempfile.gettempdir(), "_boq_convert"), ignore_errors=True)


if __name__ == "__main__":
    main()
