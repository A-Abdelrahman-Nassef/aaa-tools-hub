"""Build/extend an Excel register of uploaded BOQ files, grouped by asset and package.

File name  FF-L-CV-AR-NEW.xlsx  ->  Asset FF-L-CV, Package AR-NEW
The asset is the first ASSET_PARTS hyphen-separated parts (default 3); the rest is the package.

Output: File_Names.xlsx with two sheets
  * "Assets & Packages" - an Excel Table: No. | File Name | Asset | Package
  * "Matrix"            - assets down, packages across, a tick where the file exists
                          (COUNTIFS formulas on the table, so it updates if rows are edited)
plus File_Names.txt (asset line, its packages below, no blank lines) for pasting.

Cumulative: previous rows are read from the existing outputs and from any uploaded
File_Names.xlsx / .txt (grouped or old flat lists). Earlier assets and packages keep
their order; new packages go under their asset, new assets at the bottom. No duplicates.

Ported from the 06-03 skill (Step 02.2: Extracting excel sheets names) for the AAA Tools Hub.
main() returns (groups, rows before, new names) so the hub can report on them.
"""
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"}
PREFIX = re.compile(r"^\d{10,}_")
ASSET_PARTS = 3
REGISTER = "Assets & Packages"
MATRIX = "Matrix"


def clean(name):
    base = os.path.basename(str(name))
    if os.path.splitext(base)[1].lower() in EXTS:
        base = os.path.splitext(base)[0]
    return PREFIX.sub("", base).strip()


def split(name):
    parts = name.split("-")
    n = ASSET_PARTS if len(parts) > ASSET_PARTS else max(len(parts) - 1, 1)
    return "-".join(parts[:n]), "-".join(parts[n:])


class Groups:
    def __init__(self):
        self.order, self.codes, self.keys = [], {}, {}

    def asset(self, a):
        k = a.upper()
        if k not in self.keys:
            self.keys[k] = a
            self.order.append(a)
            self.codes[a] = []
        return self.keys[k]

    def add(self, a, code):
        a = self.asset(a)
        if code and code.upper() not in {c.upper() for c in self.codes[a]}:
            self.codes[a].append(code)
            return True
        return False

    def rows(self):
        return [(a, c) for a in self.order for c in self.codes[a]]


def read_txt(path, g):
    current = None
    with open(path, encoding="utf-8-sig", errors="replace") as fh:
        for ln in fh.read().splitlines():
            ln = clean(ln.strip())
            if not ln:
                continue
            segs = len(ln.split("-"))
            if segs > ASSET_PARTS:
                a, c = split(ln)
                g.add(a, c)
                current = g.asset(a)
            elif segs == ASSET_PARTS:
                current = g.asset(ln)
            elif current:
                g.add(current, ln)


def read_xlsx(path, g):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return
    if REGISTER not in wb.sheetnames:
        return
    ws = wb[REGISTER]
    hdr, col = None, {}
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            if row and "Asset" in row and "Package" in row:
                hdr = row
                col = {h: i for i, h in enumerate(row) if h}
            continue
        a, p = row[col["Asset"]], row[col["Package"]]
        if a and p:
            g.add(str(a).strip(), str(p).strip())


def write_xlsx(path, g, names=None):
    """names: optional {(ASSET, PACKAGE) upper-cased: real file name} for the File Name column (hub)."""
    names = names or {}
    rows = g.rows()
    wb = Workbook()
    ws = wb.active
    ws.title = REGISTER
    f = Font(name="Arial", size=10)
    fb = Font(name="Arial", size=10, bold=True)
    fh = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    headers = ["No.", "File Name", "Asset", "Package"]
    ws.append(headers)
    for i, (a, p) in enumerate(rows, 1):
        ws.append([i, names.get((a.upper(), p.upper()), f"{a}-{p}"), a, p])
    last = len(rows) + 1
    for r in ws.iter_rows(min_row=1, max_row=last):
        for c in r:
            c.font = fh if c.row == 1 else f
            c.border = box
            c.alignment = center if c.column in (1, 4) or c.row == 1 else Alignment(vertical="center")
    for c in ws[1]:
        c.fill = fill
    for col, w in zip("ABCD", (7, 24, 16, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    if rows:
        t = Table(displayName="tblAssetPackages", ref=f"A1:D{last}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)

    # Matrix: assets down, packages across
    m = wb.create_sheet(MATRIX)
    pkgs = []
    for a in g.order:
        for p in g.codes[a]:
            if p.upper() not in {x.upper() for x in pkgs}:
                pkgs.append(p)
    pkgs.sort(key=str.upper)
    m.append(["Asset"] + pkgs + ["No. of Packages"])
    rng_a = f"'{REGISTER}'!$C$2:$C${max(last, 2)}"
    rng_p = f"'{REGISTER}'!$D$2:$D${max(last, 2)}"
    for r, a in enumerate(g.order, 2):
        m.cell(r, 1, a)
        for c, p in enumerate(pkgs, 2):
            h = get_column_letter(c)
            m.cell(r, c, f'=IF(COUNTIFS({rng_a},$A{r},{rng_p},{h}$1)>0,"✓","")')
        m.cell(r, len(pkgs) + 2, f"=COUNTIF({rng_a},$A{r})")
    tr = len(g.order) + 2
    m.cell(tr, 1, "Total")
    for c in range(2, len(pkgs) + 3):
        h = get_column_letter(c)
        m.cell(tr, c, f'=COUNTIF({h}2:{h}{tr-1},"✓")' if c < len(pkgs) + 2 else f"=SUM({h}2:{h}{tr-1})")
    for row in m.iter_rows(min_row=1, max_row=tr, max_col=len(pkgs) + 2):
        for c in row:
            c.border = box
            c.font = fh if c.row == 1 else (fb if c.row == tr or c.column == 1 else f)
            c.alignment = center if c.column > 1 or c.row == 1 else Alignment(vertical="center")
            if c.row == 1:
                c.fill = fill
            elif c.row == tr:
                c.fill = PatternFill("solid", fgColor="DDEBF7")
    m.column_dimensions["A"].width = 16
    for c in range(2, len(pkgs) + 2):
        m.column_dimensions[get_column_letter(c)].width = 10
    m.column_dimensions[get_column_letter(len(pkgs) + 2)].width = 16
    m.freeze_panes = "B2"
    wb.save(path)


def main(src, out_dir):
    xlsx = os.path.join(out_dir, "File_Names.xlsx")
    txt = os.path.join(out_dir, "File_Names.txt")
    g = Groups()
    if os.path.isfile(xlsx):
        read_xlsx(xlsx, g)
    if os.path.isfile(txt):
        read_txt(txt, g)
    for f in sorted(os.listdir(src)):
        p = os.path.join(src, f)
        if f.lower().endswith(".txt"):
            read_txt(p, g)
        elif f.lower().endswith(".xlsx"):
            read_xlsx(p, g)  # only picks up a previous register (has the REGISTER sheet)
    before = len(g.rows())

    names = sorted({clean(f) for f in os.listdir(src)
                    if os.path.splitext(f)[1].lower() in EXTS and "FILE_NAMES" not in f.upper()}, key=str.upper)
    new = [n for n in names if g.add(*split(n))]

    os.makedirs(out_dir, exist_ok=True)
    write_xlsx(xlsx, g)
    lines = []
    for a in g.order:
        lines.append(a)
        lines += g.codes[a]
    with open(txt, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write("\n".join(lines))

    print(f"Assets: {len(g.order)} | Previous files: {before} | New added: {len(new)} | Total files: {len(g.rows())}")
    if new:
        print("New: " + ", ".join(new))
    for a in g.order:
        print(f"{a}: {', '.join(g.codes[a])}")
    return g, before, new


if __name__ == "__main__":
    src, out_dir = sys.argv[1], sys.argv[2]
    main(src, out_dir)
