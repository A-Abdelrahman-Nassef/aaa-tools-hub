#!/usr/bin/env python3
"""Copy Excel files to unlocked copies: remove sheet/workbook protection only, keep everything else byte-identical.

Ported from the 06-01 skill (Step-01-3: Copying files to unlocked files) for the AAA Tools Hub.
"""
import argparse, os, re, shutil, sys, tempfile, zipfile

EXTS = ('.xlsx', '.xlsm', '.xltx', '.xltm')
SHEET_RE = re.compile(r'xl/(worksheets|chartsheets)/sheet\d+\.xml$')
TAGS = [
    re.compile(r'<(\w+:)?sheetProtection\b[^>]*/>'),
    re.compile(r'<(\w+:)?sheetProtection\b[^>]*>.*?</(\w+:)?sheetProtection>', re.S),
    re.compile(r'<(\w+:)?workbookProtection\b[^>]*/>'),
    re.compile(r'<(\w+:)?workbookProtection\b[^>]*>.*?</(\w+:)?workbookProtection>', re.S),
    re.compile(r'<(\w+:)?fileSharing\b[^>]*/>'),
]

def collect(paths, tmp):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for f in sorted(os.listdir(p)):
                files += collect([os.path.join(p, f)], tmp)
        elif p.lower().endswith('.zip') and not p.lower().endswith(EXTS):
            d = os.path.join(tmp, os.path.splitext(os.path.basename(p))[0])
            zipfile.ZipFile(p).extractall(d)
            files += collect([d], tmp)
        elif p.lower().endswith(EXTS) and not os.path.basename(p).startswith('~$'):
            files.append(p)
        elif p.lower().endswith('.xls'):
            print(f"SKIP legacy .xls (needs conversion first): {os.path.basename(p)}")
    return files

def unlock(src_path, dst_path):
    if not zipfile.is_zipfile(src_path):
        return None  # encrypted / password-to-open
    removed = 0
    with zipfile.ZipFile(src_path) as src, zipfile.ZipFile(dst_path, 'w') as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if SHEET_RE.match(info.filename) or info.filename == 'xl/workbook.xml':
                s = data.decode('utf-8')
                for rx in TAGS:
                    s, n = rx.subn('', s)
                    removed += n
                data = s.encode('utf-8')
            dst.writestr(info, data, compress_type=info.compress_type)
    return removed

def verify(a_path, b_path):
    import openpyxl
    a = openpyxl.load_workbook(a_path); b = openpyxl.load_workbook(b_path)
    ok = a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        sa, sb = a[name], b[name]
        if getattr(sb, 'protection', None) is not None and sb.protection.sheet:
            ok = False
        if hasattr(sa, 'iter_rows'):
            for ra, rb in zip(sa.iter_rows(), sb.iter_rows()):
                for ca, cb in zip(ra, rb):
                    if ca.value != cb.value or ca.style_id != cb.style_id:
                        return False
            ok &= sa.max_row == sb.max_row and sa.max_column == sb.max_column
            ok &= len(sa.merged_cells.ranges) == len(sb.merged_cells.ranges)
            ok &= len(sa._images) == len(sb._images)
    ok &= not b.security or not b.security.lockStructure
    return ok

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('inputs', nargs='+')
    ap.add_argument('--out', required=True)
    ap.add_argument('--no-zip', action='store_true')
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    tmp = tempfile.mkdtemp()
    files = collect(args.inputs, tmp)
    if not files:
        sys.exit("No Excel files found.")
    bad = 0
    for f in files:
        name = os.path.basename(f)
        dst = os.path.join(args.out, name)
        n = unlock(f, dst)
        if n is None:
            print(f"ENCRYPTED (password to open, cannot unlock): {name}"); bad += 1; continue
        try:
            ok = verify(f, dst)
        except Exception as e:
            ok = False; print(f"  verify error {name}: {e}")
        status = 'unlocked' if n else 'no protection found (copied as is)'
        print(f"{name}: {status} | tags removed={n} | identical={ok}")
        bad += not ok
    if not args.no_zip:
        z = shutil.make_archive(args.out, 'zip', os.path.dirname(args.out), os.path.basename(args.out))
        print("ZIP:", z)
    sys.exit(1 if bad else 0)

if __name__ == '__main__':
    main()
