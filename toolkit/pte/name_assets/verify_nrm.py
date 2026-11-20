#!/usr/bin/env python3
"""
Prove each NRM asset group pulls the right number from the right tab.

An NRM code and a rate are planted on one item row of every asset tab, so that
tab's return column holds a known amount against a known code. Each group should
then show its own tab's number and nothing else - if a tab reference, key column
or return column were wrong, the result would be 0 or another group's figure.

Two LibreOffice quirks are worked around in the throwaway copy, neither of which
affects Excel:
  - no XLOOKUP before 24.8, so each lookup is swapped for the SUMIF that is
    equivalent when NRM codes are unique (same tab, same key, same return column)
  - outline-hidden rows are treated as filtered out and dropped from SUBTOTAL,
    which Excel does not do for function 9, so the rows are shown for the test

Usage:
  python3 verify_nrm.py --file OUT.xlsx --names "Asset A" "Asset B" ...
"""
import argparse, glob, os, re, shutil, subprocess, sys, zipfile
import openpyxl

ERR = {'#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NULL!', '#NUM!'}
# a self-closing <c .../> must not be read as an opening tag, or the
# formulas after it get swallowed and silently dropped
CELL = re.compile(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.S)


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--names', nargs='+', required=True)
    ap.add_argument('--code', default=None, help='NRM code to plant; default: the first one')
    a = ap.parse_args()
    T = '/tmp/_nrmverify'
    shutil.rmtree(T, ignore_errors=True)
    os.makedirs(T)
    subprocess.run(['unzip', '-qo', a.file, '-d', T + '/x'], check=True)

    wb = openpyxl.load_workbook(a.file)
    nrm = next(ws for ws in wb.worksheets
               if any(isinstance(c.value, str) and 'XLOOKUP' in c.value
                      for row in ws.iter_rows() for c in row))
    # the amount column of each group, and the row each lookup lives on
    amt_cols, probe_row = [], None
    for row in nrm.iter_rows():
        for c in row:
            if isinstance(c.value, str) and 'XLOOKUP' in c.value:
                if c.column_letter not in amt_cols:
                    amt_cols.append(c.column_letter)
                if probe_row is None:
                    probe_row = c.row
    amt_cols.sort(key=colnum)
    width = colnum(amt_cols[1]) - colnum(amt_cols[0]) if len(amt_cols) > 1 else 3
    code = a.code or nrm.cell(probe_row, 2).value
    print('NRM sheet %r, %d groups at %s, probing row %d (code %r)'
          % (nrm.title, len(amt_cols), ', '.join(amt_cols), probe_row, code))
    if len(amt_cols) != len(a.names):
        print('  NOTE: %d groups but %d names given' % (len(amt_cols), len(a.names)))

    # sheet part for each tab
    with zipfile.ZipFile(a.file) as z:
        wbx, rels = z.read('xl/workbook.xml').decode(), z.read('xl/_rels/workbook.xml.rels').decode()
    tgt = {m.group(1): m.group(2) for m in
           re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
    part_of = {m.group(1).replace('&amp;', '&'): 'xl/' + tgt[m.group(2)] for m in
               re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wbx)}

    # the asset tabs' key column, return column, and a blank item row to use
    f = nrm.cell(probe_row, colnum(amt_cols[0])).value
    key_col, ret_col = re.search(r"!\$([A-Z]+):\$[A-Z]+,'[^']+'!\$([A-Z]+):", f).groups()
    ss = open(T + '/x/xl/sharedStrings.xml', encoding='utf-8').read()
    texts = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', s, re.S))
             for s in re.findall(r'<si>.*?</si>', ss, re.S)]
    if code in texts:
        code_si = texts.index(code)
    else:
        code_si = len(texts)
        ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % code)
        open(T + '/x/xl/sharedStrings.xml', 'w', encoding='utf-8').write(ss)
        d = open(T + '/x/xl/sharedStrings.xml', encoding='utf-8').read()
        open(T + '/x/xl/sharedStrings.xml', 'w', encoding='utf-8').write(
            re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % (code_si + 1), d, count=1))

    exp = {}
    for i, nm in enumerate(a.names):
        ws = wb[nm]
        # the return column's formula is QTY * rate; find its two input columns
        # an item row: the return column multiplies a QTY by a rate.
        # Division rows hold SUBTOTALs, so match the multiplication specifically.
        item = re.compile(r'^=\$?([A-Z]+)\d+\*\$?([A-Z]+)\d+$')
        r = g = None
        for x in range(1, ws.max_row):
            v = ws.cell(x, colnum(ret_col)).value
            if isinstance(v, str):
                m2 = item.match(v.replace(' ', ''))
                if m2:
                    r, g = x, m2
                    break
        if not g:
            raise SystemExit('No item row found on %r' % nm)
        qty_col, rate_col = g.groups()
        qty, rate = 10 * (i + 1), 100
        p = os.path.join(T, 'x', part_of[nm])
        d = open(p, encoding='utf-8').read()
        m = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, d, re.S)
        rx = re.sub(r'<c r="%s%d"([^>]*?)\s*/>' % (key_col, r),
                    lambda mm: '<c r="%s%d"%s t="s"><v>%d</v></c>' % (key_col, r, mm.group(1), code_si),
                    m.group(0), count=1)
        for col, v in ((qty_col, qty), (rate_col, rate)):
            rx = re.sub(r'<c r="%s%d"([^>]*?)\s*/>' % (col, r),
                        lambda mm: '<c r="%s%d"%s><v>%s</v></c>' % (col, r, mm.group(1), v),
                        rx, count=1)
        open(p, 'w', encoding='utf-8').write(d[:m.start()] + rx + d[m.end():])
        exp[nm] = qty * rate

    # NRM sheet: SUMIF instead of XLOOKUP, and show the outline-hidden rows
    p2 = os.path.join(T, 'x', part_of[nrm.title])
    d2 = open(p2, encoding='utf-8').read()
    n = len(re.findall(r'_xlfn\.XLOOKUP', d2))
    d2 = re.sub(r"_xlfn\.XLOOKUP\((\$B\d+),('[^']+'!\$[A-Z]+:\$[A-Z]+),"
                r"('[^']+'!\$[A-Z]+:\$[A-Z]+),0,0,1\)", r'SUMIF(\2,\1,\3)', d2)
    assert '_xlfn.XLOOKUP' not in d2
    open(p2, 'w', encoding='utf-8').write(d2.replace(' hidden="1"', ''))
    print('swapped %d XLOOKUPs for the equivalent SUMIF' % n)

    def strip(mm):
        c = mm.group(0)
        if '<f' not in c:
            return c
        c = re.sub(r'<v>.*?</v>|<v/>', '', c, flags=re.S)
        return re.sub(r'(<c\b[^>]*?)\st="(str|s|e)"', r'\1', c)
    for fp in glob.glob(T + '/x/xl/worksheets/sheet*.xml'):
        dd = open(fp, encoding='utf-8').read()
        open(fp, 'w', encoding='utf-8').write(CELL.sub(strip, dd))

    test = T + '/t.xlsx'
    with zipfile.ZipFile(a.file) as z, zipfile.ZipFile(test, 'w', zipfile.ZIP_DEFLATED) as o:
        for n2 in z.namelist():
            o.write(os.path.join(T, 'x', n2), n2)
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', 'xlsx',
                    '--outdir', T + '/c', test],
                   env=dict(os.environ, HOME=T), check=True, capture_output=True, timeout=900)
    rc = openpyxl.load_workbook(glob.glob(T + '/c/*.xlsx')[0], data_only=True)
    ws = rc[nrm.title]
    hdr = next(r for r in range(1, 10)
               if isinstance(ws.cell(r, colnum(amt_cols[0])).value, str))
    total = next(r for r in range(ws.max_row, 1, -1)
                 if isinstance(ws.cell(r, colnum(amt_cols[0])).value, (int, float))
                 and str(ws.cell(r, 3).value or '').lower().startswith('total'))

    ok = True
    no_bua = []
    print('\n--- each group pulls from its own tab ---')
    for i, (nm, ac) in enumerate(zip(a.names, amt_cols)):
        c0 = colnum(ac)
        got = ws.cell(probe_row, c0).value
        good = got == exp[nm]
        ok &= good
        cost = ws.cell(probe_row, c0 + 1).value
        pct = ws.cell(probe_row, c0 + 2).value
        bua = ws.cell(hdr, c0 + 2).value
        gp = abs((pct or 0) - 1.0) < 1e-9
        if isinstance(bua, (int, float)) and bua:
            gc = isinstance(cost, (int, float)) and abs(cost - exp[nm] / bua) < 1e-9
            note = '%.6f' % cost if isinstance(cost, (int, float)) else str(cost)
        else:
            no_bua.append(nm)          # the BUA cell on the header row is empty
            gc = True
            note = 'BUA not set'
        ok &= good and gc and gp
        print('   %-28s %s%-3d = %-8s  cost/m2 %-14s  %% %-7s  %s'
              % (nm, ac, probe_row, got, note,
                 round(pct, 4) if isinstance(pct, (int, float)) else pct,
                 'OK' if (good and gc and gp) else 'FAIL'))

    print('\n--- grand total row %d ---' % total)
    for nm, ac in zip(a.names, amt_cols):
        got = ws.cell(total, colnum(ac)).value
        good = got == exp[nm]
        ok &= good
        print('   %s%-3d = %-8s (expected %-8s)  %s' % (ac, total, got, exp[nm],
                                                        'OK' if good else 'FAIL'))

    bad, div0 = {}, 0
    for s in rc.worksheets:
        for row in s.iter_rows():
            for c in row:
                if c.value in ERR:
                    if c.value == '#DIV/0!' and no_bua:
                        div0 += 1
                    else:
                        bad[s.title] = bad.get(s.title, 0) + 1
    if div0:
        print('\n%d #DIV/0! cells, explained by the empty BUA cell for: %s'
              % (div0, ', '.join(no_bua)))
        print('   enter each asset\'s BUA on the header row and they resolve')
    print('other error cells:', bad or 'none')
    ok &= not bad
    print('\nRESULT:', 'PASS' if ok else 'FAIL')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
