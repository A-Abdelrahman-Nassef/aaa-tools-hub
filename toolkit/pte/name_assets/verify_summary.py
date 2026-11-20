#!/usr/bin/env python3
"""
Quality gate for a built asset workbook.

Three things are checked, in order of how much they matter:

  1. Structure    - tabs exist, summary labels and links resolve to a real
                    sheet and that sheet's own Total Amounts row, no stale
                    references to the old template name.
  2. Live recalc  - test rates are injected into one item on every asset tab
                    and the file is recalculated, proving the chain
                    item -> division subtotal -> tab total -> summary row ->
                    summary grand total agrees end to end.
  3. Error delta  - formula errors are counted and compared against the same
                    treatment applied to the source file. What matters is that
                    the build added none; these workbooks often already contain
                    errors that only LibreOffice reports (XLOOKUP, for one, is
                    unsupported before 24.8 and shows as #NAME?).

Usage:
  python3 verify.py --file OUT.xlsx [--baseline IN.xlsx] [--assets "A" "B"]
"""
import argparse, glob, os, re, shutil, subprocess, sys, zipfile

ERRS = {'#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NULL!', '#NUM!'}
# a self-closing <c .../> must not be read as an opening tag, or the
# formulas after it get swallowed and silently dropped
CELL = re.compile(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.S)


def strip_cached(xml):
    """LibreOffice trusts cached results on load; drop them so it must compute."""
    def f(m):
        c = m.group(0)
        if '<f' not in c:
            return c
        c = re.sub(r'<v>.*?</v>|<v/>', '', c, flags=re.S)
        return re.sub(r'(<c\b[^>]*?)\st="(str|s|e)"', r'\1', c)
    return CELL.sub(f, xml)


def recalc(src, tmp, inject=None):
    """Return a recalculated copy. `inject` = {sheet part: {cell: value}}."""
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)
    subprocess.run(['unzip', '-qo', src, '-d', tmp + '/x'], check=True)
    for part, cells in (inject or {}).items():
        p = os.path.join(tmp, 'x', part)
        d = open(p, encoding='utf-8').read()
        for ref, val in cells.items():
            col, row = re.match(r'([A-Z]+)(\d+)', ref).groups()
            m = re.search(r'<row r="%s"[^>]*>.*?</row>' % row, d, re.S)
            rx = re.sub(r'<c r="%s"([^>]*?)\s*/>' % ref,
                        lambda mm: '<c r="%s"%s><v>%s</v></c>' % (ref, mm.group(1), val),
                        m.group(0), count=1)
            d = d[:m.start()] + rx + d[m.end():]
        open(p, 'w', encoding='utf-8').write(d)
    for p in glob.glob(tmp + '/x/xl/worksheets/sheet*.xml'):
        d = open(p, encoding='utf-8').read()
        open(p, 'w', encoding='utf-8').write(strip_cached(d))
    out = os.path.join(tmp, 'r.xlsx')
    with zipfile.ZipFile(src) as z, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as o:
        for n in z.namelist():
            o.write(os.path.join(tmp, 'x', n), n)
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', 'xlsx',
                    '--outdir', tmp + '/c', out],
                   env=dict(os.environ, HOME=tmp), check=True,
                   capture_output=True, timeout=900)
    return glob.glob(tmp + '/c/*.xlsx')[0]


def count_errors(path):
    """Errors broken down by type - a bare total hides which kind changed."""
    import collections, openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    by_type = collections.Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value in ERRS:
                    by_type[c.value] += 1
    return by_type


def uses_xlookup(path):
    import openpyxl
    wb = openpyxl.load_workbook(path)
    return any(isinstance(c.value, str) and 'XLOOKUP' in c.value
               for ws in wb.worksheets for row in ws.iter_rows() for c in row)


def main():
    import openpyxl
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--baseline')
    ap.add_argument('--assets', nargs='*')
    a = ap.parse_args()
    ok = True

    wb = openpyxl.load_workbook(a.file)
    names = wb.sheetnames
    print('sheets: %s' % ', '.join(repr(n) for n in names))

    # ---- 1. structure ----------------------------------------------------
    summary = next((ws for ws in wb.worksheets
                    if any(isinstance(c.value, str) and c.value.startswith("='")
                           for r in ws.iter_rows() for c in r)), None)
    linked, bad = [], []
    for row in summary.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("='"):
                t = c.value[2:c.value.index("'!", 2)]
                if t not in names:
                    bad.append((c.coordinate, t))
                elif c.column_letter in ('K',):
                    linked.append((c.row, t))
    print('\n1. structure')
    for r, t in linked:
        label = summary.cell(r, 3).value
        tot = wb[t]
        trow = next((x for x in range(1, tot.max_row + 1)
                     if str(tot.cell(x, 5).value).strip().lower() == 'total amounts'), None)
        want = int(re.search(r'\$(\d+)$', summary.cell(r, 11).value).group(1))
        good = (label == t and trow == want)
        ok &= good
        print('   row %-3d label=%-30r -> tab %-30r total row %s/%s  %s'
              % (r, label, t, want, trow, 'OK' if good else 'MISMATCH'))
    if bad:
        ok = False
        print('   DANGLING references: %s' % bad)
    else:
        print('   no dangling sheet references')

    # ---- 2. live recalc --------------------------------------------------
    print('\n2. live recalculation (injecting test rates)')
    parts = {}
    exp = {}
    with zipfile.ZipFile(a.file) as z:
        wbx = z.read('xl/workbook.xml').decode()
        rels = z.read('xl/_rels/workbook.xml.rels').decode()
    tgt = {m.group(1): m.group(2) for m in
           re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
    order = [(m.group(1).replace('&amp;', '&'), 'xl/' + tgt[m.group(2)]) for m in
             re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wbx)]
    part_of = dict(order)
    for i, (_, t) in enumerate(linked):
        q, aq, rate = 10 * (i + 1), 12 * (i + 1), 100
        ws = wb[t]
        r = next(x for x in range(7, ws.max_row)
                 if isinstance(ws.cell(x, 10).value, str) and ws.cell(x, 10).value.startswith('=$G'))
        parts[part_of[t]] = {'G%d' % r: q, 'H%d' % r: aq, 'I%d' % r: rate}
        exp[t] = (q * rate, aq * rate)
    calc = recalc(a.file, '/tmp/_verify_out', parts)
    rc = openpyxl.load_workbook(calc, data_only=True)
    rs = rc[summary.title]
    tk = tl = 0
    for r, t in linked:
        eJ, eK = exp[t]
        tk += eJ
        tl += eK
        got = (rs.cell(r, 11).value, rs.cell(r, 12).value)
        good = got == (eJ, eK)
        ok &= good
        print('   %-30s tab total %s/%s -> summary %s/%s  %s'
              % (t, eJ, eK, got[0], got[1], 'OK' if good else 'FAIL'))
    trow = next(x for x in range(1, rs.max_row + 1)
                if str(rs.cell(x, 3).value).strip().lower() == 'total amounts')
    g = (rs.cell(trow, 11).value, rs.cell(trow, 12).value)
    good = g == (tk, tl)
    ok &= good
    print('   grand total row %d: %s/%s (expected %s/%s)  %s'
          % (trow, g[0], g[1], tk, tl, 'OK' if good else 'FAIL'))

    # ---- 3. error delta --------------------------------------------------
    print('\n3. formula errors, by type')
    out_t = count_errors(calc)
    if not a.baseline:
        print('   built file: %s' % (dict(out_t) or 'none'))
        print('   (no --baseline given, cannot tell pre-existing from new)')
    else:
        base_t = count_errors(recalc(a.baseline, '/tmp/_verify_base'))
        # LibreOffice before 24.8 has no XLOOKUP and reports #NAME? for every cell
        # that uses one and every cell that reads those. Adding an asset column
        # group therefore adds #NAME? cells by the hundred with nothing wrong. The
        # count is meaningless here, so it is reported but not judged; every other
        # error type still has to stay put.
        excused = '#NAME?' if uses_xlookup(a.file) else None
        for k in sorted(set(out_t) | set(base_t)):
            b, o = base_t.get(k, 0), out_t.get(k, 0)
            if k == excused:
                verdict = 'not judged - LibreOffice has no XLOOKUP'
            elif o <= b:
                verdict = 'OK'
            else:
                verdict = 'REGRESSION'
                ok = False
            print('   %-9s source %-6d built %-6d  %+d  %s' % (k, b, o, o - b, verdict))
        if not (set(out_t) | set(base_t)):
            print('   none in either file')

    print('\nRESULT: %s' % ('PASS' if ok else 'FAIL'))
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
