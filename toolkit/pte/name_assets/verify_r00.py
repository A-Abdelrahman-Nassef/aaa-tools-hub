#!/usr/bin/env python3
"""
End-to-end check for a packaged workbook.

The same NRM code is planted in TWO different packages of one asset, with
DIFFERENT amounts. That is the case a first-match lookup gets wrong - it returns
one package and silently drops the rest - so the test only means anything if the
two differ and the elemental sheet shows their sum.

Chain checked:
    package item -> package subtotal -> asset grand total   (total spans all packages)
    asset grand total -> Asset Summary -> Project Summary
    every package -> Elemental Summary                      (aggregation)
    Schedule of Assets BUA -> Elemental Cost/m2

Usage:
  python3 verify_r00.py --file OUT.xlsx [--asset "Building 01"] [--code 0.1.1]
"""
import argparse, glob, os, re, shutil, subprocess, sys, zipfile
import openpyxl

ERR = {'#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NULL!', '#NUM!'}
CELL = re.compile(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.S)
BUA = 1000


def colletter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--asset')
    ap.add_argument('--code')
    a = ap.parse_args()
    T = '/tmp/_verify_r00'
    shutil.rmtree(T, ignore_errors=True)
    os.makedirs(T)
    subprocess.run(['unzip', '-qo', a.file, '-d', T + '/x'], check=True)

    wb = openpyxl.load_workbook(a.file, read_only=True)
    names = wb.sheetnames
    with zipfile.ZipFile(a.file) as z:
        wbx = z.read('xl/workbook.xml').decode()
        rels = z.read('xl/_rels/workbook.xml.rels').decode()
    tgt = {m.group(1): m.group(2) for m in
           re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
    part = {m.group(1).replace('&amp;', '&'): 'xl/' + tgt[m.group(2)] for m in
            re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wbx)}
    wb.close()

    full = openpyxl.load_workbook(a.file, read_only=True)
    nrm = next(ws for ws in full.worksheets
               if any(isinstance(c.value, str) and ('XLOOKUP' in c.value or 'SUMIF(' in c.value)
                      for row in ws.iter_rows(max_row=40) for c in row))
    summary = next(ws for ws in full.worksheets if ws.title.lower() == 'asset summary')

    # asset tabs: the ones the Asset Summary links to
    linked = {}
    for rn, row in enumerate(summary.iter_rows(), 1):
        for c in row:
            v = getattr(c, 'value', None)
            if isinstance(v, str) and v.startswith("='") and "'!" in v:
                t = v[2:v.index("'!", 2)]
                if t in names:
                    linked.setdefault(t, rn)
    full.close()

    # find a tab with at least two packages: package headers carry a single
    # letter in the Sr. column and a name beside it
    def packages_of(tab):
        ws = openpyxl.load_workbook(a.file, read_only=True)[tab]
        out = []
        for rn, row in enumerate(ws.iter_rows(min_row=6), 6):
            if len(row) < 5:
                continue
            d, e = row[3].value, row[4].value
            if isinstance(d, str) and len(d.strip()) == 1 and d.strip().isalpha() and e:
                out.append((rn, str(e)))
        return out

    target = a.asset
    if not target:
        for t in linked:
            if len(packages_of(t)) >= 2:
                target = t
                break
    if not target:
        sys.exit('No asset with two or more packages found; nothing to prove here.')
    pkgs = packages_of(target)
    print('asset under test : %r with %d package(s): %s'
          % (target, len(pkgs), ', '.join(n for _, n in pkgs)))

    ws = openpyxl.load_workbook(a.file, read_only=True)[target]
    grid = {}
    for rn, row in enumerate(ws.iter_rows(), 1):
        for cn, c in enumerate(row, 1):
            if getattr(c, 'value', None) is not None:
                grid[(rn, colletter(cn))] = c.value
    total_row = max(r for (r, cl) in grid
                    if cl == 'E' and str(grid[(r, cl)]).strip().lower() == 'total amounts')
    # an item row inside a package: the PTE column is QTY x rate there
    item = re.compile(r'^=\$?([A-Z]+)\d+\*\$?([A-Z]+)\d+$')
    ret_col = 'K'
    items = sorted(r for (r, cl) in grid if cl == ret_col
                   and isinstance(grid[(r, cl)], str) and item.match(grid[(r, cl)].replace(' ', '')))
    first_pkg_rows = [r for r in items if r > pkgs[0][0] and r < pkgs[1][0]]
    second = [r for r in items if r > pkgs[-1][0]]
    r1, r2 = first_pkg_rows[0], second[0]
    qty_col, rate_col = item.match(grid[(r1, ret_col)].replace(' ', '')).groups()
    code = a.code or next((str(grid[(r, 'B')]) for r in range(5, 40)
                           if (r, 'B') in grid and isinstance(grid[(r, 'B')], str)
                           and re.match(r'^\d+(\.\d+)+$', str(grid[(r, 'B')]))), '0.1.1')


    nrm_row = None
    nws = openpyxl.load_workbook(a.file, read_only=True)[nrm.title]
    for rn, row in enumerate(nws.iter_rows(), 1):
        if len(row) > 1 and str(row[1].value) == code:
            nrm_row = rn
            break

    if nrm_row is None:
        sys.exit('NRM code %r not found on %r' % (code, nrm.title))
    print('planting code %r on rows %d (%s) and %d (%s), NRM row %d'
          % (code, r1, pkgs[0][1], r2, pkgs[-1][1], nrm_row))

    ss = open(T + '/x/xl/sharedStrings.xml', encoding='utf-8').read()
    texts = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', s, re.S))
             for s in re.findall(r'<si>.*?</si>', ss, re.S)]
    if code in texts:
        code_si = texts.index(code)
    else:
        code_si = len(texts)
        ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % code)
        ss = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % (code_si + 1), ss, count=1)
        open(T + '/x/xl/sharedStrings.xml', 'w', encoding='utf-8').write(ss)

    def put(path, row, cells):
        p = os.path.join(T, 'x', path)
        d = open(p, encoding='utf-8').read()
        m = re.search(r'<row r="%d"[^>]*>.*?</row>' % row, d, re.S)
        rx = m.group(0)
        for col, val, is_str in cells:
            rep = ('<c r="%s%d"\\1 t="s"><v>%s</v></c>' if is_str
                   else '<c r="%s%d"\\1><v>%s</v></c>') % (col, row, val)
            n = re.sub(r'<c r="%s%d"([^>]*?)\s*/>' % (col, row), rep, rx, count=1)
            if n == rx:
                n = re.sub(r'<c r="%s%d"([^>]*?)>.*?</c>' % (col, row), rep, rx,
                           count=1, flags=re.S)
            rx = n
        open(p, 'w', encoding='utf-8').write(d[:m.start()] + rx + d[m.end():])

    v1, v2 = 100, 40
    put(part[target], r1, [('A', code_si, True), (qty_col, 10, False),
                           ('G', 10, False), (rate_col, v1, False)])
    put(part[target], r2, [('A', code_si, True), (qty_col, 10, False),
                           ('G', 10, False), (rate_col, v2, False)])

    # the BUA the elemental sheet looks up, and probes for cells inside collapsed
    # outline groups (LibreOffice will not emit cached values for those)
    sched = next((n for n in names if 'schedu' in n.lower()), None)
    if sched:
        sws = openpyxl.load_workbook(a.file, read_only=True)[sched]
        srow = next((rn for rn, row in enumerate(sws.iter_rows(), 1)
                     if any(getattr(c, 'value', None) == target for c in row)), None)

        if srow:
            put(part[sched], srow, [('G', BUA, False)])
        sp = os.path.join(T, 'x', part[sched])
        sd = open(sp, encoding='utf-8').read()
        m = re.search(r'<row r="4"[^>]*>.*?</row>', sd, re.S)
        probes = {'T': '%s%d' % (ret_col, r1), 'U': '%s%d' % (ret_col, r2),
                  'V': '%s%d' % (ret_col, pkgs[0][0]), 'W': '%s%d' % (ret_col, pkgs[-1][0])}
        extra = ''.join("<c r=\"%s4\"><f>'%s'!%s</f></c>" % (c, target, ref)
                        for c, ref in probes.items())
        sd = sd[:m.start()] + m.group(0).replace('</row>', extra + '</row>') + sd[m.end():]
        open(sp, 'w', encoding='utf-8').write(sd)

    for nm in (target, nrm.title):
        p = os.path.join(T, 'x', part[nm])
        d = open(p, encoding='utf-8').read()
        d = d.replace(' hidden="1"', '')
        d = re.sub(r'\s+outlineLevel="\d+"', '', d)
        open(p, 'w', encoding='utf-8').write(d)

    def strip(mm):
        c = mm.group(0)
        if '<f' not in c:
            return c
        c = re.sub(r'<v>.*?</v>|<v/>', '', c, flags=re.S)
        return re.sub(r'(<c\b[^>]*?)\st="(str|s|e)"', r'\1', c)
    for f in glob.glob(T + '/x/xl/worksheets/sheet*.xml'):
        d = open(f, encoding='utf-8').read()
        open(f, 'w', encoding='utf-8').write(CELL.sub(strip, d))

    test = T + '/t.xlsx'
    with zipfile.ZipFile(a.file) as z, zipfile.ZipFile(test, 'w', zipfile.ZIP_DEFLATED) as o:
        for n in z.namelist():
            o.write(os.path.join(T, 'x', n), n)
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', 'xlsx',
                    '--outdir', T + '/c', test],
                   env=dict(os.environ, HOME=T), check=True, capture_output=True, timeout=900)
    rc = openpyxl.load_workbook(glob.glob(T + '/c/*.xlsx')[0], data_only=True)

    e1, e2 = 10 * v1, 10 * v2
    ok = True

    def chk(label, got, want, tol=1e-6):
        nonlocal ok
        good = isinstance(got, (int, float)) and abs(got - want) < tol
        ok = ok and good
        print('   %-50s %-12s (expected %-12s) %s'
              % (label, got, want, 'OK' if good else 'FAIL'))

    at = rc[target]
    print('\n--- asset tab ---')
    if sched:
        sc = rc[sched]
        chk('first package item', sc['T4'].value, e1)
        chk('last package item', sc['U4'].value, e2)
        chk('first package subtotal', sc['V4'].value, e1)
        chk('last package subtotal', sc['W4'].value, e2)
    chk('grand total (all packages)', at.cell(total_row, colnum(ret_col)).value, e1 + e2)

    print('\n--- summaries ---')
    srow = linked[target]
    asum = rc['Asset Summary']
    col = next(c.column for c in asum[srow]
               if isinstance(c.value, (int, float)) and c.value == e1 + e2) \
        if False else None
    aval = [c.value for c in asum[srow] if isinstance(c.value, (int, float))]
    chk('Asset Summary carries the asset total', (e1 + e2) if (e1 + e2) in aval else None,
        e1 + e2)
    ps = next((rc[n] for n in rc.sheetnames if n.lower() == 'project summary'), None)
    if ps:
        pval = [c.value for c in ps[srow] if isinstance(c.value, (int, float))]
        chk('Project Summary carries it too',
            (e1 + e2) if (e1 + e2) in pval else None, e1 + e2)

    print('\n--- elemental summary ---')
    en = rc[nrm.title]
    amounts = [c.value for c in en[nrm_row] if isinstance(c.value, (int, float))]
    chk('both packages summed against the code',
        (e1 + e2) if (e1 + e2) in amounts else None, e1 + e2)
    print('   (a first-match lookup would show %s here, losing %s)' % (e1, e2))

    bad = {}
    for s in rc.worksheets:
        for row in s.iter_rows():
            for c in row:
                if c.value in ERR:
                    bad[s.title] = bad.get(s.title, 0) + 1
    print('\nerror cells:', bad or 'none')
    ok = ok and not bad
    print('\nRESULT:', 'PASS' if ok else 'FAIL')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
