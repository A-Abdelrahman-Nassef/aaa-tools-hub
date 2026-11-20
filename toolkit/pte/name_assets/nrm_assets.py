#!/usr/bin/env python3
"""
NRM 1 Summary (BUA): name the asset column groups, point each group's lookup at
its own asset tab, choose which AAA amount column is pulled, and add groups for
any assets the sheet does not yet have.

Each asset occupies a repeating block of columns - Amount, Cost/m2, % - with the
asset name in a merged header on the header row and that asset's BUA in the %
cell of the same row. The layout is read from the file, not assumed.

Every cell written for a new group carries an EXPLICIT formula. Cloning cells
verbatim copies their shared-formula declarations (`<f t="shared" ref="E8:E9"
si="0">D8/F$2</f>`), which leaves two cells claiming the same si and makes the
new group evaluate the old group's columns. See references/excel-traps.md.

Usage:
  python3 nrm_assets.py --file IN.xlsx --out OUT.xlsx \
      --names "B01 & B02 Str&MEP" "B01 & B02 Arch" "External Shower and Toilette" \
      [--return-col K]      # K = AAA Pre-Tender Estimate, J = AAA Estimated
      [--label-suffix " - AAA"]
"""
import argparse, os, re, shutil, zipfile
from xml.sax.saxutils import escape as xesc
import openpyxl

ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
KEY_COL = 'A'                      # NRM Code column on each asset tab


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def colletter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Pkg:
    def __init__(self, path):
        self.dir = '/tmp/_nrmbuild'
        shutil.rmtree(self.dir, ignore_errors=True)
        os.makedirs(self.dir)
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            z.extractall(self.dir)

    def rd(self, p):
        return open(os.path.join(self.dir, p), encoding='utf-8').read()

    def wr(self, p, d):
        open(os.path.join(self.dir, p), 'w', encoding='utf-8').write(d)

    def sheets(self):
        wb, rels = self.rd('xl/workbook.xml'), self.rd('xl/_rels/workbook.xml.rels')
        tgt = {m.group(1): m.group(2) for m in
               re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
        return [(m.group(1).replace('&amp;', '&'), 'xl/' + tgt[m.group(2)]) for m in
                re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wb)]

    def save(self, out):
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                z.write(os.path.join(self.dir, n), n)


def detect(pkg):
    """Locate the NRM sheet and read its asset-group grammar off the sheet."""
    for name, part in pkg.sheets():
        d = pkg.rd(part)
        if 'XLOOKUP' not in d and 'SUMIF(' not in d:
            continue
        cols = {}
        for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*>\s*<f>[^<]*(?:XLOOKUP|SUMIF\()', d):
            cols[m.group(1)] = cols.get(m.group(1), 0) + 1
        amounts = sorted((c for c, n in cols.items() if n >= 5), key=colnum)
        if len(amounts) < 1:
            continue
        width = (colnum(amounts[1]) - colnum(amounts[0])) if len(amounts) > 1 else 3
        merges = re.findall(r'<mergeCell ref="([A-Z]+)(\d+):[A-Z]+\d+"/>', d)
        hdr = next((int(r) for c, r in merges if c == amounts[0]), None)
        if hdr is None:
            continue
        return dict(name=name, part=part, amounts=amounts, width=width, header_row=hdr)
    raise SystemExit('No NRM-style sheet with asset lookup groups found.')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--names', nargs='+', required=True)
    ap.add_argument('--return-col', default='K')
    ap.add_argument('--label-suffix', default=' - AAA')
    ap.add_argument('--aggregate', choices=['xlookup', 'sumif'], default='xlookup',
                    help='sumif adds every row matching the NRM code, so an asset '
                         'split into packages counts all of them; xlookup returns '
                         'only the first match')
    a = ap.parse_args()

    pkg = Pkg(a.file)
    info = detect(pkg)
    part, amounts, width, hdr = info['part'], info['amounts'], info['width'], info['header_row']
    have, want = len(amounts), len(a.names)
    print('NRM sheet    : %r' % info['name'])
    print('groups found : %d (amount columns %s, %d columns each, header row %d)'
          % (have, ', '.join(amounts), width, hdr))
    if want < have:
        raise SystemExit('Sheet has %d groups but only %d names given; refusing to drop groups.'
                         % (have, want))

    tabs = {n for n, _ in pkg.sheets()}
    for nm in a.names:
        if nm not in tabs:
            raise SystemExit('No tab named %r - run Step 01 first.' % nm)

    # group 1's formulas, shared ones already expanded per row by openpyxl
    src_ws = openpyxl.load_workbook(a.file)[info['name']]
    g1 = [colletter(colnum(amounts[0]) + i) for i in range(width)]
    G1 = {c: {r: src_ws['%s%d' % (c, r)].value for r in range(1, src_ws.max_row + 1)}
          for c in g1}

    def xlookup(row, tab):
        t = xesc(tab)
        if a.aggregate == 'sumif':
            # A blank criterion is treated as 0 by SUMIF and matches every blank
            # key, so a row with no NRM code would total the whole column. XLOOKUP
            # returned its not-found default instead, so the guard is needed only
            # here.
            return ("IF($B%d=\"\",0,SUMIF('%s'!$%s:$%s,$B%d,'%s'!$%s:$%s))"
                    % (row, t, KEY_COL, KEY_COL, row, t, a.return_col, a.return_col))
        return ("_xlfn.XLOOKUP($B%d,'%s'!$%s:$%s,'%s'!$%s:$%s,0,0,1)"
                % (row, t, KEY_COL, KEY_COL, t, a.return_col, a.return_col))

    sh = pkg.rd(part)
    rows = {int(m.group(1)): (m.start(), m.end(), m.group(0)) for m in ROW_RE.finditer(sh)}

    # ---- labels ------------------------------------------------------------
    ss = pkg.rd('xl/sharedStrings.xml')
    def si_for(text):
        nonlocal ss
        items = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', s, re.S))
                 for s in re.findall(r'<si>.*?</si>', ss, re.S)]
        if text in items:
            return items.index(text)
        ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % xesc(text))
        return len(items)
    labels = [si_for(nm + a.label_suffix) for nm in a.names]

    # ---- rewrite -----------------------------------------------------------
    n_new_cols = (want - have) * width
    new_start = colnum(amounts[-1]) + width
    stats = dict(repointed=0, written=0, copied=0, blank=0)
    out = sh
    for r in sorted(rows, reverse=True):
        s0, e0, rx = rows[r]
        new = rx

        # existing groups: repoint the lookup at that group's own tab
        for gi, ac in enumerate(amounts):
            if gi >= want:
                break
            m = re.search(r'(<c r="%s%d"[^>]*>)<f>[^<]*(?:XLOOKUP|SUMIF)\([^<]*</f>'
                          % (ac, r), new)
            if m:
                new = (new[:m.start()] + m.group(1) + '<f>' + xlookup(r, a.names[gi])
                       + '</f>' + new[m.end():])
                stats['repointed'] += 1

        # header labels for the groups that already exist
        if r == hdr:
            for gi, ac in enumerate(amounts[:want]):
                new = re.sub(r'(<c r="%s%d"[^>]*t="s"[^>]*>)<v>\d+</v>' % (ac, r),
                             r'\g<1><v>%d</v>' % labels[gi], new, count=1)

        # new groups, built from group 1 with the columns shifted
        add = ''
        for gi in range(have, want):
            off = (gi - have) * width + (new_start - colnum(g1[0]))
            cmap = {c: colletter(colnum(c) + off) for c in g1}
            shift = lambda f: re.sub(
                r'(\$?)([A-Z]+)(\$?)(\d+)',
                lambda m: (m.group(1) + cmap.get(m.group(2), m.group(2))
                           + m.group(3) + m.group(4)), f)
            for src in g1:
                dst = cmap[src]
                cm = re.search(r'<c r="%s%d"([^>]*?)(/>|>(.*?)</c>)' % (src, r), new, re.S)
                if not cm:
                    continue
                attrs, body = cm.group(1), cm.group(3)
                if r == hdr:
                    clean = attrs.replace(' t="s"', '')
                    if src == g1[0]:                      # the asset name
                        add += '<c r="%s%d"%s t="s"><v>%d</v></c>' % (dst, r, clean, labels[gi])
                    else:
                        hf = G1[src].get(r)
                        if isinstance(hf, str) and hf.startswith('='):
                            # the BUA cell is a formula in this template - it looks the
                            # asset up in the schedule - so it shifts like any other
                            add += ('<c r="%s%d"%s><f>%s</f><v>0</v></c>'
                                    % (dst, r, attrs, xesc(shift(hf[1:]))))
                        else:
                            add += '<c r="%s%d"%s/>' % (dst, r, clean)
                    continue
                if not body:
                    add += '<c r="%s%d"%s/>' % (dst, r, attrs)
                    stats['blank'] += 1
                    continue
                f = G1[src].get(r)
                if not (isinstance(f, str) and f.startswith('=')):
                    add += '<c r="%s%d"%s>%s</c>' % (dst, r, attrs, body)
                    stats['copied'] += 1
                    continue
                f = f[1:]
                if '#REF!' in f:                   # never propagate a broken formula
                    add += '<c r="%s%d"%s/>' % (dst, r, attrs.replace(' t="str"', ''))
                    stats['blank'] += 1
                    continue
                nf = (xlookup(r, a.names[gi])
                      if ('XLOOKUP' in f or 'SUMIF(' in f) else shift(f))
                add += '<c r="%s%d"%s><f>%s</f><v>0</v></c>' % (dst, r, attrs, nf)
                stats['written'] += 1
        if add:
            new = new.replace('</row>', add + '</row>')
        out = out[:s0] + new + out[e0:]

    # ---- spans, columns, merges, dimension ---------------------------------
    def fix_spans(m):
        b = m.group(0)
        cols = [colnum(c) for c in re.findall(r'<c r="([A-Z]+)\d+"', b)]
        if not cols:
            return re.sub(r'\s+spans="[^"]*"', '', b)
        return re.sub(r'spans="[^"]*"', 'spans="%d:%d"' % (min(cols), max(cols)), b)
    out = ROW_RE.sub(fix_spans, out)

    if n_new_cols:
        defs = re.findall(r'<col [^>]*/>', out)
        tpl = {}
        for d in defs:
            mn = int(re.search(r'min="(\d+)"', d).group(1))
            mx = int(re.search(r'max="(\d+)"', d).group(1))
            for c in g1:
                if mn <= colnum(c) <= mx:
                    tpl[c] = d
        rebuilt = []
        for d in defs:                              # push everything at or past the new block
            mn = int(re.search(r'min="(\d+)"', d).group(1))
            mx = int(re.search(r'max="(\d+)"', d).group(1))
            if mn >= new_start:
                d = re.sub(r'min="\d+"', 'min="%d"' % (mn + n_new_cols), d)
                d = re.sub(r'max="\d+"', 'max="%d"' % min(mx + n_new_cols, 16384), d)
            rebuilt.append(d)
        block = ''
        for gi in range(have, want):
            for i, c in enumerate(g1):
                n = new_start + (gi - have) * width + i
                t = tpl.get(c, '<col min="%d" max="%d" width="10" customWidth="1"/>' % (n, n))
                t = re.sub(r'min="\d+"', 'min="%d"' % n, t)
                t = re.sub(r'max="\d+"', 'max="%d"' % n, t)
                block += t
        out = out.replace(''.join(defs), _splice(rebuilt, block, new_start), 1)

        # merged header for each new group's name
        mm = re.search(r'<mergeCells count="(\d+)">', out)
        if mm:
            extra = ''
            for gi in range(have, want):
                s = new_start + (gi - have) * width
                extra += '<mergeCell ref="%s%d:%s%d"/>' % (colletter(s), hdr,
                                                           colletter(s + width - 2), hdr)
            out = out.replace(mm.group(0),
                              '<mergeCells count="%d">' % (int(mm.group(1)) + (want - have)))
            out = out.replace('</mergeCells>', extra + '</mergeCells>')

    last_col = colletter(new_start + n_new_cols - 1) if n_new_cols else amounts[-1]
    out = re.sub(r'(<dimension ref="[A-Z]+\d+:)[A-Z]+(\d+"/>)', r'\g<1>%s\g<2>' % last_col, out)
    pkg.wr(part, out)
    pkg.wr('xl/sharedStrings.xml', ss)

    wb = pkg.rd('xl/workbook.xml')
    wb = re.sub(r"(Print_Area\" localSheetId=\"\d+\">'%s'!\$[A-Z]+\$\d+:\$)[A-Z]+(\$\d+)"
                % re.escape(info['name']), r'\g<1>%s\g<2>' % last_col, wb)
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\g<1> fullCalcOnLoad="1"/>', wb, count=1)
    pkg.wr('xl/workbook.xml', wb)

    # ---- the counts Excel refuses to open a file over ----------------------
    ss2 = pkg.rd('xl/sharedStrings.xml')
    uniq = len(re.findall(r'<si[ >]', ss2))
    refs = sum(len(re.findall(r't="s"', pkg.rd(p))) for _, p in pkg.sheets())
    ss2 = re.sub(r'count="\d+" uniqueCount="\d+"',
                 'count="%d" uniqueCount="%d"' % (refs, uniq), ss2, count=1)
    pkg.wr('xl/sharedStrings.xml', ss2)

    pkg.save(a.out)
    print('lookups repointed  : %d' % stats['repointed'])
    print('new-group formulas : %d explicit, %d literals, %d blanks'
          % (stats['written'], stats['copied'], stats['blank']))
    print('return column      : %s' % a.return_col)
    print('shared strings     : count=%d uniqueCount=%d' % (refs, uniq))
    print('saved %s (%d bytes)' % (a.out, os.path.getsize(a.out)))


def _splice(rebuilt, block, new_start):
    """Put the new <col> definitions in ascending order among the existing ones."""
    outp = ''
    placed = False
    for d in rebuilt:
        mn = int(re.search(r'min="(\d+)"', d).group(1))
        if not placed and mn >= new_start:
            outp += block
            placed = True
        outp += d
    if not placed:
        outp += block
    return outp


if __name__ == '__main__':
    main()
