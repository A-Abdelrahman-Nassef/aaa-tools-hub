#!/usr/bin/env python3
"""
Name the assets across the CLIENT/PROJECT R00 workbook, with per-asset packages.

Sheets touched:
  <asset tabs>              one per asset, from the packaged or unpackaged
                            template depending on whether that asset has packages
  Scheduale of Assets       asset name per row
  Asset Summary            name + links to that asset's own Total Amounts row
  Project Summary          name (its amount columns already read Asset Summary)
  Elemental Summary (NRM01) a column group per asset, keyed on NRM code

Usage:
  python3 build_r00.py --file IN.xlsx --out OUT.xlsx \
      --assets "Building 01: Structure, Architecture, Plumbing, Electrical" \
               "Building 02: Structure, MEP" \
               "External Works"
  (an asset with no colon has no packages and uses the unpackaged template)
"""
import argparse, os, re, shutil, sys, uuid, zipfile
from xml.sax.saxutils import escape as xesc
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pkgblock import build_asset_rows, FIRST, BLOCK, TPL_TOTAL

ROW = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
NS_WS = ('http://schemas.openxmlformats.org/officeDocument/2006/'
         'relationships/worksheet')
CT_WS = ('application/vnd.openxmlformats-officedocument.spreadsheetml.'
         'worksheet+xml')


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
        self.dir = '/tmp/_r00'
        shutil.rmtree(self.dir, ignore_errors=True)
        os.makedirs(self.dir)
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            z.extractall(self.dir)

    def rd(self, p):
        return open(os.path.join(self.dir, p), encoding='utf-8').read()

    def wr(self, p, d):
        open(os.path.join(self.dir, p), 'w', encoding='utf-8').write(d)
        if p not in self.names:
            self.names.append(p)

    def sheets(self):
        wb, rels = self.rd('xl/workbook.xml'), self.rd('xl/_rels/workbook.xml.rels')
        tgt = {m.group(1): m.group(2) for m in
               re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
        return [(m.group(1).replace('&amp;', '&'), 'xl/' + tgt[m.group(2)],
                 m.group(2)) for m in
                re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wb)]

    def part_of(self, name):
        return next(p for n, p, _ in self.sheets() if n == name)

    def save(self, out):
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                z.write(os.path.join(self.dir, n), n)


def parse_assets(items):
    out = []
    for s in items:
        if ':' in s:
            name, pk = s.split(':', 1)
            pkgs = [p.strip() for p in pk.split(',') if p.strip()]
        else:
            name, pkgs = s, []
        out.append((name.strip(), pkgs))
    return out


def set_cell(rowxml, col, row, text=None, formula=None, si=None):
    """Replace one cell's content, keeping its style."""
    pat = re.compile(r'<c r="%s%d"([^>]*?)(?:/>|>(.*?)</c>)' % (col, row), re.S)
    m = pat.search(rowxml)
    if not m:
        return rowxml
    attrs = re.sub(r'\st="(s|str|inlineStr|e)"', '', m.group(1))
    if formula is not None:
        new = '<c r="%s%d"%s><f>%s</f><v>0</v></c>' % (col, row, attrs, xesc(formula))
    elif si is not None:
        new = '<c r="%s%d"%s t="s"><v>%d</v></c>' % (col, row, attrs, si)
    elif text is not None:
        new = '<c r="%s%d"%s t="inlineStr"><is><t>%s</t></is></c>' % (
            col, row, attrs, xesc(text))
    else:
        new = '<c r="%s%d"%s/>' % (col, row, attrs)
    return rowxml[:m.start()] + new + rowxml[m.end():]


def edit_rows(xml, fn):
    """Apply fn(row_number, row_xml) -> row_xml to every row, back to front."""
    rows = [(int(m.group(1)), m.start(), m.end(), m.group(0)) for m in ROW.finditer(xml)]
    for r, s, e, rx in reversed(rows):
        new = fn(r, rx)
        if new is not None and new != rx:
            xml = xml[:s] + new + xml[e:]
    return xml


def fix_spans(xml):
    def f(m):
        b = m.group(0)
        cols = [colnum(c) for c in re.findall(r'<c r="([A-Z]+)\d+"', b)]
        if not cols:
            return re.sub(r'\s+spans="[^"]*"', '', b)
        return re.sub(r'spans="[^"]*"', 'spans="%d:%d"' % (min(cols), max(cols)), b)
    return ROW.sub(f, xml)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--assets', nargs='+', required=True)
    a = ap.parse_args()
    assets = parse_assets(a.assets)

    for nm, _ in assets:
        if len(nm) > 31 or re.search(r'[:\\/?*\[\]]', nm):
            sys.exit('Invalid tab name: %r (max 31 chars, no : \\ / ? * [ ])' % nm)
    if len({n for n, _ in assets}) != len(assets):
        sys.exit('Asset names must be unique.')

    pkg = Pkg(a.file)
    src = openpyxl.load_workbook(a.file)
    names = [n for n, _, _ in pkg.sheets()]
    tpl_pkg = next(n for n in names if '{with package}' in n)
    tpl_flat = next(n for n in names if '{no package}' in n)
    print('templates   : %r (packaged), %r (flat)' % (tpl_pkg, tpl_flat))
    print('assets      : %d' % len(assets))
    for nm, p in assets:
        print('   %-28s %s' % (nm, ', '.join(p) if p else '(no packages)'))

    # ---------- template row/formula/value maps ----------------------------
    def maps(sheetname):
        part = pkg.part_of(sheetname)
        rows = {int(m.group(1)): m.group(0) for m in ROW.finditer(pkg.rd(part))}
        ws = src[sheetname]
        f, v = {}, {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith('='):
                    f[(c.column_letter, c.row)] = c.value
                else:
                    v[(c.column_letter, c.row)] = c.value
        return part, rows, f, v

    part_pkg, rows_pkg, f_pkg, v_pkg = maps(tpl_pkg)
    part_flat, rows_flat, f_flat, v_flat = maps(tpl_flat)
    flat_total = max(r for (c, r) in f_flat if c == 'J' and
                     f_flat[(c, r)].startswith('=SUBTOTAL(9,J6'))

    # ---------- build one worksheet part per asset -------------------------
    used = {int(m.group(1)) for m in re.finditer(r'worksheets/sheet(\d+)\.xml',
                                                 pkg.rd('xl/_rels/workbook.xml.rels'))}
    rid_max = max(int(m.group(1)) for m in
                  re.finditer(r'Id="rId(\d+)"', pkg.rd('xl/_rels/workbook.xml.rels')))
    sid_max = max(int(m.group(1)) for m in
                  re.finditer(r'sheetId="(\d+)"', pkg.rd('xl/workbook.xml')))

    new_parts, totals = [], {}
    for idx, (nm, pkgs) in enumerate(assets):
        if pkgs:
            body, total = build_asset_rows(rows_pkg, f_pkg, v_pkg, pkgs)
            base = pkg.rd(part_pkg)
        else:
            seeded = [r for (c, r) in f_flat
                      if c == 'J' and str(f_flat[(c, r)]).startswith('=$G')]
            def clear_flat(rx, rn):
                if rn not in seeded:
                    return rx
                for col in ('A', 'G', 'H', 'I'):
                    rx = re.sub(r'<c r="%s%d"([^>]*?)(?:/>|>.*?</c>)' % (col, rn),
                                lambda m: '<c r="%s%d"%s/>'
                                % (col, rn, re.sub(r'\st="[a-z]+"', '', m.group(1))),
                                rx, count=1, flags=re.S)
                return rx
            body = ''.join(clear_flat(rows_flat[r], r) for r in sorted(rows_flat))
            total = flat_total
            base = pkg.rd(part_flat)
        totals[nm] = total
        head = base[:base.index('<sheetData>') + len('<sheetData>')]
        tail = base[base.index('</sheetData>'):]
        sheet = head + body + tail
        sheet = re.sub(r'xr:uid="\{[0-9A-Fa-f-]+\}"',
                       lambda m: 'xr:uid="{%s}"' % str(uuid.uuid4()).upper(), sheet)
        last = max(int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', sheet))
        sheet = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>',
                       r'<dimension ref="A1:\g<1>%d"/>' % last, sheet)
        # only the autofilter spans the body; the merges are all header
        # furniture on rows 1-5 and must be left exactly as they are
        sheet = re.sub(r'(<autoFilter[^>]*ref="[A-Z]+\d+:[A-Z]+)\d+',
                       lambda m: m.group(1) + str(total - 1), sheet)
        sheet = fix_spans(sheet)
        n = max(used) + 1
        used.add(n)
        p = 'xl/worksheets/sheet%d.xml' % n
        pkg.wr(p, sheet)
        new_parts.append((nm, sid_max + 1 + idx, 'rId%d' % (rid_max + 1 + idx), p))
        print('   built %-28s %d package(s), total row %d' % (nm, max(len(pkgs), 1), total))

    # ---------- register the new sheets, drop the templates ----------------
    ct = pkg.rd('[Content_Types].xml')
    ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)
    ct = ct.replace('<Override PartName="/xl/theme/theme1.xml"',
                    ''.join('<Override PartName="/%s" ContentType="%s"/>' % (p, CT_WS)
                            for _, _, _, p in new_parts)
                    + '<Override PartName="/xl/theme/theme1.xml"')
    pkg.wr('[Content_Types].xml', ct)

    rels = pkg.rd('xl/_rels/workbook.xml.rels')
    rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', rels)
    rels = rels.replace('</Relationships>',
                        ''.join('<Relationship Id="%s" Type="%s" Target="%s"/>'
                                % (rid, NS_WS, p.replace('xl/', ''))
                                for _, _, rid, p in new_parts) + '</Relationships>')
    pkg.wr('xl/_rels/workbook.xml.rels', rels)
    if os.path.exists(os.path.join(pkg.dir, 'xl/calcChain.xml')):
        os.remove(os.path.join(pkg.dir, 'xl/calcChain.xml'))
        pkg.names.remove('xl/calcChain.xml')

    wb = pkg.rd('xl/workbook.xml')
    sheet_els = re.findall(r'<sheet [^>]*/>', wb)
    keep = [s for s in sheet_els if '{with package}' not in s and '{no package}' not in s]
    added = ''.join('<sheet name="%s" sheetId="%d" r:id="%s"/>'
                    % (xesc(nm), sid, rid) for nm, sid, rid, _ in new_parts)
    wb = wb.replace(''.join(sheet_els), ''.join(keep) + added)
    # sheet-local defined names are keyed by tab position; the templates are gone
    for tn in (tpl_pkg, tpl_flat):
        wb = re.sub(r'<definedName [^>]*>\s*\'?%s\'?![^<]*</definedName>'
                    % re.escape(tn), '', wb)
    wb = re.sub(r'<definedName [^>]*localSheetId="[67]"[^>]*>.*?</definedName>', '', wb, flags=re.S)
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\g<1> fullCalcOnLoad="1"/>', wb, count=1)
    pkg.wr('xl/workbook.xml', wb)

    # ---------- shared strings helper --------------------------------------
    ss = pkg.rd('xl/sharedStrings.xml')

    def si_for(text):
        nonlocal ss
        items = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', s, re.S))
                 for s in re.findall(r'<si>.*?</si>', ss, re.S)]
        if text in items:
            return items.index(text)
        ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % xesc(text))
        return len(items)

    # ---------- Scheduale of Assets ----------------------------------------
    part = pkg.part_of('Scheduale of Assets')
    d = pkg.rd(part)
    sched_rows = sorted(r for r in
                        {int(m.group(1)) for m in ROW.finditer(d)} if r >= 6)

    def sched(r, rx):
        i = sched_rows.index(r) if r in sched_rows else -1
        if 0 <= i < len(assets):
            return set_cell(rx, 'D', r, si=si_for(assets[i][0]))
        return rx
    pkg.wr(part, edit_rows(d, sched))
    print('Scheduale of Assets : %d name(s) written' % len(assets))

    # ---------- Asset Summary ----------------------------------------------
    part = pkg.part_of('Asset Summary')
    d = pkg.rd(part)
    aws = src['Asset Summary']
    total_row = max(r for r in range(1, aws.max_row + 1)
                    if str(aws.cell(r, 3).value or '').strip().lower() == 'total amounts')
    asset_rows = [r for r in range(6, total_row)
                  if isinstance(aws.cell(r, 2).value, (int, float))
                  and aws.cell(r, 3).value]
    # learn the column mapping from the row that is already linked
    linked = next(r for r in asset_rows
                  if isinstance(aws.cell(r, 9).value, str) and "'!" in str(aws.cell(r, 9).value))
    pairs = []
    for c in range(9, aws.max_column + 1):
        v = aws.cell(linked, c).value
        m = isinstance(v, str) and re.match(r"='[^']+'!\$([A-Z]+)\$(\d+)$", v)
        if m:
            pairs.append((aws.cell(linked, c).column_letter, m.group(1)))
    print('Asset Summary       : rows %s..%s, total row %d, %d linked columns'
          % (asset_rows[0], asset_rows[-1], total_row, len(pairs)))

    def asum(r, rx):
        if r == total_row:                       # weight total stopped short of the body
            return re.sub(r'(<c r="A%d"[^>]*>)<f>SUBTOTAL\(9,A6:A\d+\)</f>' % r,
                          r'\g<1><f>SUBTOTAL(9,A6:A%d)</f>' % (total_row - 1), rx)
        if r in asset_rows:
            i = asset_rows.index(r)
            rx = set_cell(rx, 'B', r, formula=None) if False else rx
            rx = re.sub(r'(<c r="B%d"[^>]*>)<v>\d+</v>' % r,
                        r'\g<1><v>%d</v>' % (i + 1), rx)     # Sr. - B28 read 1, not 11
            if i < len(assets):
                nm = assets[i][0]
                rx = set_cell(rx, 'C', r, si=si_for(nm))
                for sc, fc in pairs:
                    rx = set_cell(rx, sc, r,
                                  formula="'%s'!$%s$%d" % (nm, fc, totals[nm]))
        return rx
    pkg.wr(part, edit_rows(d, asum))

    # ---------- Project Summary --------------------------------------------
    part = pkg.part_of('Project Summary')
    d = pkg.rd(part)
    pws = src['Project Summary']
    ptotal = max(r for r in range(1, pws.max_row + 1)
                 if str(pws.cell(r, 3).value or '').strip().lower() == 'total amounts')
    prows = [r for r in range(6, ptotal)
             if isinstance(pws.cell(r, 2).value, (int, float)) and pws.cell(r, 3).value]

    def psum(r, rx):
        if r == ptotal:
            return re.sub(r'(<c r="A%d"[^>]*>)<f>SUBTOTAL\(9,A6:A\d+\)</f>' % r,
                          r'\g<1><f>SUBTOTAL(9,A6:A%d)</f>' % (ptotal - 1), rx)
        if r in prows:
            i = prows.index(r)
            rx = re.sub(r'(<c r="B%d"[^>]*>)<v>\d+</v>' % r, r'\g<1><v>%d</v>' % (i + 1), rx)
            if i < len(assets):
                rx = set_cell(rx, 'C', r, si=si_for(assets[i][0]))
        return rx
    pkg.wr(part, edit_rows(d, psum))
    print('Project Summary     : rows %s..%s, total row %d' % (prows[0], prows[-1], ptotal))

    pkg.wr('xl/sharedStrings.xml', ss)
    ss2 = pkg.rd('xl/sharedStrings.xml')
    uniq = len(re.findall(r'<si[ >]', ss2))
    refs = sum(len(re.findall(r't="s"', pkg.rd(p))) for _, p, _ in pkg.sheets())
    pkg.wr('xl/sharedStrings.xml',
           re.sub(r'count="\d+" uniqueCount="\d+"',
                  'count="%d" uniqueCount="%d"' % (refs, uniq), ss2, count=1))

    # The template ships cached results computed from its own sample data. Those
    # caches are stale the moment the sample data goes, and a viewer that does not
    # recalculate would show them as real money. Drop every cached formula result
    # so the first calculation is the only one anyone sees.
    CELL = re.compile(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.S)

    def drop_cache(m):
        c = m.group(0)
        if '<f' not in c:
            return c
        c = re.sub(r'<v>.*?</v>|<v/>', '', c, flags=re.S)
        return re.sub(r'(<c\b[^>]*?)\st="(str|e)"', r'\1', c)

    for _, part_p, _ in pkg.sheets():
        pkg.wr(part_p, CELL.sub(drop_cache, pkg.rd(part_p)))

    pkg.save(a.out)
    print('saved %s (%d bytes)' % (a.out, os.path.getsize(a.out)))


if __name__ == '__main__':
    main()
