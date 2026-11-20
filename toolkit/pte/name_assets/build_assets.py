#!/usr/bin/env python3
"""
Name the template asset sheet and clone it once per asset, then re-link the
Asset Summary.

Everything happens on the raw OOXML package rather than through openpyxl,
because these workbooks routinely carry thousands of legacy defined names,
web extensions, VML drawings, printer settings and outline grouping that a
load/save round-trip silently drops.

Usage:
  python3 build_assets.py --file IN.xlsx --out OUT.xlsx \
      --names "B01 & B02 Str&MEP" "B01 & B02 Arch" "External Shower and Toilette" \
      [--rows 15]          # grow the summary to this many asset rows
      [--summary-sheet "Asset Summary"] [--template-sheet "Asset #01"]
"""
import argparse, os, re, shutil, uuid, zipfile
from xml.sax.saxutils import escape as xesc

NS_WS = ('http://schemas.openxmlformats.org/officeDocument/2006/'
         'relationships/worksheet')
CT_WS = ('application/vnd.openxmlformats-officedocument.spreadsheetml.'
         'worksheet+xml')


# ----------------------------------------------------------------- utilities
def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def rows_of(xml):
    """{row number: (start, end, xml)} for every <row> element."""
    return {int(m.group(1)): (m.start(), m.end(), m.group(0))
            for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', xml, re.S)}


def renumber(rowxml, old, new):
    rowxml = re.sub(r'(<row r=")%d(")' % old, r'\g<1>%d\g<2>' % new, rowxml)
    return re.sub(r'(<c r="[A-Z]+)%d(")' % old, r'\g<1>%d\g<2>' % new, rowxml)


def sst_text(si):
    """Plain text of a <si> element, concatenating rich-text runs."""
    return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))


class Pkg:
    """A working copy of the xlsx package, edited as text."""

    def __init__(self, path):
        self.dir = '/tmp/_assetbuild'
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

    def drop(self, p):
        if p in self.names:
            self.names.remove(p)
        fp = os.path.join(self.dir, p)
        if os.path.exists(fp):
            os.remove(fp)

    def save(self, out):
        os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                z.write(os.path.join(self.dir, n), n)


# --------------------------------------------------------------- discovery
def sheet_index(pkg):
    """[(name, sheetId, rid, part_path, state)] in tab order."""
    wb = pkg.rd('xl/workbook.xml')
    rels = pkg.rd('xl/_rels/workbook.xml.rels')
    target = {m.group(1): m.group(2) for m in
              re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="(worksheets/[^"]+)"', rels)}
    out = []
    for m in re.finditer(r'<sheet name="([^"]*)" sheetId="(\d+)"([^>]*?)r:id="([^"]+)"', wb):
        name, sid, mid, rid = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        state = 'hidden' if 'hidden' in mid else 'visible'
        out.append((name, sid, rid, 'xl/' + target[rid], state))
    return out


def find_summary_and_template(pkg, sheets, want_summary=None, want_template=None):
    """The summary sheet is the one whose rows link to another sheet's total row."""
    if want_summary and want_template:
        s = next(x for x in sheets if x[0] == want_summary)
        t = next(x for x in sheets if x[0] == want_template)
        return s, t, None
    for sh in sheets:
        xml = pkg.rd(sh[3])
        for rn, (_, _, rx) in sorted(rows_of(xml).items()):
            refs = re.findall(r"<f>'([^']+)'!\$([A-Z]+)\$(\d+)</f>", rx)
            if len(refs) >= 4 and len({r[0] for r in refs}) == 1:
                tname = refs[0][0].replace('&amp;', '&')
                t = next((x for x in sheets if x[0] == tname), None)
                if t:
                    return sh, t, rn
    raise SystemExit('Could not find a summary sheet linking to a template tab. '
                     'Pass --summary-sheet and --template-sheet explicitly.')


def summary_layout(xml):
    """Asset rows, total row, and the repeating step, read from the sheet itself."""
    rows = rows_of(xml)
    total = max((r for r, (_, _, x) in rows.items()
                 if re.search(r'<c r="A%d"[^>]*>\s*<f>SUBTOTAL' % r, x)), default=None)
    if total is None:
        raise SystemExit('No "Total Amounts" SUBTOTAL row found on the summary sheet.')
    assets = sorted(r for r, (_, _, x) in rows.items()
                    if 5 < r < total
                    and re.search(r'<c r="B%d"[^>]*>\s*<v>\d+</v>' % r, x)
                    and re.search(r'<c r="C%d"[^>]*t="s"[^>]*>\s*<v>\d+</v>' % r, x))
    return assets, total, (assets[1] - assets[0] if len(assets) > 1 else 2)


# --------------------------------------------------------------- operations
def extend_summary(pkg, sheet_part, want):
    """Clone the (asset row, spacer row) unit until the sheet holds `want` assets."""
    sh, ss = pkg.rd(sheet_part), pkg.rd('xl/sharedStrings.xml')
    assets, total, step = summary_layout(sh)
    if want <= len(assets):
        return 0
    rows = rows_of(sh)
    last_body, spacer_tpl = total - 1, rows[total - 1][2]
    asset_tpl, labels = rows[assets[-1]][2], [sst_text(s) for s in
                                              re.findall(r'<si>.*?</si>', ss, re.S)]
    block, cur = '', last_body
    for i in range(want - len(assets)):
        sr = len(assets) + i + 1
        label = 'Asset #%02d' % sr
        if label in labels:
            idx = labels.index(label)
        else:
            labels.append(label)
            idx = len(labels) - 1
            ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % label)
        cur += 1
        r = cur
        a = renumber(asset_tpl, assets[-1], r)
        a = re.sub(r'(<c r="B%d"[^>]*>)<v>\d+</v>' % r, r'\g<1><v>%d</v>' % sr, a)
        a = re.sub(r'(<c r="C%d"[^>]*>)<v>\d+</v>' % r, r'\g<1><v>%d</v>' % idx, a)
        for col, f in (('E', 'IFERROR(L{r}/$D{r}, 0)'), ('G', 'SUM(H{r}:I{r})'),
                       ('J', 'IFERROR(L{r}/$G{r}, 0)')):
            a = re.sub(r'(<c r="%s%d"[^>]*>)<f>.*?</f>' % (col, r),
                       lambda m: m.group(1) + '<f>' + f.format(r=r) + '</f>', a, count=1)
        block += a
        cur += 1
        block += renumber(spacer_tpl, last_body, cur)
    new_last, new_total = cur, cur + 1

    tot = renumber(rows[total][2], total, new_total)
    tot = re.sub(r'(SUBTOTAL\(9,[A-Z]+\d+:[A-Z]+)%d(\))' % last_body,
                 r'\g<1>%d\g<2>' % new_last, tot)
    tot = tot.replace('ref="L%d:AF%d"' % (total, total), 'ref="L%d:AF%d"' % (new_total, new_total))
    sh = sh[:rows[last_body][1]] + block + tot + sh[rows[total][1]:]

    sh = sh.replace('$L$%d' % total, '$L$%d' % new_total)                 # weight denominators
    sh = sh.replace('ref="A%d:A%d"' % (assets[0] + 1, last_body),
                    'ref="A%d:A%d"' % (assets[0] + 1, new_last))          # shared formula range
    sh = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>',
                r'<dimension ref="A1:\g<1>%d"/>' % new_total, sh)
    pkg.wr(sheet_part, sh)
    pkg.wr('xl/sharedStrings.xml', ss)

    wb = pkg.rd('xl/workbook.xml')
    wb = re.sub(r"(Print_Area\" localSheetId=\"\d+\">'[^']*'!\$A\$1:\$[A-Z]+\$)%d\b" % total,
                r'\g<1>%d' % new_total, wb)
    pkg.wr('xl/workbook.xml', wb)
    return want - len(assets)


def clone_tabs(pkg, sheets, template, names):
    """Rename the template to names[0] and add a clone for every other name."""
    old = template[0]
    t_idx = sheets.index(template)
    n_new = len(names) - 1

    used = {int(m.group(1)) for m in
            re.finditer(r'worksheets/sheet(\d+)\.xml', pkg.rd('xl/_rels/workbook.xml.rels'))}
    rid_max = max(int(m.group(1)) for m in
                  re.finditer(r'Id="rId(\d+)"', pkg.rd('xl/_rels/workbook.xml.rels')))
    sid_max = max(s[1] for s in sheets)
    ps_used = {int(m.group(1)) for m in
               re.finditer(r'printerSettings(\d+)\.bin', ' '.join(pkg.names))} or {0}

    src = pkg.rd(template[3])
    t_rels_path = 'xl/worksheets/_rels/%s.rels' % os.path.basename(template[3])
    t_rels = pkg.rd(t_rels_path) if t_rels_path in pkg.names else None

    new_sheets = []
    for i, nm in enumerate(names[1:]):
        n = max(used) + 1
        used.add(n)
        part = 'xl/worksheets/sheet%d.xml' % n
        pkg.wr(part, re.sub(r'xr:uid="\{[0-9A-Fa-f-]+\}"',
                            lambda m: 'xr:uid="{%s}"' % str(uuid.uuid4()).upper(), src))
        if t_rels:                       # give the clone its own printer-settings part
            r = t_rels
            for pm in re.finditer(r'printerSettings(\d+)\.bin', t_rels):
                newps = max(ps_used) + 1
                ps_used.add(newps)
                shutil.copy(os.path.join(pkg.dir, 'xl/printerSettings/printerSettings%s.bin' % pm.group(1)),
                            os.path.join(pkg.dir, 'xl/printerSettings/printerSettings%d.bin' % newps))
                pkg.names.append('xl/printerSettings/printerSettings%d.bin' % newps)
                r = r.replace('printerSettings%s.bin' % pm.group(1),
                              'printerSettings%d.bin' % newps)
            pkg.wr('xl/worksheets/_rels/sheet%d.xml.rels' % n, r)
        new_sheets.append((nm, sid_max + 1 + i, 'rId%d' % (rid_max + 1 + i), part))

    ct = pkg.rd('[Content_Types].xml')
    ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)
    ct = ct.replace('<Override PartName="/xl/theme/theme1.xml"',
                    ''.join('<Override PartName="/%s" ContentType="%s"/>' % (p, CT_WS)
                            for _, _, _, p in new_sheets)
                    + '<Override PartName="/xl/theme/theme1.xml"')
    pkg.wr('[Content_Types].xml', ct)

    rels = pkg.rd('xl/_rels/workbook.xml.rels')
    rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', rels)
    rels = rels.replace('</Relationships>',
                        ''.join('<Relationship Id="%s" Type="%s" Target="%s"/>'
                                % (rid, NS_WS, p.replace('xl/', ''))
                                for _, _, rid, p in new_sheets) + '</Relationships>')
    pkg.wr('xl/_rels/workbook.xml.rels', rels)
    pkg.drop('xl/calcChain.xml')

    wb = pkg.rd('xl/workbook.xml')
    # sheets after the template shift right; bump their sheet-local defined names first
    for old_i in range(len(sheets) - 1, t_idx, -1):
        wb = wb.replace('localSheetId="%d"' % old_i, 'localSheetId="%d"' % (old_i + n_new))
    ent = re.search(r'<sheet name="%s" sheetId="%d"[^>]*/>' % (re.escape(xesc(old)), template[1]), wb)
    wb = wb.replace(ent.group(0),
                    ent.group(0).replace('name="%s"' % xesc(old), 'name="%s"' % xesc(names[0]))
                    + ''.join('<sheet name="%s" sheetId="%d" r:id="%s"/>'
                              % (xesc(nm), sid, rid) for nm, sid, rid, _ in new_sheets))
    # the template's own defined names follow the rename; clones get copies
    tmpl_dn = re.findall(r'<definedName [^>]*localSheetId="%d"[^>]*>.*?</definedName>' % t_idx,
                         wb, re.S)
    add = ''
    for j, (nm, _, _, _) in enumerate(new_sheets):
        for dn in tmpl_dn:
            add += (dn.replace('localSheetId="%d"' % t_idx, 'localSheetId="%d"' % (t_idx + 1 + j))
                      .replace("'%s'!" % xesc(old), "'%s'!" % xesc(nm)))
    wb = wb.replace('</definedNames>', add + '</definedNames>')
    wb = wb.replace("'%s'!" % xesc(old), "'%s'!" % xesc(names[0]))
    wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\g<1> fullCalcOnLoad="1"/>', wb, count=1)
    pkg.wr('xl/workbook.xml', wb)

    # every other sheet that referenced the old tab name follows the rename
    followed = 0
    for part in [s[3] for s in sheet_index(pkg)]:
        x = pkg.rd(part)
        k = x.count("'%s'!" % xesc(old))
        if k:
            pkg.wr(part, x.replace("'%s'!" % xesc(old), "'%s'!" % xesc(names[0])))
            followed += k
    return new_sheets, followed


def link_summary(pkg, summary_part, names, linked_row):
    """Write each asset's name and its links to that tab's Total Amounts row."""
    sh, ss = pkg.rd(summary_part), pkg.rd('xl/sharedStrings.xml')
    assets, total, _ = summary_layout(sh)
    rows = rows_of(sh)

    # learn the column mapping from the row that is already linked
    src = rows[linked_row][2]
    pairs = [(m.group(1), m.group(2), int(m.group(3))) for m in
             re.finditer(r'<c r="([A-Z]+)%d"[^>]*>\s*<f>\'[^\']+\'!\$([A-Z]+)\$(\d+)</f>'
                         % linked_row, src)]
    if not pairs:
        raise SystemExit('No linked asset row to learn the column mapping from.')
    print('   mapping learned from row %d: %d columns, target row %d'
          % (linked_row, len(pairs), pairs[0][2]))

    si = [sst_text(s) for s in re.findall(r'<si>.*?</si>', ss, re.S)]
    for i, nm in enumerate(names):
        r = assets[i]
        rowxml = rows[r][2]
        # label
        if nm in si:
            idx = si.index(nm)
        else:
            si.append(nm)
            idx = len(si) - 1
            ss = ss.replace('</sst>', '<si><t>%s</t></si></sst>' % xesc(nm))
        rowxml = re.sub(r'(<c r="C%d"[^>]*t="s"[^>]*>)<v>\d+</v>' % r,
                        r'\g<1><v>%d</v>' % idx, rowxml)
        # links
        for sc, fc, trow in pairs:
            f = "<f>'%s'!$%s$%d</f><v>0</v>" % (xesc(nm), fc, trow)
            if re.search(r'<c r="%s%d"[^>]*/>' % (sc, r), rowxml):
                rowxml = re.sub(r'<c r="%s%d"([^>]*?)\s*/>' % (sc, r),
                                lambda m: '<c r="%s%d"%s>%s</c>' % (sc, r, m.group(1), f),
                                rowxml, count=1)
            else:
                rowxml = re.sub(r'(<c r="%s%d"[^>]*>).*?</c>' % (sc, r),
                                lambda m: m.group(1) + f + '</c>', rowxml, count=1)
        s0, e0, _ = rows[r]
        sh = sh[:s0] + rowxml + sh[e0:]
        rows = rows_of(sh)
    pkg.wr(summary_part, sh)
    pkg.wr('xl/sharedStrings.xml', ss)


def fix_sst_counts(pkg):
    ss = pkg.rd('xl/sharedStrings.xml')
    uniq = len(re.findall(r'<si>', ss))
    refs = sum(len(re.findall(r't="s"', pkg.rd(s[3]))) for s in sheet_index(pkg))
    pkg.wr('xl/sharedStrings.xml',
           re.sub(r'count="\d+" uniqueCount="\d+"',
                  'count="%d" uniqueCount="%d"' % (refs, uniq), ss, count=1))


# ------------------------------------------------------------------- driver
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--names', nargs='+', required=True)
    ap.add_argument('--rows', type=int, help='grow the summary to this many asset rows')
    ap.add_argument('--summary-sheet')
    ap.add_argument('--template-sheet')
    a = ap.parse_args()

    for nm in a.names:
        if len(nm) > 31 or re.search(r'[:\\/?*\[\]]', nm):
            raise SystemExit('Invalid Excel tab name: %r (max 31 chars, no : \\ / ? * [ ])' % nm)
    if len(set(a.names)) != len(a.names):
        raise SystemExit('Asset names must be unique.')

    pkg = Pkg(a.file)
    sheets = sheet_index(pkg)
    summary, template, linked_row = find_summary_and_template(
        pkg, sheets, a.summary_sheet, a.template_sheet)
    print('summary sheet : %r' % summary[0])
    print('template tab  : %r  (linked from summary row %s)' % (template[0], linked_row))

    assets, total, step = summary_layout(pkg.rd(summary[3]))
    print('summary holds : %d asset rows (rows %d..%d, step %d), total row %d'
          % (len(assets), assets[0], assets[-1], step, total))

    want = max(a.rows or 0, len(a.names))
    if want > len(assets):
        n = extend_summary(pkg, summary[3], want)
        print('extended      : +%d asset rows -> %d' % (n, want))

    new_sheets, followed = clone_tabs(pkg, sheets, template, a.names)
    print('tabs          : %r renamed to %r, %d clone(s) added'
          % (template[0], a.names[0], len(new_sheets)))
    if followed:
        print('rename follow : %d reference(s) on other sheets repointed' % followed)

    link_summary(pkg, summary[3], a.names, linked_row)
    fix_sst_counts(pkg)
    print('linked        : %d asset row(s) wired to their tabs' % len(a.names))

    pkg.save(a.out)
    print('saved         : %s (%d bytes)' % (a.out, os.path.getsize(a.out)))


if __name__ == '__main__':
    main()
