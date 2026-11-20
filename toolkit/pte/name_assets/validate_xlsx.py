#!/usr/bin/env python3
"""
Check the things Excel refuses to open a file over. LibreOffice is far more
forgiving than Excel, so "it converts" is not evidence the file is sound.
"""
import re, sys, zipfile, xml.etree.ElementTree as ET
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PR = '{http://schemas.openxmlformats.org/package/2006/relationships}'

path = sys.argv[1]
z = zipfile.ZipFile(path)
parts = set(z.namelist())
ok = True


def fail(msg):
    global ok
    ok = False
    print('  FAIL', msg)


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


print('1. every part parses')
for n in sorted(parts):
    if n.endswith(('.xml', '.rels')):
        try:
            ET.fromstring(z.read(n))
        except Exception as e:
            fail('%s: %s' % (n, e))
print('   %d parts' % len(parts))

print('2. content types and relationships resolve')
ct = z.read('[Content_Types].xml').decode()
declared = set(re.findall(r'PartName="/([^"]+)"', ct))
for d in declared:
    if d not in parts:
        fail('Content_Types declares a missing part: %s' % d)
for n in sorted(parts):
    if n.endswith('.xml') and n != '[Content_Types].xml' and n not in declared:
        if not re.search(r'Extension="%s"' % n.rsplit('.', 1)[1], ct):
            fail('part not declared: %s' % n)
for rels in [n for n in parts if n.endswith('.rels')]:
    base = rels.rsplit('_rels/', 1)[0].rstrip('/')
    for r in ET.fromstring(z.read(rels)).findall(PR + 'Relationship'):
        t = r.get('Target')
        if r.get('TargetMode') == 'External' or t.startswith('http'):
            continue
        p = t[1:] if t.startswith('/') else (base + '/' + t if base else t)
        while '/../' in p:
            p = re.sub(r'[^/]+/\.\./', '', p, count=1)
        if p not in parts:
            fail('%s points at a missing target: %s' % (rels, t))

print('3. sheets in workbook.xml all resolve to a part')
wb = ET.fromstring(z.read('xl/workbook.xml'))
rel = {r.get('Id'): r.get('Target') for r in
       ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).findall(PR + 'Relationship')}
sheets = []
for s in wb.iter(NS + 'sheet'):
    rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    tgt = rel.get(rid)
    if not tgt or ('xl/' + tgt) not in parts:
        fail('sheet %r has no part' % s.get('name'))
    sheets.append(('xl/' + tgt, s.get('name')))
names = [n for _, n in sheets]
if len(names) != len(set(names)):
    fail('duplicate sheet names')
ids = [s.get('sheetId') for s in wb.iter(NS + 'sheet')]
if len(ids) != len(set(ids)):
    fail('duplicate sheetIds')
print('   %d sheets: %s' % (len(names), ', '.join(repr(n) for n in names)))

print('4. shared string table')
ss = z.read('xl/sharedStrings.xml').decode()
uniq = len(re.findall(r'<si[ >]', ss))
declared_u = int(re.search(r'uniqueCount="(\d+)"', ss).group(1))
declared_c = int(re.search(r'count="(\d+)"', ss).group(1))
refs = 0
hi = -1
for part, _ in sheets:
    d = z.read(part).decode()
    refs += len(re.findall(r't="s"', d))
    for v in re.findall(r't="s"[^>]*>\s*<v>(\d+)</v>', d):
        hi = max(hi, int(v))
if declared_u != uniq:
    fail('uniqueCount=%d but %d <si> elements' % (declared_u, uniq))
if declared_c != refs:
    fail('count=%d but %d references' % (declared_c, refs))
if hi >= uniq:
    fail('a cell points at index %d, table holds %d' % (hi, uniq))
print('   uniqueCount=%d count=%d highest index used=%d' % (uniq, refs, hi))

print('5. cells: row agreement, ordering, duplicates, spans')
for part, name in sheets:
    root = ET.fromstring(z.read(part))
    for row in root.iter(NS + 'row'):
        r = row.get('r')
        cols = []
        for c in row.findall(NS + 'c'):
            m = re.match(r'([A-Z]+)(\d+)$', c.get('r') or '')
            if not m or m.group(2) != r:
                fail('%s row %s holds cell %s' % (name, r, c.get('r')))
            else:
                cols.append(colnum(m.group(1)))
            if c.get('t') == 's' and c.find(NS + 'f') is not None:
                fail('%s %s is both a shared string and a formula' % (name, c.get('r')))
        if cols != sorted(cols):
            fail('%s row %s cells out of order' % (name, r))
        if len(cols) != len(set(cols)):
            fail('%s row %s duplicate cell refs' % (name, r))
        sp = row.get('spans')
        if sp and cols:
            a, b = (int(x) for x in sp.split(':'))
            if a > min(cols) or b < max(cols):
                fail('%s row %s spans=%s but cells run %d-%d' % (name, r, sp, min(cols), max(cols)))

print('6. shared formulas')
for part, name in sheets:
    d = z.read(part).decode()
    masters = re.findall(r'<f t="shared"[^>]*\bref="([^"]+)"[^>]*\bsi="(\d+)"', d)
    seen = {}
    for ref, si in masters:
        if si in seen:
            fail('%s: si=%s declared master twice (%s and %s)' % (name, si, seen[si], ref))
        seen[si] = ref
    used = set(re.findall(r'<f t="shared"[^>]*\bsi="(\d+)"', d))
    for si in used - set(seen):
        fail('%s: si=%s used with no master' % (name, si))
    # every follower must sit inside its master's ref range
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*>\s*<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"', d):
        c0, r0, si = colnum(m.group(1)), int(m.group(2)), m.group(3)
        ref = seen.get(si)
        if not ref:
            continue
        a, b = (ref.split(':') + [ref])[:2]
        ma, mb = re.match(r'([A-Z]+)(\d+)', a), re.match(r'([A-Z]+)(\d+)', b)
        if not (colnum(ma.group(1)) <= c0 <= colnum(mb.group(1))
                and int(ma.group(2)) <= r0 <= int(mb.group(2))):
            fail('%s: %s%s uses si=%s whose master range is %s'
                 % (name, m.group(1), m.group(2), si, ref))

print('7. merges and dimensions')
for part, name in sheets:
    d = z.read(part).decode()
    m = re.search(r'<mergeCells count="(\d+)"', d)
    if m:
        n = len(re.findall(r'<mergeCell ', d))
        if int(m.group(1)) != n:
            fail('%s: mergeCells count=%s but %d children' % (name, m.group(1), n))
    # overlapping merges are invalid and Excel repairs the file over them
    boxes = []
    for mm in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', d):
        c1, r1, c2, r2 = (colnum(mm.group(1)), int(mm.group(2)),
                          colnum(mm.group(3)), int(mm.group(4)))
        boxes.append((min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2), mm.group(0)))
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a1, b1, a2, b2, t1 = boxes[i]
            c1, d1, c2, d2, t2 = boxes[j]
            if a1 <= c2 and c1 <= a2 and b1 <= d2 and d1 <= b2:
                fail('%s: overlapping merges %s and %s' % (name, t1, t2))
                break
        else:
            continue
        break
    root = ET.fromstring(z.read(part))
    dim = root.find(NS + 'dimension')
    if dim is not None and ':' in (dim.get('ref') or ''):
        end = re.match(r'([A-Z]+)(\d+)', dim.get('ref').split(':')[1])
        mc = mr = 0
        for row in root.iter(NS + 'row'):
            for c in row.findall(NS + 'c'):
                g = re.match(r'([A-Z]+)(\d+)$', c.get('r') or '')
                if g:
                    mc = max(mc, colnum(g.group(1)))
                    mr = max(mr, int(g.group(2)))
        if colnum(end.group(1)) < mc or int(end.group(2)) < mr:
            fail('%s: dimension %s smaller than the used range' % (name, dim.get('ref')))

print('\nRESULT:', 'PASS' if ok else 'FAIL')
sys.exit(0 if ok else 1)
