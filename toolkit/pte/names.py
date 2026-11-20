"""names.py - asset / package register for the hub's 06-03 step.

The 06-03 skill splits names like FF-L-CV-AR-NEW (asset = first 3 hyphen parts). Real BOQ names are often
messier, e.g. "Sample-Compound. L-VT -96-AL-1.Elec.xlsx", so the hub:
  1. guesses asset and package per file (guess), with how it got there,
  2. shows the guesses on a review screen where they can be corrected,
  3. builds File_Names.xlsx / .txt from the approved rows (build), with the skill's layout,
     keeping each BOQ's real file name in the File Name column (later steps find the BOQs by it).
"""
import os, re

from openpyxl import load_workbook

import extract_names as E

# A BOQ that covers a whole asset (all disciplines in one BOQ) has the package code "Whole" (Ahmed, 19 Nov 2026):
# it goes in the register's Package column and in column B (Client Code) when step 05 loads it.
WHOLE = 'Whole'


def is_whole(package):
    return re.sub(r' -\d+$', '', (package or '').strip()).lower() == WHOLE.lower()

# first word of a package / sheet name -> discipline label
DISCIPLINES = [('arch', 'Arch'), ('struc', 'Struc'), ('elec', 'Elec'), ('plum', 'Plum'), ('fire', 'Fire'),
               ('ff', 'Fire'), ('hvac', 'HVAC'), ('mech', 'Mech'), ('land', 'Land'), ('infra', 'Infra'),
               ('civil', 'Civil'), ('lv', 'LV'), ('hv', 'HV')]
CODE = re.compile(r'[A-Za-z0-9]+(?:-[A-Za-z0-9]+){3,}')                          # FF-L-CV-AR: the skill's convention
ASSET = re.compile(r'(?<![A-Z])(?:[A-Z]+\s*-\s*)+[\d,]+\s*-\s*[A-Za-z0-9]+\s*-\s*\d+')   # L-VT -96,97-A3-02


def discipline(text):
    w = re.split(r'[\s._\-]+', (text or '').strip().lower())
    w = [x for x in w if x]
    if not w:
        return None
    for key, label in DISCIPLINES:
        if w[0] == key or (len(key) > 2 and w[0].startswith(key)):
            return label
    return None


def package_name(text):
    """'.Plum Third' -> 'Plum Third', 'architecture' -> 'Arch'."""
    t = re.sub(r'\s+', ' ', (text or '').strip(' .-_')).strip()
    if not t:
        return ''
    words = t.split(' ')
    d = discipline(words[0])
    if d:
        rest = ' '.join(x[:1].upper() + x[1:] for x in words[1:])
        return (d + ' ' + rest).strip()
    return t


def asset_code(raw):
    parts = [p for p in re.sub(r'\s+', '', raw).split('-') if p]
    changed = True
    while changed:                       # 'L-VT' and 'L-VT-L-VT' typed twice in the name
        changed = False
        for n in (2, 1):
            for i in range(len(parts) - 2 * n + 1):
                if parts[i:i + n] == parts[i + n:i + 2 * n]:
                    del parts[i + n:i + 2 * n]; changed = True; break
            if changed:
                break
    return '-'.join(parts)


def sheet_names(path):
    try:
        wb = load_workbook(path, read_only=True)
        names = list(wb.sheetnames); wb.close()
        return names
    except Exception:
        return []


def guess(filename, sheets=()):
    """-> (asset, package, how)."""
    stem = E.clean(filename)
    if CODE.fullmatch(stem):
        a, p = E.split(stem)
        return a, p, 'file name code (asset = first 3 parts)'
    s = re.sub(r'^B\(\d+\)\.?\s*', '', stem)
    m = None
    for m in ASSET.finditer(s):
        pass
    if m:
        asset = asset_code(m.group(0))
        pkg = package_name(s[m.end():])
        if pkg:
            return asset, pkg, 'asset code in the name, package after it'
        for sh in sheets:
            if discipline(sh):
                return asset, discipline(sh), f'asset code in the name, package from sheet "{sh.strip()}"'
        return asset, '', 'asset code in the name, no package found'
    segs = [x.strip() for x in s.split('.') if x.strip()]
    if len(segs) >= 2 and discipline(segs[-1]):
        return segs[-2], package_name(segs[-1]), 'last two parts of the name (asset . package)'
    return stem, '', 'no pattern found'


def merge_qualifiers(rows):
    """'Plum Third' -> 'Plum' when it is the asset's only Plum BOQ (Ahmed, 18 Nov 2026: plumbing BOQs named
    Plum / Plum Second / Plum Third / Plum Groud are one package). Only for guessed names, not the skill's
    code convention (AR-NEW next to AR may be a separate BOQ), and only when the asset has one BOQ of that
    discipline, so two real packages are never merged. The review screen can still change it."""
    count = {}
    for r in rows:
        d = discipline(r['package'])
        if d:
            count[(r['asset'].upper(), d)] = count.get((r['asset'].upper(), d), 0) + 1
    for r in rows:
        d = discipline(r['package'])
        if (d and r['package'] != d and ' ' in r['package'] and not r['how'].startswith('file name code')
                and count[(r['asset'].upper(), d)] == 1):
            r['how'] += f'; "{r["package"]}" taken as {d} (the asset\'s only {d} BOQ)'
            r['package'] = d
    return rows


def read_register(path):
    """Rows of a previous File_Names.xlsx: [{file, asset, package}]."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    if E.REGISTER not in wb.sheetnames:
        return []
    rows, col = [], None
    for row in wb[E.REGISTER].iter_rows(values_only=True):
        if col is None:
            if row and 'Asset' in row and 'Package' in row:
                col = {h: i for i, h in enumerate(row) if h}
            continue
        a, p = row[col['Asset']], row[col['Package']]
        if a and p:
            f = row[col['File Name']] if 'File Name' in col else None
            rows.append(dict(file=str(f or f'{a}-{p}').strip(), asset=str(a).strip(), package=str(p).strip()))
    return rows


def problems(rows):
    """Things that must be fixed before the register is built."""
    out, seen = [], {}
    for r in rows:
        if not r['asset'].strip():
            out.append(f"{r['file']}: no asset")
        if not r['package'].strip():
            out.append(f"{r['file']}: no package")
        k = (r['asset'].strip().upper(), r['package'].strip().upper())
        if all(k) and k in seen:
            out.append(f"{r['file']} and {seen[k]} are both {r['asset']} / {r['package']}")
        seen.setdefault(k, r['file'])
    return out


def number_duplicates(rows):
    """Two BOQs of one asset with the same code get numbered: "Arch -1", "Arch -2" (Ahmed, 18 Nov 2026).
    The code goes into column B (Client Code) of every item loaded from that BOQ."""
    groups = {}
    for r in rows:
        p = r['package'].strip()
        if p:
            groups.setdefault((r['asset'].strip().upper(), p.upper()), []).append(r)
    for rs in groups.values():
        if len(rs) > 1:
            base = rs[0]['package'].strip()
            for i, r in enumerate(rs, 1):
                r['package'] = f'{base} -{i}'
                if 'how' in r:
                    r['how'] += f'; numbered: {len(rs)} BOQs of {r["asset"]} share the code {base}'
    return rows


def remarks(rows):
    """Worth telling the QS (from the skill): packages only some assets have, extra words in a package."""
    assets = sorted({r['asset'] for r in rows})
    by_pkg = {}
    for r in rows:                                  # Arch -1 / Arch -2 count as Arch here
        by_pkg.setdefault(re.sub(r' -\d+$', '', r['package']), set()).add(r['asset'])
    out = []
    for a in assets:                                # a whole-asset BOQ next to package BOQs of the same asset
        mine = [r['package'] for r in rows if r['asset'] == a]
        if any(is_whole(p) for p in mine) and not all(is_whole(p) for p in mine):
            out.append(f"{a}: has a whole-asset BOQ and package BOQs ({', '.join(p for p in mine if not is_whole(p))}); "
                       f"check they do not price the same work twice.")
    if len(assets) > 1:
        for p, have in sorted(by_pkg.items()):
            if is_whole(p):                         # some assets priced whole, others by package: expected
                continue
            if len(have) < len(assets):
                lack = [a for a in assets if a not in have]
                out.append(f"{p}: only in {', '.join(sorted(have))}" if len(have) <= len(lack)
                           else f"{p}: missing in {', '.join(lack)}")
    for p in sorted(by_pkg):
        if re.search(r' -\d+$', p):                 # numbered duplicates (Arch -1, Arch -2) are deliberate
            continue
        if ' ' in p or '-' in p:
            base = p.replace('-', ' ').split(' ')[0]
            if base in by_pkg or discipline(base):
                out.append(f'"{p}" next to "{base}": a separate BOQ, or a revision that replaces it? Check before naming assets.')
    return out


def build(rows, out_dir):
    """Write File_Names.xlsx and .txt from approved rows (previous-register rows first). Returns the groups."""
    g = E.Groups()
    names = {}
    for r in rows:
        a, p = r['asset'].strip(), r['package'].strip()
        g.add(a, p)
        names[(a.upper(), p.upper())] = r['file']
    os.makedirs(out_dir, exist_ok=True)
    E.write_xlsx(os.path.join(out_dir, 'File_Names.xlsx'), g, names)
    lines = []
    for a in g.order:
        lines.append(a)
        lines += g.codes[a]
    with open(os.path.join(out_dir, 'File_Names.txt'), 'w', encoding='utf-8', newline='\r\n') as fh:
        fh.write('\n'.join(lines))
    return g
