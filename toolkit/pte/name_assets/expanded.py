"""expanded.py - name the assets in the AAA PTE master that ships ready-made asset tabs (hub, 06-04).

The 06-04 skill handles two template families: "packaged" (Asset #01 {no package} + Asset #02 {with package})
and "flat" (one template tab). Some masters, e.g. "0000.00.00 - CLIENT - PROJECT - PTE & Tender Evaluation
- R01.xlsx", are a third, "expanded" layout: 25 identical tabs Asset #01..#25 already linked from Schedule of
Assets, Asset Summary, Project Summary and Elemental Summary (NRM01), each a single block of divisions with
its own Total Amounts row, and an extra "Client Code" column B (so AAA PTE is column L, not K).

Everything is done by Excel (toolkit/pte/excel_ops.ps1), not by editing the XML or saving through openpyxl:
the file carries web extensions, images and data-validation extensions openpyxl would drop, and Excel itself
repoints every reference when a tab is renamed or rows/columns move. openpyxl only READS the file here.

  detect(path)                    -> layout (rows, columns, links), defects found, problems that stop the build
  plan(layout, assets, ...)       -> list of Excel operations + what the result must look like
  verify_static(out, expect)      -> every link and formula checked in the saved file
  verify_numbers(out, expect)     -> amounts planted in a throwaway copy, recalculated by Excel, read back

assets: [{"name": "VT", "packages": ["Arch", "Elec", "Plum"]}, ...] in tab order; [] packages = one block.
"""
import re, sys, warnings
from pathlib import Path

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import excel  # noqa: E402

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
SCHEDULE_NAMES = ('Schedule of Assets', 'Scheduale of Assets')
LINK = re.compile(r"^='?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$")
ABS_ROW = re.compile(r"(\$?[A-Z]{1,3})\$(\d+)\b")
BAD_NAME = re.compile(r"[\[\]:*?/\\']")


class LayoutError(Exception):
    pass


def _find(ws, col, text, rows=range(1, 400)):
    for r in rows:
        v = ws[f'{col}{r}'].value
        if isinstance(v, str) and v.strip().lower().startswith(text.lower()):
            return r
    return None


def _header_cols(ws, row):
    out = {}
    for c in ws[row]:
        if isinstance(c.value, str):
            out[re.sub(r'\s+', ' ', c.value).strip()] = c.column_letter
    return out


def detect(path):
    wb = openpyxl.load_workbook(path)
    names = wb.sheetnames
    if any('{with package}' in n for n in names):
        raise LayoutError('This is the packaged R00 template ("Asset #02 {with package}"). The hub builds the master '
                          'that ships ready-made Asset #01..#25 tabs; use the 06-04 skill in Claude for this one.')
    sched = next((n for n in SCHEDULE_NAMES if n in names), None)
    elem = next((n for n in names if n.startswith('Elemental Summary')), None)
    missing = [n for n, ok in (('Schedule of Assets', sched), ('Asset Summary', 'Asset Summary' in names),
                               ('Project Summary', 'Project Summary' in names), ('Elemental Summary', elem)) if not ok]
    if missing:
        raise LayoutError('Not the expected PTE master: no ' + ', '.join(missing) + ' sheet.')
    L = dict(path=str(path), sheets=names, schedule=sched, elemental=elem)

    # ---- Asset Summary: asset rows, the tab each links to, and which tab column feeds which summary column
    # Asset rows are the numbered rows under the "Sr." header. In the delivered master only assets 1-2 are
    # linked to their tabs; for the others the tab is the name the row shows (from the Schedule).
    s = wb[sched]
    a = wb['Asset Summary']
    a_total = _find(a, 'C', 'Total Amounts')
    a_hdr = _find(a, 'B', 'Sr.', range(1, 30))
    rows, link_map = [], {}
    for r in range(a_hdr + 1, a_total):
        sr = a[f'B{r}'].value
        if not isinstance(sr, (int, float)):
            continue
        tab, cols = None, {}
        for c in a[r]:
            m = LINK.match(str(c.value or ''))
            if m and m.group(1) != sched:
                tab = m.group(1); cols[c.column_letter] = (m.group(2), int(m.group(3)))
        m = LINK.match(str(a[f'C{r}'].value or ''))
        sched_row = int(m.group(3)) if m else None
        linked = bool(tab)
        if not tab and sched_row:
            tab = s[f'E{sched_row}'].value
        if tab not in names:
            raise LayoutError(f'Asset Summary row {r}: cannot tell which asset tab it belongs to.')
        rows.append(dict(row=r, sr=sr, tab=tab, sched_row=sched_row, linked=linked))
        if cols and not link_map:
            link_map = {k: v[0] for k, v in cols.items()}
    if not rows or not link_map:
        raise LayoutError('Asset Summary has no rows linked to asset tabs.')
    tabs = [x['tab'] for x in rows]
    L['asum'] = dict(total=a_total, rows=rows, link_map=link_map)

    p = wb['Project Summary']
    p_total = _find(p, 'C', 'Total Amounts')
    p_hdr = _find(p, 'B', 'Sr.', range(1, 30))
    prow = [r for r in range(p_hdr + 1, p_total) if isinstance(p[f'B{r}'].value, (int, float))]
    if len(prow) != len(rows):
        raise LayoutError(f'Project Summary has {len(prow)} asset rows, Asset Summary {len(rows)}.')
    L['psum'] = dict(total=p_total, rows=prow)

    L['sched_rows'] = [x['sched_row'] for x in rows]
    hdr = _find(s, 'E', 'Asset Name', range(1, 30)) or 6
    s_cols = [c.column_letter for c in s[hdr] if c.value not in (None, '') and c.column > CI('E')]
    L['sched_cols'] = s_cols
    L['sched_samples'] = [f'{c}{r}' for r in L['sched_rows'] for c in s_cols
                          if isinstance(s[f'{c}{r}'].value, (int, float))]

    # ---- Elemental Summary: one group per asset, found by the header link to the Schedule row
    e = wb[elem]
    groups = []
    for c in e[2]:
        m = LINK.match(str(c.value or ''))
        if m and m.group(1) == sched:
            groups.append(dict(col=c.column_letter, sched_row=int(m.group(3))))
    by_sched = {g['sched_row']: g for g in groups}
    first_f = last_f = None
    for g in groups:
        sumif = []
        for r in range(3, e.max_row + 1):
            v = e[f"{g['col']}{r}"].value
            if isinstance(v, str) and v.startswith('='):
                first_f = r if first_f is None else min(first_f, r)
                last_f = max(last_f or r, r)
                if 'SUMIF(' in v:
                    sumif.append(r)
        g['sumif_rows'] = sumif
    L['elem'] = dict(groups=groups, first=first_f, last=last_f, width=3,
                     total=_find(e, 'C', 'Total', range(5, e.max_row + 1)))
    for x in rows:
        if x['sched_row'] not in by_sched:
            raise LayoutError(f"Elemental Summary has no column group for {x['tab']}.")

    # ---- the asset tabs (all must be laid out like the first and still unpriced)
    t = wb[tabs[0]]
    hdr_row = _find(t, 'A', 'NRM Code', range(1, 20))
    h = _header_cols(t, hdr_row)
    sub = _header_cols(t, hdr_row + 1)
    col = dict(nrm=h.get('NRM Code', 'A'), wa=h.get('Weight/ Asset'), wd=h.get('Weight/ Division'), sr=h.get('Sr.'),
               desc=h.get('BOQ Item Description'))
    aaa = h.get('AAA')
    if not all(col.values()) or not aaa:
        raise LayoutError(f'{tabs[0]}: header row {hdr_row} is not the AAA asset-tab layout.')
    a0 = CI(aaa)
    col.update(qty_aaa=CL(a0), rate=CL(a0 + 1), est=CL(a0 + 2), pte=CL(a0 + 3), qty_des=CL(a0 - 1))
    if not str(t[f"{col['pte']}{hdr_row + 1}"].value or '').startswith('Pre-Tender'):
        raise LayoutError(f"{tabs[0]}: column {col['pte']} is not 'Pre-Tender Estimate Amount'.")
    total = _find(t, col['desc'], 'Total Amounts', range(hdr_row, t.max_row + 1))
    divs = [r for r in range(hdr_row + 2, total)
            if str(t[f"{col['desc']}{r}"].value or '').startswith('Division')]
    sum_cols = [c.column_letter for c in t[total] if str(c.value or '').upper().startswith('=SUBTOTAL(')]
    m = re.search(r'\(9,[A-Z]+(\d+):', str(t[f"{col['pte']}{total}"].value))
    body_start = int(m.group(1)) if m else hdr_row + 2
    tmpl = {}
    for r in range(divs[0], total):
        tmpl[r] = {c: t[f'{c}{r}'].value for c in [col['wa'], col['wd'], col['sr']]}
        tmpl[r]['_level'] = t.row_dimensions[r].outline_level or 0
    L['tab'] = dict(header=hdr_row, first_div=divs[0], total=total, divs=divs, sum_cols=sum_cols, cols=col,
                    body_start=body_start, template=tmpl,
                    print_last_col=re.sub(r'.*\$([A-Z]+)\$\d+$', r'\1', str(t.print_area or '')) or CL(t.max_column),
                    filter=t.auto_filter.ref)
    L['tabs'] = tabs

    def sig(ws):
        return [(c.coordinate, c.value) for row in ws.iter_rows(min_row=1, max_row=total + 1, max_col=CI(sum_cols[-1]))
                for c in row if c.value is not None]
    s0 = sig(t)
    problems = []
    for n in tabs[1:]:
        if sig(wb[n]) != s0:
            problems.append(f'{n} is not laid out like {tabs[0]} (already edited or loaded?)')
    for n in tabs:
        ws = wb[n]
        priced = [f'{c}{r}' for r in range(divs[0], total) for c in (col['qty_des'], col['qty_aaa'], col['rate'])
                  if isinstance(ws[f'{c}{r}'].value, (int, float))]
        if priced:
            problems.append(f'{n} already has quantities or rates ({", ".join(priced[:3])}); this step is for the blank master')
    L['problems'] = problems

    # ---- defects worth reporting (most are fixed by the build)
    d = []
    unlinked = [x['sr'] for x in rows if not x['linked']]
    if unlinked:
        d.append(f'Asset Summary: assets {unlinked[0]}-{unlinked[-1]} are not linked to their tabs, so their amounts '
                 'would never reach the summary or the project total (links written for every asset kept).')
    mism = 0
    for g in groups:
        for r in g['sumif_rows']:
            f = e[f"{g['col']}{r}"].value
            q = re.findall(r"SUMIF\('([^']+)'!", f) + re.findall(r",'([^']+)'!\$[A-Z]+:\$[A-Z]+\)", f)
            if len(set(q)) > 1:
                mism += 1
    if mism:
        d.append(f'Elemental Summary: {mism} SUMIF cells look up the NRM code in one tab but add up another tab '
                 '(rewritten so each group reads only its own asset).')
    for sh, info in (('Asset Summary', L['asum']), ('Project Summary', L['psum'])):
        ws = wb[sh]
        f = str(ws[f"A{info['total']}"].value or '')
        m = re.search(r'A(\d+):A(\d+)', f)
        if m and int(m.group(2)) < info['total'] - 1:
            d.append(f'{sh}: the weight total {f} stops at row {m.group(2)}, before the last assets.')
        srs = [ws[f'B{r}'].value for r in (info['rows'] if sh == 'Project Summary' else [x['row'] for x in rows])]
        bad = [f'{i + 1} is numbered {v}' for i, v in enumerate(srs) if v != i + 1]
        if bad:
            d.append(f'{sh}: Sr. ' + ', '.join(bad) + '.')
    letters = [t[f"{col['sr']}{r}"].value for r in divs]
    dup = sorted({x for x in letters if letters.count(x) > 1})
    if dup:
        d.append(f'Asset tabs: division letter {", ".join(dup)} is used twice (packages renumber divisions A.01, A.02...).')
    nowt = [r for r in range(divs[0], total) if r not in divs and t[f"{col['est']}{r}"].value
            and not t[f"{col['wa']}{r}"].value]
    if nowt:
        d.append(f'Asset tabs: item row {", ".join(map(str, nowt))} has amount formulas but no weight formulas '
                 '(left as the template has it).')
    if L['sched_samples']:
        d.append(f"{sched}: sample areas in {len(L['sched_samples'])} cells (10, 20, 30...) would show as real BUA "
                 'and cost/m2; cleared for the named assets.')
    L['defects'] = d
    L['ref_baseline'] = ref_counts(wb)
    return L


def ref_counts(wb):
    """#REF! formulas per sheet (the template already has some; only new ones count against a build)."""
    out = {}
    for ws in wb.worksheets:
        n = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and '#REF!' in c.value)
        if n:
            out[ws.title] = n
    return out


def _n(v):
    """Compare sheet references with or without quotes (Excel drops them for simple names like VT)."""
    return str(v).replace("'", '')


def check_names(assets, existing):
    """Problems with the asset names / packages chosen on the review screen."""
    out, seen = [], set()
    for x in assets:
        n = x['name']
        if not n:
            out.append('An asset has no name.')
        elif len(n) > 31:
            out.append(f'"{n}" is over 31 characters (Excel tab name limit).')
        elif BAD_NAME.search(n):
            out.append(f'"{n}" has a character not allowed in a tab name ([ ] : * ? / \\ or an apostrophe).')
        elif n.lower() in seen:
            out.append(f'"{n}" is used twice.')
        elif n.lower() in {e.lower() for e in existing}:
            out.append(f'"{n}" is already the name of another sheet.')
        seen.add(n.lower())
        if len(set(p.lower() for p in x['packages'])) != len(x['packages']):
            out.append(f'{n}: a package is listed twice.')
    return out


def _move(f, r_from, r_to, rows_map):
    """Template formula from row r_from, rewritten for row r_to: relative refs follow, absolute rows are mapped."""
    if not (isinstance(f, str) and f.startswith('=')):
        return f
    f = Translator(f, origin=f'A{r_from}').translate_formula(f'A{r_to}')
    return ABS_ROW.sub(lambda m: f'{m.group(1)}${rows_map.get(int(m.group(2)), int(m.group(2)))}', f)


def plan(L, assets, delete_spare=True, amount='pte'):
    """Excel operations to name the assets; returns (ops, expect)."""
    n = len(assets)
    tabs = L['tabs']
    if n > len(tabs):
        raise LayoutError(f'{n} assets, but the master has only {len(tabs)} asset tabs.')
    T = L['tab']; C = T['cols']
    amt = C['pte'] if amount == 'pte' else C['est']
    ops, expect = [], dict(assets=[], sched=L['schedule'], elem=L['elemental'], delete_spare=delete_spare, amount_col=amt,
                           link_map=L['asum']['link_map'], cols=C, sum_cols=T['sum_cols'], body_start=T['body_start'])
    E = L['elem']; sched = L['schedule']

    # 1. Elemental: each kept group reads only its own tab, from the chosen amount column
    for k in range(n):
        g = next(g for g in E['groups'] if g['sched_row'] == L['asum']['rows'][k]['sched_row'])
        tab = tabs[k]
        wb_vals = _elem_column(L, g['col'])
        for r in g['sumif_rows']:
            wb_vals[r - E['first']] = (f'=IF($B{r}="",0,SUMIF(\'{tab}\'!${C["nrm"]}:${C["nrm"]},$B{r},'
                                       f'\'{tab}\'!${amt}:${amt}))')
        wb_vals = [v if isinstance(v, (str, int, float)) else str(getattr(v, 'text', v)) for v in wb_vals]
        ops.append(dict(op='column', sheet=L['elemental'], col=g['col'], **{'from': E['first']}, values=wb_vals))
    spare_tabs = tabs[n:] if delete_spare else []
    if spare_tabs:
        g0 = next(g for g in E['groups'] if g['sched_row'] == L['asum']['rows'][n]['sched_row'])
        last = max(CI(g['col']) for g in E['groups']) + E['width'] - 1
        ops.append(dict(op='delete_cols', sheet=L['elemental'], cols=f"{g0['col']}:{CL(last)}"))
        # 2. summaries and schedule: drop the spare rows (Excel shrinks every range and moves the totals)
        a_rows = L['asum']['rows']
        ops.append(dict(op='delete_rows', sheet='Asset Summary', rows=f"{a_rows[n]['row']}:{L['asum']['total'] - 1}"))
        ops.append(dict(op='delete_rows', sheet='Project Summary', rows=f"{L['psum']['rows'][n]}:{L['psum']['total'] - 1}"))
        ops.append(dict(op='delete_rows', sheet=sched, rows=f"{L['sched_rows'][n]}:{L['sched_rows'][-1]}"))
        for t in spare_tabs:
            ops.append(dict(op='delete_sheet', sheet=t))

    # 2b. every kept summary row linked to its own tab's Total Amounts row (written before the tabs grow, so
    #     Excel moves these references down with the total row) and numbered 1, 2, 3...
    fd, T0 = T['first_div'], T['total']
    kept = len(tabs) if not delete_spare else n
    for k in range(kept):
        ar, tab = L['asum']['rows'][k]['row'], tabs[k]
        for scol, tcol in L['asum']['link_map'].items():
            ops.append(dict(op='set', sheet='Asset Summary', cell=f'{scol}{ar}', f=f"='{tab}'!${tcol}${T0}"))
        ops.append(dict(op='set', sheet='Asset Summary', cell=f'B{ar}', f=k + 1))
        ops.append(dict(op='set', sheet='Project Summary', cell=f"B{L['psum']['rows'][k]}", f=k + 1))
    # the weight (A) total of both summaries stopped at row 47 in the template: make it span every row
    for sh, info, rws in (('Asset Summary', L['asum'], [x['row'] for x in L['asum']['rows']]),
                          ('Project Summary', L['psum'], L['psum']['rows'])):
        new_total = rws[n] if spare_tabs else info['total']
        ops.append(dict(op='set', sheet=sh, cell=f'A{new_total}', f=f'=SUBTOTAL(9,A{rws[0] - 2}:A{new_total - 1})'))
    expect['asum_total'] = L['asum']['rows'][n]['row'] if spare_tabs else L['asum']['total']

    # 3. asset tabs: a block per package, weights and totals rewritten explicitly
    n0 = T0 - fd                      # template rows from the first division to the row above Total Amounts
    for k, x in enumerate(assets):
        tab, pk = tabs[k], x['packages']
        info = dict(name=x['name'], tab=tab, packages=pk, headers=[])
        if pk:
            size = n0 + 1
            ops.append(dict(op='insert_rows', sheet=tab, at=fd, count=1))
            ops.append(dict(op='copy_rows', sheet=tab, src=f'{fd + 1}:{fd + 1}', dst=f'{fd}:{fd}'))
            for i in range(1, len(pk)):
                at = fd + size * i
                ops.append(dict(op='insert_rows', sheet=tab, at=at, count=size))
                ops.append(dict(op='copy_rows', sheet=tab, src=f'{fd}:{fd + size - 1}', dst=f'{at}:{at + size - 1}'))
            total = fd + size * len(pk)
            cols = {C['wa']: [], C['wd']: [], C['sr']: []}
            levels = []
            for i, name in enumerate(pk):
                h = fd + size * i
                letter = chr(ord('A') + i)
                info['headers'].append(dict(row=h, letter=letter, name=name, first=h + 1, last=h + n0))
                rows_map = {T0: total}
                rows_map.update({r: h + 1 + (r - fd) for r in range(fd, T0)})
                for c in (C['wa'], C['wd']):
                    cols[c].append(f'=SUBTOTAL(9,{c}{h + 1}:{c}{h + n0})')
                cols[C['sr']].append(letter)
                levels.append((h, 1))
                for off in range(n0):
                    src, dst = fd + off, h + 1 + off
                    tm = T['template'][src]
                    for c in (C['wa'], C['wd']):
                        v = tm[c]
                        if dst not in [h + 1 + (d - fd) for d in T['divs']] and isinstance(v, str) and 'SUBTOTAL' not in v \
                                and src >= T['divs'][-1] + 4:
                            v = None                                  # stray weight formulas on the trailing blank row
                        cols[c].append(_move(v, src, dst, rows_map) or '')
                    if src in T['divs']:
                        cols[C['sr']].append(f'{letter}.{T["divs"].index(src) + 1:02d}')
                    else:
                        cols[C['sr']].append(tm[C['sr']] or '')
                    levels.append((dst, tm['_level'] + 2))
                for c in T['sum_cols']:
                    if c not in (C['wa'], C['wd']):
                        ops.append(dict(op='set', sheet=tab, cell=f'{c}{h}', f=f'=SUBTOTAL(9,{c}{h + 1}:{c}{h + n0})'))
                ops.append(dict(op='set', sheet=tab, cell=f"{C['desc']}{h}", f=f"'Package {letter} : {name}"))
            for c, vals in cols.items():
                ops.append(dict(op='column', sheet=tab, col=c, **{'from': fd}, values=vals))
            for c in T['sum_cols']:
                ops.append(dict(op='set', sheet=tab, cell=f'{c}{total}', f=f"=SUBTOTAL(9,{c}{T['body_start']}:{c}{total - 1})"))
            # outline: package header 1, division headers 2, items 3 (runs of equal level in one call)
            run_start, run_lvl = None, None
            for r, lv in levels + [(None, None)]:
                if run_lvl is not None and (lv != run_lvl or r != prev + 1):
                    ops.append(dict(op='outline', sheet=tab, rows=f'{run_start}:{prev}', level=run_lvl))
                    run_start = None
                if r is None:
                    break
                if run_start is None:
                    run_start, run_lvl = r, lv
                prev = r
            ops.append(dict(op='print_area', sheet=tab, range=f"$A$1:${T['print_last_col']}${total}"))
            if T['filter']:
                a1, b1 = T['filter'].split(':')
                last_col = b1.rstrip('0123456789')
                ops.append(dict(op='autofilter', sheet=tab, range=f'{a1}:{last_col}{total - 1}'))
        else:
            total = T0
        info['total'] = total
        expect['assets'].append(info)

    # 4. names: Schedule (feeds Asset Summary and the elemental headers), Project Summary label, tab names
    for k, x in enumerate(assets):
        sr = L['sched_rows'][k]
        ops.append(dict(op='set', sheet=sched, cell=f'E{sr}', f="'" + x['name']))     # ' = always text
        prow = L['psum']['rows'][k]
        ops.append(dict(op='set', sheet='Project Summary', cell=f'C{prow}', f=f"='{sched}'!$E${sr}"))
    for cell in L['sched_samples']:
        if int(re.sub(r'^[A-Z]+', '', cell)) in L['sched_rows'][:n]:
            ops.append(dict(op='clear', sheet=sched, range=cell))
    for k, x in enumerate(assets):
        if x['name'] != tabs[k]:
            ops.append(dict(op='rename', sheet=tabs[k], to=x['name']))
    expect['spare'] = tabs[n:] if not delete_spare else []
    expect['first_div'] = fd
    expect['ref_baseline'] = L['ref_baseline']
    expect['deleted'] = spare_tabs
    expect['asum_rows'] = [r['row'] for r in L['asum']['rows'][:n]]
    expect['psum_rows'] = L['psum']['rows'][:n]
    expect['sched_rows'] = L['sched_rows'][:n]
    expect['elem_cols'] = [next(g['col'] for g in E['groups'] if g['sched_row'] == L['sched_rows'][k]) for k in range(n)]
    expect['elem_sumif_rows'] = E['groups'][0]['sumif_rows']
    return ops, expect


_elem_cache = {}


def _elem_column(L, col):
    key = (L['path'], col)
    if key not in _elem_cache:
        ws = openpyxl.load_workbook(L['path'])[L['elemental']]
        E = L['elem']
        for c in [g['col'] for g in E['groups']]:
            _elem_cache[(L['path'], c)] = [('' if ws[f'{c}{r}'].value is None else ws[f'{c}{r}'].value)
                                           for r in range(E['first'], E['last'] + 1)]
    return list(_elem_cache[key])


# ---------------- checks on the saved file ----------------
def verify_static(out, expect):
    """Every link and formula the build is responsible for, read back from the saved file."""
    wb = openpyxl.load_workbook(out)
    C = expect['cols']; fails = []
    names = [a['name'] for a in expect['assets']]
    for t in expect['deleted']:
        if t in wb.sheetnames:
            fails.append(f'{t} still there')
    for a in expect['assets']:
        if a['name'] not in wb.sheetnames:
            fails.append(f"tab {a['name']} missing"); continue
        ws = wb[a['name']]; T = a['total']
        if ws[f"{C['desc']}{T}"].value != 'Total Amounts':
            fails.append(f"{a['name']}: Total Amounts is not on row {T}")
        for c in expect['sum_cols']:
            want = f"=SUBTOTAL(9,{c}{expect['body_start']}:{c}{T - 1})"
            if ws[f'{c}{T}'].value != want:
                fails.append(f"{a['name']}!{c}{T} = {ws[f'{c}{T}'].value}, expected {want}")
        for h in a['headers']:
            if ws[f"{C['sr']}{h['row']}"].value != h['letter'] or h['name'] not in str(ws[f"{C['desc']}{h['row']}"].value):
                fails.append(f"{a['name']}: package {h['letter']} header wrong on row {h['row']}")
            for c in expect['sum_cols']:
                want = f"=SUBTOTAL(9,{c}{h['first']}:{c}{h['last']})"
                if ws[f"{c}{h['row']}"].value != want:
                    fails.append(f"{a['name']}!{c}{h['row']} = {ws[f'{c}{h['row']}'].value}, expected {want}")
        div = None
        for r in range(expect['body_start'], T):
            if str(ws[f"{C['desc']}{r}"].value or '').startswith('Division'):
                div = r
            fa, fdv = ws[f"{C['wa']}{r}"].value, ws[f"{C['wd']}{r}"].value
            if isinstance(fa, str) and 'IFERROR' in fa and f"/${C['pte']}${T})" not in fa:
                fails.append(f"{a['name']}!{C['wa']}{r} weight not over the total row {T}: {fa}")
            if isinstance(fdv, str) and 'IFERROR' in fdv and f"/{C['pte']}${div})" not in fdv:
                fails.append(f"{a['name']}!{C['wd']}{r} weight not over its division row {div}: {fdv}")
    s = wb['Asset Summary']
    for k, (a, r) in enumerate(zip(expect['assets'], expect['asum_rows'])):
        for scol, tcol in expect['link_map'].items():
            want = f"='{a['name']}'!${tcol}${a['total']}"
            if _n(s[f'{scol}{r}'].value) != _n(want):
                fails.append(f"Asset Summary {scol}{r} = {s[f'{scol}{r}'].value}, expected {want}")
        if _n(s[f'C{r}'].value) != _n(f"='{expect['sched']}'!$E${expect['sched_rows'][k]}"):
            fails.append(f'Asset Summary C{r} does not link to the Schedule name')
    p = wb['Project Summary']
    for k, r in enumerate(expect['psum_rows']):
        if _n(p[f'C{r}'].value) != _n(f"='{expect['sched']}'!$E${expect['sched_rows'][k]}"):
            fails.append(f'Project Summary C{r} does not link to the Schedule name')
    sc = wb[expect['sched']]
    for k, a in enumerate(expect['assets']):
        if sc[f"E{expect['sched_rows'][k]}"].value != a['name']:
            fails.append(f"Schedule E{expect['sched_rows'][k]} is not {a['name']}")
    e = wb[expect['elem']]
    for k, (a, col) in enumerate(zip(expect['assets'], expect['elem_cols'])):
        if _n(e[f'{col}2'].value) != _n(f"='{expect['sched']}'!$E${expect['sched_rows'][k]}"):
            fails.append(f'Elemental {col}2 does not show the Schedule name of {a["name"]}')
        q = _n(f"'{a['name']}'!")
        for r in expect['elem_sumif_rows']:
            f = _n(e[f'{col}{r}'].value or '')
            if f.count(q) != 2 or f"${expect['amount_col']}:${expect['amount_col']}" not in f:
                fails.append(f'Elemental {col}{r} does not read only {a["name"]} column {expect["amount_col"]}: {f}'); break
    gone = [re.escape(t) for t in expect['deleted']]
    pat = re.compile('|'.join(f"'{g}'!" for g in gone)) if gone else None
    for sh, cnt in ref_counts(wb).items():
        base = expect['ref_baseline'].get(sh, 0)
        if cnt > base:
            fails.append(f'{sh}: {cnt - base} new #REF! formula(s)')
    if pat:
        for ws in wb.worksheets:
            hits = [c.coordinate for row in ws.iter_rows() for c in row if isinstance(c.value, str) and pat.search(c.value)]
            if hits:
                fails.append(f'{ws.title}: {len(hits)} formula(s) still point at deleted tabs, e.g. {hits[0]}')
    return fails


def verify_numbers(out, expect, code=None):
    """Plant amounts in a throwaway copy (never saved), let Excel recalculate, read every total back.

    Asset k gets rate 1000*k (AAA qty 1, Designer qty 2) in its first package and 7*k in its last package,
    under the same NRM code, so the elemental group must show 1007*k: their SUM, from its OWN tab."""
    C = expect['cols']; ops, want = [], {}
    code = code or '0.1.1'
    for k, a in enumerate(expect['assets'], 1):
        hs = a['headers']
        first = hs[0]['first'] if hs else expect['first_div']
        r1 = first + 1                                     # first item row of the first division
        r2 = (hs[-1]['first'] if hs else first) + 5        # first item row of the second division, last package
        ops += [dict(op='set', sheet=a['name'], cell=f"{C['nrm']}{r1}", f="'" + code),
                dict(op='set', sheet=a['name'], cell=f"{C['qty_des']}{r1}", f=2),
                dict(op='set', sheet=a['name'], cell=f"{C['qty_aaa']}{r1}", f=1),
                dict(op='set', sheet=a['name'], cell=f"{C['rate']}{r1}", f=1000 * k),
                dict(op='set', sheet=a['name'], cell=f"{C['nrm']}{r2}", f="'" + code),
                dict(op='set', sheet=a['name'], cell=f"{C['qty_aaa']}{r2}", f=1),
                dict(op='set', sheet=a['name'], cell=f"{C['rate']}{r2}", f=7 * k)]
        want[f"{a['name']}!{C['pte']}{a['total']}"] = 1007 * k
        want[f"{a['name']}!{C['est']}{a['total']}"] = 2000 * k
        want[f"{a['name']}!{C['wa']}{r1}"] = 1000 / 1007
        if len(hs) > 1:
            want[f"{a['name']}!{C['pte']}{hs[0]['row']}"] = 1000 * k
            want[f"{a['name']}!{C['pte']}{hs[-1]['row']}"] = 7 * k
    ops.append(dict(op='calc'))
    tot = sum(1007 * k for k in range(1, len(expect['assets']) + 1))
    for k, a in enumerate(expect['assets'], 1):
        r = expect['asum_rows'][k - 1]
        inv = {v: s for s, v in expect['link_map'].items()}
        want[f"Asset Summary!{inv[C['pte']]}{r}"] = 1007 * k
        want[f"Asset Summary!{inv[C['est']]}{r}"] = 2000 * k
        want[f"Asset Summary!C{r}"] = a['name']
        want[f"Asset Summary!A{r}"] = 1007 * k / tot
        pr = expect['psum_rows'][k - 1]
        want[f"Project Summary!C{pr}"] = a['name']
        # Project Summary E.. = Asset Summary I.. x Repetition (1 in the template)
        want[f"Project Summary!{CL(CI(inv[C['pte']]) - 4)}{pr}"] = 1007 * k
        el = expect['elem_cols'][k - 1]
        want[f"{expect['elem']}!{el}2"] = a['name']
    reads = {}
    for key in want:
        sh, cell = key.rsplit('!', 1)
        reads.setdefault(sh, []).append(cell)
    # the elemental row of the planted code
    wb = openpyxl.load_workbook(out, read_only=True)
    e = wb[expect['elem']]
    code_row = next(r for r in expect['elem_sumif_rows'] if str(e[f'B{r}'].value) == code)
    wb.close()
    for k, col in enumerate(expect['elem_cols'], 1):
        planted = 1007 * k if expect['amount_col'] == C['pte'] else 2000 * k
        want[f"{expect['elem']}!{col}{code_row}"] = planted
        reads.setdefault(expect['elem'], []).append(f'{col}{code_row}')
    for sh, cells in reads.items():
        ops.append(dict(op='read', sheet=sh, cells=cells))
    got = excel.run_ops(out, ops)                     # no out= : the copy is closed without saving
    fails = []
    for key, w in want.items():
        g = got.get(key)
        ok = abs(g - w) < 1e-6 if isinstance(w, (int, float)) and isinstance(g, (int, float)) else g == w
        if not ok:
            fails.append(f'{key} = {g!r}, expected {w!r}')
    return fails, len(want)
