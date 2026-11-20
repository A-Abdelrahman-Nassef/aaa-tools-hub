"""load_master.py - load designer BOQ items into the named PTE master (hub, 06-06 Path A).

Input: the master built by 06-04 (one tab per asset, a block per package: "Package A : Arch", divisions
"A.01 Division 01 : ...", each division a header + 3 template rows, Total Amounts at the end), the register
(file -> asset, package) and the BOQs read by boq_items.py.

What it does, in Excel (excel_ops.ps1) so every reference moves with the rows:
  1. bottom to top, each division gets exactly as many rows as its BOQ content (rows inserted inside the block,
     so Excel stretches the SUBTOTALs and moves the Asset Summary links with the Total Amounts row);
  2. whole columns written in one call each: B Client Code (= package, per the 06-05 house rules), C/D weights,
     E Ser., F description, G unit, H Designer QTY, and every amount column (the master's own item formulas);
     SUBTOTALs for division, package and total rows; text as text ('1.01' stays 1.01, '- 40 mm' is not a formula);
  3. Section headings bold, every other BOQ line plain and as written (no joining), descriptions wrapped,
     row heights fitted;
  4. spacing (Ahmed, 18 Nov 2026): at least 1 empty row under each loaded division heading and 3 empty rows
     between the last row of a loaded division and the next division heading.
Never touched: column A NRM Code, the AAA / tenderer QTY and Rate inputs, columns BC-BP, the other sheets.
Divisions with no BOQ content keep the template's 3 placeholder rows.
"""
import re, sys, warnings
from pathlib import Path

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(HERE / 'name_assets'))
import excel  # noqa: E402
from expanded import _header_cols, _find, _n  # noqa: E402

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
ABS = re.compile(r"(\$?[A-Z]{1,3})\$(\d+)\b")
DIVNUM = re.compile(r'Division\s+(\d+)', re.I)


class LoadError(Exception):
    pass


def _txt(v):
    """Text for Excel: always stored as text (a leading ' is not shown)."""
    return '' if v in (None, '') else "'" + str(v)


def _move(f, r_from, r_to, amap):
    if not (isinstance(f, str) and f.startswith('=')):
        return f or ''
    f = Translator(f, origin=f'A{r_from}').translate_formula(f'A{r_to}')
    return ABS.sub(lambda m: f'{m.group(1)}${amap.get(int(m.group(2)), int(m.group(2)))}', f)


def detect(path, tabs):
    """Layout of the named master's asset tabs (only the tabs we load)."""
    wb = openpyxl.load_workbook(path)
    out = {}
    for tab in tabs:
        if tab not in wb.sheetnames:
            raise LoadError(f'the master has no tab "{tab}"')
        ws = wb[tab]
        hr = _find(ws, 'A', 'NRM Code', range(1, 20))
        h = _header_cols(ws, hr)
        aaa = CI(h['AAA'])
        C = dict(nrm='A', client=h.get('Client Code [If any]', 'B'), wa=h['Weight/ Asset'], wd=h['Weight/ Division'],
                 sr=h['Sr.'], desc=h['BOQ Item Description'], unit=h.get('Unit', 'G'), qty_des=CL(aaa - 1),
                 qty_aaa=CL(aaa), rate=CL(aaa + 1), est=CL(aaa + 2), pte=CL(aaa + 3))
        T = _find(ws, C['desc'], 'Total Amounts', range(hr, ws.max_row + 1))
        sum_cols = [c.column_letter for c in ws[T] if str(c.value or '').upper().startswith('=SUBTOTAL(')]
        m = re.search(r'\(9,[A-Z]+(\d+):', str(ws[f"{C['pte']}{T}"].value))
        body_start = int(m.group(1))
        pk_rows = [r for r in range(hr + 2, T) if str(ws[f"{C['desc']}{r}"].value or '').startswith('Package ')]
        div_rows = [r for r in range(hr + 2, T) if str(ws[f"{C['desc']}{r}"].value or '').startswith('Division')]
        if not div_rows:
            raise LoadError(f'{tab}: no division rows')
        blocks = []
        starts = pk_rows or [div_rows[0] - 1]
        for i, s in enumerate(starts):
            end = (starts[i + 1] - 1) if i + 1 < len(starts) else T - 1
            divs = [r for r in div_rows if s < r <= end]
            name = str(ws[f"{C['desc']}{s}"].value).split(':', 1)[-1].strip() if pk_rows else ''
            letter = ws[f"{C['sr']}{s}"].value if pk_rows else ''
            dl = []
            for j, d in enumerate(divs):
                nxt = divs[j + 1] if j + 1 < len(divs) else None
                last = (nxt - 1) if nxt else end - (1 if pk_rows else 0)     # the package's trailing blank row
                mm = DIVNUM.search(str(ws[f"{C['desc']}{d}"].value))
                dl.append(dict(row=d, num=int(mm.group(1)) if mm else None, last=last,
                               sr=ws[f"{C['sr']}{d}"].value, text=ws[f"{C['desc']}{d}"].value))
            blocks.append(dict(header=s if pk_rows else None, name=name, letter=letter, divs=dl, end=end))
        # template rows of the first division (header + its rows) for placeholders and item formulas
        d0 = blocks[0]['divs'][0]
        tmpl = {}
        for k, r in enumerate(range(d0['row'] + 1, d0['last'] + 1)):
            tmpl[k] = (r, {c: ws[f'{c}{r}'].value for c in [C['wa'], C['wd']] + sum_cols})
        item_src = next(r for r, v in tmpl.values() if str(v.get(C['est']) or '').startswith('='))
        out[tab] = dict(cols=C, total=T, sum_cols=sum_cols, body_start=body_start, blocks=blocks, tmpl=tmpl,
                        item_row=item_src, item=tmpl[[k for k, (r, _) in tmpl.items() if r == item_src][0]][1],
                        title_row=hr, unloaded=_is_blank(ws, C, blocks))
    return out


def _is_blank(ws, C, blocks):
    """True when nothing is in Description / Unit / Designer QTY under any division (not loaded yet)."""
    for b in blocks:
        for d in b['divs']:
            for r in range(d['row'] + 1, d['last'] + 1):
                if ws[f"{C['desc']}{r}"].value or ws[f"{C['unit']}{r}"].value or ws[f"{C['qty_des']}{r}"].value:
                    return False
    return True


def plan(L, loads):
    """loads: {tab: {block index: [ {"num": master division number, "rows": [...], "package": "Arch"} ... ]}}.

    Returns (ops, expect). expect holds, per tab, the final row of everything so the checks can read it back."""
    ops, expect = [], dict(tabs={})
    for tab, by_block in loads.items():
        Lt = L[tab]; C = Lt['cols']; T0 = Lt['total']
        # content per division row (old numbering): rows to write
        content = {}
        for bi, parts in by_block.items():
            b = Lt['blocks'][bi]
            for part in parts:
                d = next((d for d in b['divs'] if d['num'] == part['num']), None)
                if d is None:
                    raise LoadError(f"{tab} block {b['name'] or bi + 1}: no Division {part['num']:02d}")
                rows = [dict(r, package=part['package']) for r in part['rows']]
                while rows and rows[0]['type'] == 'blank':
                    rows.pop(0)
                while rows and rows[-1]['type'] == 'blank':
                    rows.pop()
                if d['row'] in content:                      # a second part in the same division: 1 empty row between
                    content[d['row']].append(dict(type='blank'))
                content.setdefault(d['row'], []).extend(rows)
        # spacing rule: 1 empty row under the division heading, 3 empty rows before the next division heading
        for k in content:
            content[k] = [dict(type='blank')] + content[k] + [dict(type='blank')] * 3
        # 1. structure, bottom to top
        all_divs = [d for b in Lt['blocks'] for d in b['divs']]
        for d in sorted(all_divs, key=lambda d: -d['row']):
            have, want = d['last'] - d['row'], len(content.get(d['row'], [])) or None
            if want is None or want == have:
                continue
            if want > have:
                ops.append(dict(op='insert_rows', sheet=tab, at=d['row'] + 2, count=want - have))
            else:
                ops.append(dict(op='delete_rows', sheet=tab, rows=f"{d['row'] + want + 1}:{d['last']}"))
        # 2. the final layout, top to bottom
        shift, layout = 0, []                      # layout: (old row or None, new row, kind, payload)
        first = Lt['blocks'][0]['header'] or Lt['blocks'][0]['divs'][0]['row']
        r_new = first
        final_blocks = []
        for b in Lt['blocks']:
            fb = dict(name=b['name'], letter=b['letter'], divs=[])
            if b['header']:
                fb['header'] = r_new
                layout.append(('pkg', r_new, b)); r_new += 1
            for d in b['divs']:
                rows = content.get(d['row'])
                fd = dict(num=d['num'], header=r_new, sr=d['sr'], text=d['text'])
                layout.append(('div', r_new, d)); r_new += 1
                if rows:
                    fd['rows'] = list(range(r_new, r_new + len(rows)))
                    fd['items'] = [(r_new + k, x) for k, x in enumerate(rows) if x['type'] == 'item']
                    for x in rows:
                        layout.append(('row', r_new, x)); r_new += 1
                else:
                    fd['rows'] = list(range(r_new, r_new + (d['last'] - d['row'])))
                    fd['items'] = []
                    for k in range(d['last'] - d['row']):
                        layout.append(('ph', r_new, k)); r_new += 1
                fd['last'] = r_new - 1
                fb['divs'].append(fd)
            if b['header']:                                   # the package's trailing blank row
                layout.append(('sep', r_new, None)); r_new += 1
            fb['end'] = r_new - 1
            final_blocks.append(fb)
        T = r_new
        # 3. columns, one call each
        cols = [C['client'], C['wa'], C['wd'], C['sr'], C['desc'], C['unit'], C['qty_des']] + \
               [c for c in Lt['sum_cols'] if c not in (C['wa'], C['wd'])]
        vals = {c: [] for c in cols}
        bold, italic, plain, wrap_rows = [], [], [], []
        div_of = {}
        for fb in final_blocks:
            for fd in fb['divs']:
                for r in fd['rows']:
                    div_of[r] = fd['header']
        pk_of = {}
        for fb in final_blocks:
            if fb.get('header'):
                for r in range(fb['header'], fb['end'] + 1):
                    pk_of[r] = fb
        for kind, r, p in layout:
            v = {c: '' for c in cols}
            if kind == 'pkg':
                fb = pk_of[r]
                for c in [C['wa'], C['wd']] + Lt['sum_cols']:
                    if c in v:
                        v[c] = f'=SUBTOTAL(9,{c}{r + 1}:{c}{fb["end"]})'
                v[C['sr']] = _txt(p['letter']); v[C['desc']] = _txt(f"Package {p['letter']} : {p['name']}")
            elif kind == 'div':
                fd = next(fd for fb in final_blocks for fd in fb['divs'] if fd['header'] == r)
                for c in [C['wa'], C['wd']] + Lt['sum_cols']:
                    if c in v:
                        v[c] = f"=SUBTOTAL(9,{c}{r + 1}:{c}{fd['last']})"
                v[C['sr']] = _txt(p['sr']); v[C['desc']] = _txt(p['text'])
            elif kind == 'ph':                                # untouched division: the template's own rows
                if p not in Lt['tmpl']:                       # e.g. the last division's extra blank row
                    for c in cols:
                        vals[c].append(v[c])
                    continue
                src, f = Lt['tmpl'][p]
                amap = {T0: T, Lt['blocks'][0]['divs'][0]['row']: div_of[r]}
                for c in [C['wa'], C['wd']] + Lt['sum_cols']:
                    if c in v:
                        v[c] = _move(f.get(c), src, r, amap)
            elif kind == 'row':
                x = p
                if x['type'] == 'item':
                    amap = {T0: T, Lt['blocks'][0]['divs'][0]['row']: div_of[r]}
                    for c in [C['wa'], C['wd']] + Lt['sum_cols']:
                        if c in v:
                            v[c] = _move(Lt['item'].get(c), Lt['item_row'], r, amap)
                    v[C['client']] = _txt(x['package'])
                    v[C['sr']] = _txt(x.get('ser'))
                    v[C['desc']] = _txt(x['desc'])
                    v[C['unit']] = _txt(x.get('unit'))
                    v[C['qty_des']] = '' if x.get('qty') is None else x['qty']
                    wrap_rows.append(r)
                elif x['type'] in ('sec', 'text', 'mat'):
                    v[C['client']] = _txt(x.get('package'))
                    v[C['sr']] = _txt(x.get('ser'))
                    v[C['desc']] = _txt(x['desc'])
                    {'sec': bold, 'mat': italic}.get(x['type'], plain).append(f"{C['desc']}{r}")
                    wrap_rows.append(r)
            for c in cols:
                vals[c].append(v[c])
        for c in cols:
            ops.append(dict(op='column', sheet=tab, col=c, **{'from': first}, values=vals[c]))
        for c in Lt['sum_cols']:
            ops.append(dict(op='set', sheet=tab, cell=f'{c}{T}', f=f"=SUBTOTAL(9,{c}{Lt['body_start']}:{c}{T - 1})"))
        if bold:
            ops.append(dict(op='font', sheet=tab, cells=bold, bold=True, italic=False))
        if italic:
            ops.append(dict(op='font', sheet=tab, cells=italic, italic=True, bold=False))
        if plain:
            ops.append(dict(op='font', sheet=tab, cells=plain, italic=False, bold=False))
        # outline: package 1, division 2, rows 3; wrap + fit heights on loaded rows
        runs = []
        for kind, r, p in layout:
            lv = 1 if kind == 'pkg' else (2 if kind in ('div', 'sep') else 3)
            if runs and runs[-1][2] == lv and runs[-1][1] == r - 1:
                runs[-1][1] = r
            else:
                runs.append([r, r, lv])
        if Lt['blocks'][0]['header'] is None:           # no packages: divisions 1, rows 2
            runs = [[a, b, lv - 1] for a, b, lv in runs]
        for a, b, lv in runs:
            ops.append(dict(op='outline', sheet=tab, rows=f'{a}:{b}', level=lv))
        loaded = sorted(set(wrap_rows))
        if loaded:
            ops.append(dict(op='wrap', sheet=tab, range=f"{C['desc']}{first}:{C['desc']}{T - 1}"))
            run = [loaded[0], loaded[0]]
            for r in loaded[1:] + [None]:
                if r is not None and r == run[1] + 1:
                    run[1] = r; continue
                ops.append(dict(op='autofit', sheet=tab, rows=f'{run[0]}:{run[1]}'))
                if r is not None:
                    run = [r, r]
        expect['tabs'][tab] = dict(total=T, blocks=final_blocks, cols=C, sum_cols=Lt['sum_cols'],
                                   body_start=Lt['body_start'])
    return ops, expect


# ---------------- checks ----------------
def _norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def verify_static(out, expect, asum_rows):
    """Every loaded row read back against its BOQ row; totals, subtotals and the Asset Summary link."""
    wb = openpyxl.load_workbook(out)
    fails, counted = [], 0
    for tab, e in expect['tabs'].items():
        ws = wb[tab]; C = e['cols']; T = e['total']
        if ws[f"{C['desc']}{T}"].value != 'Total Amounts':
            fails.append(f'{tab}: Total Amounts is not on row {T}')
        for c in e['sum_cols']:
            want = f"=SUBTOTAL(9,{c}{e['body_start']}:{c}{T - 1})"
            if ws[f'{c}{T}'].value != want:
                fails.append(f'{tab}!{c}{T} = {ws[f"{c}{T}"].value}')
        for fb in e['blocks']:
            if fb.get('header'):
                f = ws[f"{C['pte']}{fb['header']}"].value
                if f != f"=SUBTOTAL(9,{C['pte']}{fb['header'] + 1}:{C['pte']}{fb['end']})":
                    fails.append(f"{tab}: package {fb['letter']} subtotal is {f}")
            for fd in fb['divs']:
                f = ws[f"{C['pte']}{fd['header']}"].value
                if f != f"=SUBTOTAL(9,{C['pte']}{fd['header'] + 1}:{C['pte']}{fd['last']})":
                    fails.append(f"{tab}: division row {fd['header']} subtotal is {f}")
                for r, x in fd['items']:
                    counted += 1
                    got = (_norm(ws[f"{C['desc']}{r}"].value), _norm(ws[f"{C['unit']}{r}"].value),
                           ws[f"{C['qty_des']}{r}"].value, _norm(ws[f"{C['sr']}{r}"].value))
                    want = (_norm(x['desc']), _norm(x.get('unit')), x.get('qty'), _norm(x.get('ser')))
                    if got[:2] != want[:2] or got[3] != want[3] or (want[2] is not None and abs((got[2] or 0) - want[2]) > 1e-9):
                        fails.append(f'{tab}!row {r}: {got} vs BOQ {want}')
                    for c, pat in ((C['est'], f"=${C['qty_des']}{r}*{C['rate']}{r}"),
                                   (C['pte'], f"=${C['qty_aaa']}{r}*{C['rate']}{r}")):
                        if ws[f'{c}{r}'].value != pat:
                            fails.append(f'{tab}!{c}{r} = {ws[f"{c}{r}"].value}, expected {pat}')
                    fw = str(ws[f"{C['wd']}{r}"].value or '')
                    if f"/{C['pte']}${fd['header']})" not in fw or f"/${C['pte']}${T})" not in str(ws[f"{C['wa']}{r}"].value):
                        fails.append(f'{tab}!row {r}: weight formulas not over division {fd["header"]} / total {T}')
    s = wb['Asset Summary']
    for tab, (ar, ac) in asum_rows.items():
        e = expect['tabs'][tab]
        v = _n(s[f'{ac}{ar}'].value)
        if not v.endswith(f"!${e['cols']['pte']}${e['total']}") or _n(tab) + '!' not in v:
            fails.append(f'Asset Summary {ac}{ar} = {s[f"{ac}{ar}"].value}, expected the {tab} total row {e["total"]}')
    return fails[:40], counted


def asset_summary_links(path, tabs, L):
    """{tab: (row, column)} of the Asset Summary cell that carries each tab's Pre-Tender total."""
    ws = openpyxl.load_workbook(path)['Asset Summary']
    out = {}
    for row in ws.iter_rows():
        for c in row:
            v = _n(c.value or '')
            for tab in tabs:
                if v == _n(f"='{tab}'!${L[tab]['cols']['pte']}${L[tab]['total']}"):
                    out[tab] = (c.row, c.column_letter)
    return out


def verify_numbers(out, expect, asum_rows):
    """Throwaway copy: AAA QTY = Designer QTY and Rate = 1 on every loaded item, recalculated by Excel.
    Each package, tab total and Asset Summary row must equal the BOQ quantities; no formula errors anywhere."""
    ops, want = [], {}
    for tab, e in expect['tabs'].items():
        C = e['cols']
        items = [(r, x) for fb in e['blocks'] for fd in fb['divs'] for r, x in fd['items']]
        if not items:
            continue
        lo, hi = min(r for r, _ in items), max(r for r, _ in items)
        q = {r: x.get('qty') or 0 for r, x in items}
        ops.append(dict(op='column', sheet=tab, col=C['qty_aaa'], **{'from': lo},
                        values=[q.get(r, '') if r in q else '' for r in range(lo, hi + 1)]))
        ops.append(dict(op='column', sheet=tab, col=C['rate'], **{'from': lo},
                        values=[1 if r in q else '' for r in range(lo, hi + 1)]))
        total = sum(q.values())
        want[f"{tab}!{C['pte']}{e['total']}"] = total
        want[f"{tab}!{C['est']}{e['total']}"] = total
        for fb in e['blocks']:
            if fb.get('header'):
                want[f"{tab}!{C['pte']}{fb['header']}"] = sum(q[r] for fd in fb['divs'] for r, _ in fd['items'])
        if tab in asum_rows:
            ar, ac = asum_rows[tab]
            want[f"Asset Summary!{ac}{ar}"] = total
    ops.append(dict(op='errors'))
    reads = {}
    for k in want:
        sh, cell = k.rsplit('!', 1)
        reads.setdefault(sh, []).append(cell)
    for sh, cells in reads.items():
        ops.append(dict(op='read', sheet=sh, cells=cells))
    got = excel.run_ops(out, ops)
    fails = [f'{k} = {got.get(k)!r}, expected {w:g}' for k, w in want.items()
             if not isinstance(got.get(k), (int, float)) or abs(got[k] - w) > 1e-6 * max(1, abs(w))]
    errs = {k: v for k, v in got.items() if k.startswith('errors!') and not k.endswith('!first')}
    for k, v in errs.items():
        fails.append(f"{k[7:]}: {int(v)} formula error(s), first at {got.get(k + '!first')}")
    return fails, len(want)
