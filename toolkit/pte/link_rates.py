"""link_rates.py - link repeated items across assets to one AAA rate (hub, 06-05 rules + 06-06 Path C).

Rules (06-05 "Asset Linking & Highlighting" prompt, Ahmed's choices of 18 Nov 2026):
  * assets in their order in the master (Asset Summary order), items top to bottom;
  * each item is compared first with the items ABOVE it in its own asset (Ahmed, 19 Nov 2026: when several BOQs
    are loaded into one asset, the repeats sit below each other), then with items of EARLIER assets, one asset
    at a time in order; the first place with a match wins and the search stops;
  * within one BOQ (same asset and column B code) an item links only under the same heading run: the same
    "Ø 75mm" line under Roof drain and under Floor drain is two items;
  * the BOQ-code headings 06-06 adds ("Arch - A2-1 - DIVISION 3 ...") are not part of an item's heading run;
  * BLUE   (0070C0): same division + heading run above the item + description + unit (canonical text);
  * ORANGE (FFC000): same division + description + unit under a different heading, or the same wording apart
    from spacing / punctuation / plurals / filler words, with the same unit and numbers. (A plain "90% alike"
    test was tried and rejected: it linked internal to external plaster and one-gang to two-gang switches.)
  * Division 08 Openings: never BLUE. Doors share codes but not sizes (Ahmed), so even identical text is only
    ORANGE, and nothing looser links;
  * no match: no link, no fill.
A link is the AAA Rate cell (column J): ='<source asset>'!J<row> (its own tab for a link within the asset).
Quantities stay each item's own.
"""
import re, sys, warnings
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(HERE / 'name_assets'))
import excel  # noqa: E402
from expanded import _header_cols, _find, _n  # noqa: E402

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
BLUE, ORANGE = 0xC07000, 0x00C0FF          # Excel colours are BGR: 0070C0 and FFC000
NUM = re.compile(r'\d+(?:[.,]\d+)?')
TAB = re.compile(r"^='?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$")


def canon(s):
    s = re.sub(r'\s+', ' ', str(s or '')).strip().lower()
    return s.strip(' .,:;-–')


def numbers(s):
    return sorted(x.replace(',', '.') for x in NUM.findall(str(s or '')))


FILLER = {'a', 'an', 'the', 'and', 'by', 'for', 'of', 'with', 'to', 'in', 'on', 'at', 'as', 'or', 'wire', 'wires',
          'complete', 'including', 'incl', 'all', 'type', 'item', 'items', 'mm'}


def words(s):
    """Meaningful words: letters/digits only, lower case, no plural s, no filler words; spacing ignored."""
    out = []
    for w in re.findall(r'[a-z0-9]+', str(s or '').lower()):
        if len(w) > 3 and w.endswith('s') and not w.endswith('ss'):
            w = w[:-1]
        if w not in FILLER:
            out.append(w)
    return out


def close(a, b):
    """Same wording apart from spacing, punctuation, plurals and filler words ('3X3mm2' = '3X3 mm2').
    A single different real word (internal/external, one/two gang, exhaust fan/lighting) is NOT close."""
    ca, cb = re.sub(r'[^a-z0-9]', '', a.lower()), re.sub(r'[^a-z0-9]', '', b.lower())
    return ca == cb or ''.join(words(a)) == ''.join(words(b))


def _code_heading(text, code):
    """True for the BOQ-code heading load_assign writes: "<code in column B> - <BOQ division label>"."""
    code = str(code or '').strip()
    return bool(code) and str(text).strip().lower().startswith(code.lower() + ' - ')


def read_items(path):
    """Assets in order, each with its item rows: division, heading run, description, unit."""
    wb = openpyxl.load_workbook(path)
    s = wb['Asset Summary']
    tabs = []
    for row in s.iter_rows():
        for c in row:
            m = TAB.match(str(c.value or ''))
            if m and m.group(1) in wb.sheetnames and m.group(1) not in tabs and not m.group(1).startswith(('Schedule', 'Scheduale')):
                tabs.append(m.group(1))
    assets = []
    for tab in tabs:
        ws = wb[tab]
        hr = _find(ws, 'A', 'NRM Code', range(1, 20))
        h = _header_cols(ws, hr)
        aaa = CI(h['AAA'])
        C = dict(sr=h['Sr.'], desc=h['BOQ Item Description'], unit=h.get('Unit', 'G'), qty=CL(aaa - 1),
                 rate=CL(aaa + 1), est=CL(aaa + 2), client=h.get('Client Code [If any]', 'B'))
        T = _find(ws, C['desc'], 'Total Amounts', range(hr, ws.max_row + 1))
        # BOQ lines are loaded one per row (not joined), so an item is rebuilt from the rows:
        #   sections  bold Section / BOQ-code headings          titles  unnumbered title lines after items
        #   group     the lines of a numbered group whose sub-items carry their own Ser. (1 -> 1.01, 1.02)
        #   block     an open numbered item: "1 Excavation...", "The item shall include..." up to its priced line
        items, div = [], None
        sections, titles, group, block, block_sr, last = [], [], [], None, '', ''
        for r in range(hr + 2, T):
            d = ws[f"{C['desc']}{r}"]
            text = d.value
            f = str(ws[f"{C['est']}{r}"].value or '')
            sr = str(ws[f"{C['sr']}{r}"].value or '').strip()
            if isinstance(text, str) and (text.startswith('Division') or text.startswith('Package ')) \
                    and not f.startswith('=$'):
                m = re.search(r'Division\s+(\d+)', text)
                div = int(m.group(1)) if m and text.startswith('Division') else None
                sections, titles, group, block, last = [], [], [], None, ''
                continue
            unit = ws[f"{C['unit']}{r}"].value
            priced = f.startswith('=') and (text or unit or ws[f"{C['qty']}{r}"].value is not None)
            if not text and not priced:
                continue
            text = str(text or '')
            if priced:                                               # a priced line (its text may be empty)
                if block is not None and sr:                         # numbered sub-item under an open group
                    group, block, desc = block, None, text
                elif block is not None:                              # the last line of a multi-line item
                    desc, block = '\n'.join(block + [text]), None
                else:
                    desc = text
                items.append(dict(row=r, div=div, parent=' / '.join(sections + titles + group), desc=desc,
                                  unit=str(ws[f"{C['unit']}{r}"].value or ''), sr=sr,
                                  package=ws[f"{C['client']}{r}"].value, rate=ws[f"{C['rate']}{r}"].value))
                last = 'item'
            elif d.font is not None and d.font.b and _code_heading(text, ws[f"{C['client']}{r}"].value):
                # "Arch - A2-1 - DIVISION 3 - CONCRETE": the heading 06-06 adds when two BOQs share a master division.
                # It only says which BOQ follows, so it starts a fresh heading run but is not part of it (else every
                # repeat between two BOQs would differ by this line and never be BLUE).
                sections, titles, group, block, last = [], [], [], None, ''
            elif d.font is not None and d.font.b:                    # Section heading
                sections = (sections if last == 'sec' else []) + [text]
                titles, group, block, last = [], [], None, 'sec'
            elif sr:                                                 # a numbered line starts an item or a group
                if block is not None and block_sr and sr.startswith(block_sr + '.'):
                    block.append(text)                               # e.g. "3.01 Ø 25mm" under an open "3"
                else:
                    group, block, block_sr = [], [text], sr
                last = 'num'
            elif block is not None:                                  # "The item shall include ..."
                block.append(text)
            else:                                                    # an unnumbered title line
                titles = (titles if last == 'title' else []) + [text]
                group, last = [], 'title'
        assets.append(dict(tab=tab, cols=C, items=items))
    return assets


def match(assets):
    """-> list of links {tab, row, src_tab, src_row, colour, why, same_asset}.
    Each item looks first at the items above it in its own asset, then at earlier assets in order (Ahmed,
    19 Nov 2026: repeated items sit below each other when several BOQs are loaded into one asset)."""
    links = []
    for k, a in enumerate(assets):
        for i, it in enumerate(a['items']):
            d, p, x, u = it['div'], canon(it['parent']), canon(it['desc']), canon(it['unit'])
            own = dict(tab=a['tab'], items=a['items'][:i], same=True)
            for src in [own] + [dict(b, same=False) for b in assets[:k]]:    # own asset above it, then earlier assets
                best = None
                for s in src['items']:
                    if s['div'] != d or canon(s['unit']) != u:
                        continue
                    sx, sp = canon(s['desc']), canon(s['parent'])
                    if src['same'] and s['package'] == it['package'] and sp != p:
                        continue    # one BOQ never repeats an item under another heading: "Ø 75mm" under Roof
                                    # drain is not the "Ø 75mm" under Floor drain
                    if d == 8:                                   # Openings: same code can mean a different size
                        if sx == x and sp == p:
                            best = (s, 'orange', 'Openings: same door / window code and text, but sizes can differ '
                                                 'between assets - check the schedule'); break
                        continue
                    if sx == x and sp == p:
                        best = (s, 'blue', 'same division, heading, description and unit'); break
                    if best:
                        continue
                    if sx == x:
                        best = (s, 'orange', 'same description and unit, different heading')
                    elif numbers(sx) == numbers(x) and close(sx, x) and \
                            (sp == p or close(sp, p) or numbers(sp) == numbers(p)):
                        best = (s, 'orange', 'same wording apart from spacing / punctuation / filler words, '
                                             'same unit and numbers')
                if best:
                    s, colour, why = best
                    if src['same']:
                        why = 'same asset, an item above it: ' + why.replace('between assets', 'between BOQs')
                    links.append(dict(tab=a['tab'], row=it['row'], src_tab=src['tab'], src_row=s['row'], colour=colour,
                                      why=why, desc=it['desc'], src_desc=s['desc'], unit=it['unit'], div=d,
                                      package=it['package'], src_package=s['package'], same_asset=src['same']))
                    break                                            # stop at the first asset with a match
    return links


def problems(assets):
    out = []
    for a in assets:
        done = [it['row'] for it in a['items'] if it['rate'] not in (None, '')]
        if done:
            out.append(f"{a['tab']}: {len(done)} item(s) already have a rate or a link (e.g. row {done[0]}); "
                       'start from the master as 06-06 loaded it')
    if not assets or not any(a['items'] for a in assets):
        out.append('no loaded items found; run 06-06 first')
    return out


def plan(assets, links):
    rate = {a['tab']: a['cols']['rate'] for a in assets}
    ops = [dict(op='set', sheet=l['tab'], cell=f"{rate[l['tab']]}{l['row']}",
                f=f"='{l['src_tab']}'!{rate[l['src_tab']]}{l['src_row']}") for l in links]
    for tab in rate:
        for colour, code in (('blue', BLUE), ('orange', ORANGE)):
            cells = [f"{rate[tab]}{l['row']}" for l in links if l['tab'] == tab and l['colour'] == colour]
            if cells:
                ops.append(dict(op='fill', sheet=tab, cells=cells, color=code))
    return ops


def verify(out, assets, links):
    """Static: every link formula as planned. Numbers: unique rates on all unlinked items in a throwaway copy;
    each linked rate must show its source's rate."""
    wb = openpyxl.load_workbook(out)
    rate = {a['tab']: a['cols']['rate'] for a in assets}
    fails = []
    for l in links:
        v = wb[l['tab']][f"{rate[l['tab']]}{l['row']}"].value
        if _n(v) != _n(f"='{l['src_tab']}'!{rate[l['src_tab']]}{l['src_row']}"):
            fails.append(f"{l['tab']}!{rate[l['tab']]}{l['row']} = {v}")
    linked = {(l['tab'], l['row']) for l in links}
    ops, want, n = [], {}, 0
    for a in assets:
        for it in a['items']:
            if (a['tab'], it['row']) not in linked:
                n += 1
                ops.append(dict(op='set', sheet=a['tab'], cell=f"{rate[a['tab']]}{it['row']}", f=1000 + n))
    srcval = {}
    n = 0
    for a in assets:
        for it in a['items']:
            if (a['tab'], it['row']) not in linked:
                n += 1
                srcval[(a['tab'], it['row'])] = 1000 + n
    to = {(l['tab'], l['row']): (l['src_tab'], l['src_row']) for l in links}

    def root(key):                    # a source may itself be linked to an earlier asset: follow the chain
        seen = set()
        while key in to and key not in seen:
            seen.add(key); key = to[key]
        return key
    reads = {}
    for l in links:
        key = f"{l['tab']}!{rate[l['tab']]}{l['row']}"
        want[key] = srcval.get(root((l['src_tab'], l['src_row'])))
        reads.setdefault(l['tab'], []).append(f"{rate[l['tab']]}{l['row']}")
    ops.append(dict(op='calc'))
    for sh, cells in reads.items():
        ops.append(dict(op='read', sheet=sh, cells=cells))
    ops.append(dict(op='errors'))
    got = excel.run_ops(out, ops) if links else {}
    for k, w in want.items():
        if got.get(k) != w:
            fails.append(f'{k} = {got.get(k)!r}, expected {w!r} from its source')
    for k, v in got.items():
        if k.startswith('errors!') and not k.endswith('!first'):
            fails.append(f"{k[7:]}: {int(v)} formula error(s), first at {got.get(k + '!first')}")
    return fails, len(want)


def report(path, assets, links):
    """Link report workbook: one row per linked item, for the manual check of the orange ones."""
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'Links'
    ws.append(['Colour', 'Asset', 'Row', 'Package', 'Division', 'Description', 'Unit', 'Linked to asset', 'Row',
               'Package', 'Source description', 'Why'])
    from openpyxl.styles import Font, PatternFill
    fills = dict(blue=PatternFill('solid', fgColor='0070C0'), orange=PatternFill('solid', fgColor='FFC000'))
    for l in sorted(links, key=lambda l: (l['colour'] != 'orange', l['tab'], l['row'])):
        ws.append([l['colour'].upper(), l['tab'], l['row'], l['package'], l['div'], l['desc'][:500], l['unit'],
                   l['src_tab'], l['src_row'], l['src_package'], l['src_desc'][:500], l['why']])
        ws.cell(ws.max_row, 1).fill = fills[l['colour']]
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip('ABCDEFGHIJKL', (9, 22, 7, 9, 9, 60, 7, 22, 7, 9, 60, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet('Per asset')
    s.append(['Asset', 'Items', 'Blue', 'Orange', 'Not linked', 'Within the asset'])
    for k, a in enumerate(assets):
        b = sum(1 for l in links if l['tab'] == a['tab'] and l['colour'] == 'blue')
        o = sum(1 for l in links if l['tab'] == a['tab'] and l['colour'] == 'orange')
        own = sum(1 for l in links if l['tab'] == a['tab'] and l.get('same_asset'))
        s.append([a['tab'], len(a['items']), b, o, len(a['items']) - b - o, own])
    s.column_dimensions['A'].width = 30
    wb.save(path)
