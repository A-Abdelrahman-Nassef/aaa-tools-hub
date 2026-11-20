"""boq_items.py - read a designer BOQ into items and headings, grouped by division (hub, 06-06 loading).

Adapted from the 06-06 skill's extract_boq.py for BOQs like the VT ones:
  * header row "Ser. | Item Description | Unit | Qty." found, not assumed;
  * old 16-division markers ("DIVISION 2 – SITE WORK", "DIVISION 16 – ELECTRICAL WORK"), repeated on every
    page; the page repeat of a division or Section heading is dropped;
  * an item may run over several rows: "1  Excavation ...", "The item shall include ...", and only the last
    row carries the unit and quantity. Every line stays its own row, wording as written, not joined
    (Ahmed, 18 Nov 2026); the line with the unit and quantity is the priced row;
  * junk dropped: "Total – Carried to Next Page", "Total of Previous Page", "Total Amount of ...",
    and everything before the first DIVISION marker (title block, General Notes).

read_boq(path) -> {"title": "B (5): Electrical Work", "divisions": [{"label", "num", "name", "rows": [...]}],
                   "counts": {...}, "notes": [...]}
row types: item {ser, desc, unit, qty, src} (priced), sec {desc} (Section heading, bold), text {ser, desc}
(every other line, plain), blank
"""
import re, warnings

import openpyxl

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
DIV = re.compile(r'^DIVISION\s*(\d+)\s*[–\-:]*\s*(.*)$', re.I)
JUNK = re.compile(r'^(total\b|carried\b|brought\b|page\s*\d+$)', re.I)


def _t(v):
    return re.sub(r'[ \t\r\f\v]+', ' ', str(v)).strip() if v is not None else ''


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _header(ws):
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = {c: _t(ws.cell(r, c).value).lower() for c in range(1, min(ws.max_column, 20) + 1)}
        cols = {}
        for c, v in vals.items():
            if 'ser' == v.rstrip('.') or v.startswith('item no') or v in ('item', 'no.', 'sr.', 's.n'):
                cols.setdefault('ser', c)
            elif 'description' in v:
                cols['desc'] = c
            elif v.rstrip('.') == 'unit':
                cols['unit'] = c
            elif v.rstrip('.') in ('qty', 'quantity', "q'ty"):
                cols['qty'] = c
        if {'desc', 'unit', 'qty'} <= set(cols):
            cols.setdefault('ser', cols['desc'] - 1)
            return r, cols
    raise ValueError('no "Item Description | Unit | Qty." header row found')


def read_boq(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = max(wb.worksheets, key=lambda s: s.max_row * min(s.max_column, 20))
    hr, C = _header(ws)
    title = ''
    for r in range(1, hr):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = _t(ws.cell(r, c).value)
            if re.match(r'^B\s*\(\d+\)', v):
                title = v
    divisions, notes = [], []
    cur = None
    sec_now = None
    has_div = any(DIV.match(_t(ws.cell(r, C['desc']).value)) for r in range(hr + 1, ws.max_row + 1))
    if not has_div:                         # e.g. the VT B(5) Electrical BOQ: items start right under the header
        cur = dict(label=f'(no DIVISION marker) {title or ws.title}', num=None, name=(title or '').upper(), rows=[])
        divisions.append(cur)
        notes.append('no DIVISION markers in this BOQ: all items are in one group, mapped by package')
    dropped_pre = 0

    def emit(row):
        cur['rows'].append(row)

    for r in range(hr + 1, ws.max_row + 1):
        ser, desc = _t(ws.cell(r, C['ser']).value), _t(ws.cell(r, C['desc']).value)
        unit, qty = _t(ws.cell(r, C['unit']).value), _num(ws.cell(r, C['qty']).value)
        if not (ser or desc or unit):
            continue
        m = DIV.match(desc)
        if m:
            num, name = int(m.group(1)), _t(m.group(2)).strip(' :-–').upper()
            label = f'DIVISION {num} - {name}'
            if not cur or cur['label'] != label:      # a new division (a page repeat of the same one changes nothing)
                cur = next((d for d in divisions if d['label'] == label), None)
                if not cur:
                    cur = dict(label=label, num=num, name=name, rows=[])
                    divisions.append(cur)
                sec_now = None
            continue
        if cur is None:
            dropped_pre += 1
            continue
        if JUNK.match(desc) and not unit:
            continue
        if re.match(r'^section\b', desc, re.I):
            key = re.sub(r'\W+', '', desc.lower())
            if key != sec_now:                     # a page repeat of the same Section heading is dropped
                emit(dict(type='sec', desc=desc, src=r))
                sec_now = key
            continue
        # every BOQ line stays its own row, as written (Ahmed, 18 Nov 2026: no joining of the wording);
        # the line that carries the unit and quantity is the priced row
        if qty is not None:
            emit(dict(type='item', ser=ser, desc=desc, unit=unit, qty=qty, src=r))
            if not unit:
                notes.append(f'row {r}: Ser. {ser or "-"} has a quantity ({qty:g}) but no unit')
        elif unit:                                 # a unit but no quantity: keep it visible, priced later
            emit(dict(type='item', ser=ser, desc=desc, unit=unit, qty=None, src=r))
            notes.append(f'row {r}: unit "{unit}" but no quantity')
        else:
            emit(dict(type='text', ser=ser, desc=desc, src=r))
    # a numbered line with no priced line under it before the next numbered line (and not a group whose
    # sub-items are numbered under it, 1 -> 1.01): the designer's layout, flagged
    for d in divisions:
        rows = d['rows']
        for i, x in enumerate(rows):
            if x['type'] == 'text' and x.get('ser'):
                nxt = next((y for y in rows[i + 1:] if y['type'] == 'sec' or y.get('ser') or y['type'] == 'item'), None)
                if nxt is None or (nxt['type'] != 'item' and not str(nxt.get('ser', '')).startswith(x['ser'] + '.')):
                    if not (nxt and nxt['type'] == 'item'):
                        notes.append(f"row {x['src']}: Ser. {x['ser']} \"{x['desc'][:40]}\" has no quantity and no sub-items")
    # one blank row before each new numbered item or heading that follows a priced line keeps items apart
    for d in divisions:
        out = []
        for x in d['rows']:
            if out and out[-1]['type'] == 'item' and (x['type'] == 'sec' or (x['type'] == 'text')):
                out.append(dict(type='blank'))
            out.append(x)
        d['rows'] = out
    items = sum(1 for d in divisions for x in d['rows'] if x['type'] == 'item')
    heads = sum(1 for d in divisions for x in d['rows'] if x['type'] in ('sec', 'text'))
    if dropped_pre:
        notes.insert(0, f'{dropped_pre} text row(s) before the first DIVISION (title block, general notes) left out')
    return dict(title=title, sheet=ws.title, divisions=divisions, notes=notes,
                counts=dict(items=items, headings=heads, qty=sum(x['qty'] or 0 for d in divisions for x in d['rows']
                                                                 if x['type'] == 'item')))


# ---- old (16-division) to current CSI divisions; approved by Ahmed 18 Nov 2026 for the VT BOQs
OLD_TO_NEW = {1: 1, 2: 31, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10, 11: 11, 12: 12, 13: 13, 14: 14, 16: 26}


def map_division(num, name, package):
    """-> (new division number or None, reason)."""
    p = (package or '').lower()
    if p.startswith('fire'):
        return 21, 'fire package -> 21 Fire Suppression'
    if num is None:                                   # BOQ without DIVISION markers: go by package
        for key, n in (('elec', 26), ('plum', 22), ('hvac', 23), ('mech', 23)):
            if p.startswith(key):
                return n, f'no DIVISION marker; {package} package -> {n}'
        return None, 'no DIVISION marker; choose a division'
    if num >= 20:
        return num, 'already a current division number'
    if num == 9 and 'PLUMB' in name.upper():
        return 22, '"9 - PLUMBING" -> 22 Plumbing'
    if num == 15:
        if p.startswith(('hvac', 'mech')):
            return 23, '15 Mechanical in an HVAC package -> 23'
        return 22, '15 Mechanical in a plumbing package -> 22 Plumbing'
    if num in OLD_TO_NEW:
        return OLD_TO_NEW[num], 'old -> current division'
    return None, 'no rule'
