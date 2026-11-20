"""build_workbook.py - one collected AAA Areas Calculations workbook for one or more assets.
usage: python build_workbook.py template.xlsx out.xlsx asset1.json [asset2.json ...]
template = the blank AAA Areas Calculations template (sheets Inclusions, Summary, BS 00, Building 01).
Each asset JSON comes from asset_lib.dump(). One sheet per asset (cloned from 'Building 01'),
one shared Room Schedule, Summary rebuilt for the assets in use only, Notes sheet first.
Changed template cells are shaded orange; measured inputs are blue; links are green."""
import sys, json, copy, warnings, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings('ignore')
sys.path.insert(0, __file__.rsplit('/', 1)[0] if '/' in __file__ else '.')
from asset_lib import OPEN_CODE, IN_BUA, IN_IPMS, VOID

TPL, OUT, FILES = sys.argv[1], sys.argv[2], sys.argv[3:]
A = [json.load(open(f)) for f in FILES]
wb = openpyxl.load_workbook(TPL)
ORANGE = PatternFill('solid', fgColor='FFF4B183')
BLUE = Font(name='Aptos', size=11, color='FF0000FF'); GREEN = Font(name='Aptos', color='FF008000')
NUM = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'
RS = "'Room Schedule'!"
ITEM_ROWS = {15: 'A1', 27: 'B1.1', 28: 'B1.2', 32: 'B2.1', 39: 'C1.1', 40: 'C1.2', 41: 'C1.3', 42: 'C1.4', 48: 'D1.1', 49: 'D1.2',
             55: 'E1.1', 61: 'F1.1', 62: 'F1.2', 63: 'F1.3', 64: 'F1.4', 70: 'G1.1', 71: 'G1.2', 75: 'G2.1',
             81: 'H1.1', 82: 'H1.2', 83: 'H1.3', 84: 'H1.4'}
COLS = 'CDEFGHI'

# ---------- Room Schedule (all assets) ----------
rs = wb.create_sheet('Room Schedule')
hdr = ['Asset', 'Level', 'Space', 'Drawing ref', 'Code', 'Area (m²)', 'In external skin', 'AAA BUA', 'IPMS 1', 'Void counted here', 'Note']
for i, h in enumerate(hdr, 1):
    c = rs.cell(row=2, column=i, value=h); c.font = Font(name='Aptos', bold=True, color='FFFFFFFF')
    c.fill = PatternFill('solid', fgColor='FF1E3A66'); c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
rs['A1'] = 'Spaces measured from the drawings. Rooms to inner wall faces; open areas to their outer edge. Atrium voids (B1.3): Yes = counted on this floor, No = removed from this floor.'
rs['A1'].font = Font(name='Aptos', bold=True)
NOTE = {'perg': 'Pergola of open slats: treated as uncovered', 'cov': 'Covered by the slab above', 'unc': 'Open to the sky',
        'plant': 'Plant roof, maintenance access only: not counted', 'bal': 'Balcony', 'shel': 'Sheltered area', 'nb': 'Notional boundary area'}
r = 3
for a in A:
    lab = {l['id']: l['label'] for l in a['levels']}
    for Lv in a['levels']:
        k = Lv['id']
        allrooms = list(a['rooms'][k]) + [dict(code=VOID, name=o['name'].replace(' (removed)', ''), ref='', area=o['area'], void=True, counted=False)
                                          for o in a['opens'] if o['level'] == k and o['key'] == 'void']
        for x in allrooms:
            v = x.get('void')
            vals = [a['asset'], lab[k], x['name'], x.get('ref', ''), x['code'], x['area'], 'Yes',
                    'Yes' if (not v or x['counted']) else 'No', 'Yes' if (not v or x['counted']) else 'No',
                    ('Yes' if x['counted'] else 'No') if v else '', 'Atrium void' + ('' if x['counted'] else ', removed from this floor') if v else '']
            for i, val in enumerate(vals, 1): rs.cell(row=r, column=i, value=val)
            rs.cell(row=r, column=6).font = BLUE; r += 1
        for o in a['opens']:
            if o['level'] != k or o['key'] == 'void': continue
            vals = [a['asset'], lab[k], o['name'], '', OPEN_CODE[o['key']], o['area'], 'No', 'Yes' if IN_BUA[o['key']] else 'No',
                    'Yes' if IN_IPMS[o['key']] else 'No', '', o.get('note') or NOTE[o['key']]]
            for i, val in enumerate(vals, 1): rs.cell(row=r, column=i, value=val)
            rs.cell(row=r, column=6).font = BLUE; r += 1
last = r - 1
for row in rs.iter_rows(min_row=3, max_row=last):
    for c in row:
        if c.font.color is None or c.font.color.rgb != 'FF0000FF': c.font = Font(name='Aptos')
        c.border = Border(bottom=Side(style='thin', color='FFD5DBE3'))
    row[5].number_format = NUM
for col, w in zip('ABCDEFGHIJK', [16, 10, 36, 11, 11, 12, 14, 10, 10, 12, 46]): rs.column_dimensions[col].width = w
rs.freeze_panes = 'A3'
rs.cell(row=last + 2, column=1, value='Blue figures are measured inputs. Each asset sheet picks them up with SUMIFS by asset, level and code.').font = Font(name='Aptos', italic=True)

# ---------- asset sheets ----------
tpl = wb['Building 01']
sheets = [tpl] + [wb.copy_worksheet(tpl) for _ in A[1:]]
def sumifs(col, code, extra=''):
    return f"=SUMIFS({RS}$F:$F,{RS}$A:$A,$C$2,{RS}$B:$B,{col}$3,{RS}$E:$E,\"{code}\"{extra})"
for a, ws in zip(A, sheets):
    ws.title = a['asset'][:31]; ws['C2'] = a['asset']
    lv = a['levels']
    start = 0 if any(l['z0'] < -0.01 for l in lv) else 1
    if start + len(lv) > 7: raise SystemExit(f"{a['asset']}: more than {7-start} levels - group typical floors or split the asset")
    used = {}
    for i, l in enumerate(lv):
        c = COLS[start + i]; used[l['id']] = c; ws[f'{c}3'] = l['label']
    def setin(addr, v, note):
        ws[addr] = v; ws[addr].font = BLUE; ws[addr].number_format = NUM; ws[addr].comment = Comment(note, 'AAA')
    # row 11: atrium void removed on this floor (new)
    ws['B11'] = '01.3 Atrium Void (removed on other floors)'; ws['B11']._style = copy.copy(ws['B10']._style); ws['B11'].fill = ORANGE
    ws.row_dimensions[11].outlineLevel = ws.row_dimensions[10].outlineLevel
    # row 29: atrium void counted once (new, inside B1 Vertical Circulation)
    ws['B29'] = 'B1.3 Atrium Void (counted on one floor)'; ws['B29']._style = copy.copy(ws['B28']._style); ws['B29'].fill = ORANGE
    ws.row_dimensions[29].outlineLevel = ws.row_dimensions[28].outlineLevel
    for col in COLS + 'J':
        ws[f'{col}29']._style = copy.copy(ws[f'{col}28']._style); ws[f'{col}11']._style = copy.copy(ws[f'{col}10']._style)
    ws['J29'] = '=SUM(C29:I29)'; ws['J11'] = '=SUM(C11:I11)'
    for k, c in used.items():
        s = a['sched'][k]
        setin(f'{c}9', s['skin_gross'], 'Measured to the outer face of the external walls (gross, atrium voids not deducted).')
        if s['internal'] is not None: setin(f'{c}10', s['internal'], 'Measured to the inner face of the external walls.')
        for row, code in ITEM_ROWS.items():
            ws[f'{c}{row}'] = sumifs(c, code); ws[f'{c}{row}'].number_format = NUM; ws[f'{c}{row}'].font = GREEN
        ws[f'{c}29'] = sumifs(c, VOID, f',{RS}$J:$J,"Yes"'); ws[f'{c}29'].font = GREEN; ws[f'{c}29'].fill = ORANGE
        ws[f'{c}11'] = sumifs(c, VOID, f',{RS}$J:$J,"No"'); ws[f'{c}11'].font = GREEN; ws[f'{c}11'].fill = ORANGE
    ws['B82'] = 'H1.2 Covered Terraces'; ws['B82'].fill = ORANGE
    ws['B84'] = 'H1.4 Uncovered Terraces'; ws['B84'].fill = ORANGE; ws['B84'].font = Font(name='Aptos', bold=True)
    ws['J84'] = '=SUM(C84:I84)'; ws['J84'].fill = ORANGE; ws['J84'].number_format = NUM; ws['J84'].font = Font(name='Aptos', bold=True)
    for col in COLS:
        for row in (19, 20, 21):   # A5/A6/A7: H lies outside the skin; atrium voids are neither wall nor floor
            f = ws[f'{col}{row}'].value
            f = f.replace(f'-{col}77', '').replace(f'-{col}66', f'-{col}66-{col}11')
            ws[f'{col}{row}'] = f; ws[f'{col}{row}'].fill = ORANGE
        ws[f'{col}101'] = f'={col}13-{col}15+{col}23+{col}35+{col}44+{col}51+{col}57+{col}66+{col}81+{col}82'; ws[f'{col}101'].fill = ORANGE
    ws['B101'].comment = Comment('AAA BUA = A to G (less A1 notional boundary areas) + H1.1 balconies + H1.2 covered terraces. Atrium voids count once (B1.3) and are removed on the other floors (row 11, taken out of A5/A6/A7).', 'AAA')
    ws['B11'].comment = Comment('Atrium void on floors where it is NOT counted. It is removed from the floor area through A5/A6/A7, so IPMS 1 and BUA AAA drop by this amount.', 'AAA')
    ws['B29'].comment = Comment('Atrium void counted once, on the floor marked Yes in the Room Schedule (normally the lowest floor it opens from).', 'AAA')
for n in ('BS 00',):
    if n in wb.sheetnames: del wb[n]

# ---------- Summary: assets in use only ----------
sm = wb['Summary']
ROWS = range(2, sm.max_row + 1)
grab = lambda col: {r: copy.copy(sm.cell(r, col)._style) for r in ROWS}
st_asset, st_space, st_tot, st_blank = grab(4), grab(9), grab(10), grab(12)
for r in ROWS:
    for c in range(3, 12): sm.cell(r, c).value = None; sm.cell(r, c)._style = copy.copy(st_blank[r])
for k in list(sm.column_dimensions):
    if k not in ('A', 'B'): del sm.column_dimensions[k]
n = len(A); sp, tc = 3 + n, 4 + n; T = L(tc)
COMP = [4, 6, 8, 10, 12, 14, 16, 18]
for i, a in enumerate(A):
    c = 3 + i; C = L(c)
    for r in ROWS: sm.cell(r, c)._style = copy.copy(st_asset[r])
    sm.cell(2, c, a['asset'][:31])
    for r in COMP:
        sm.cell(r, c, f"=IFERROR(INDEX(INDIRECT(\"'\"&{C}$2&\"'!J:J\"),MATCH($B{r},INDIRECT(\"'\"&{C}$2&\"'!B:B\"),0)),0)").fill = ORANGE
    sm.cell(21, c, f'=SUM({C}3:{C}20)'); sm.cell(23, c, f'=IFERROR({C}14/{C}21,0)')
    sm.cell(25, c, f"=IFERROR(INDEX(INDIRECT(\"'\"&{C}$2&\"'!J:J\"),MATCH(\"BUA AAA\",INDIRECT(\"'\"&{C}$2&\"'!B:B\"),0)),0)").fill = ORANGE
    sm.cell(27, c, f'=IFERROR({C}14/{C}25,0)')
    sm.column_dimensions[C].width = max(16.5, len(a['asset']) * 1.3)
for r in ROWS: sm.cell(r, sp)._style = copy.copy(st_space[r]); sm.cell(r, tc)._style = copy.copy(st_tot[r])
sm.column_dimensions[L(sp)].width = 1.7; sm.column_dimensions[T].width = 22
sm.cell(2, tc, 'Total (all assets)' if n > 1 else 'Total')
for r in COMP + [21, 25]: sm.cell(r, tc, f'=SUM(C{r}:{L(2 + n)}{r})')
for r in (23, 27): sm.cell(r, tc, f'=IFERROR({T}14/{T}{r - 2},0)')
order = ['Inclusions', 'Summary'] + [s.title for s in sheets] + ['Room Schedule']
wb._sheets = [wb[t] for t in order]

# ---------- Notes ----------
nt = wb.create_sheet('Notes', 0)
lines = [('Area calculation – ' + ', '.join(a['asset'] for a in A), True)]
for a in A:
    lines.append((f"{a['asset']}: {a['meta'].get('sources', '')}", False))
lines += [('Measured by AAA from the drawing linework. External skin = outer face of external walls; rooms = inner faces.', False), ('', False),
 ('Measurement rules applied', True),
 ('AAA BUA = external skin area (A2 to G) + balconies (H1.1) + covered terraces (H1.2). Notional boundary areas (A1), uncovered terraces (H1.4) and sheltered areas (H1.3) are excluded.', False),
 ('IPMS 1 = all of the above + uncovered terraces, sheltered areas and notional boundary areas. Plant roofs are not counted by either.', False),
 ('Terraces under an open-slat pergola are treated as uncovered terraces.', False),
 ('Atrium voids: counted once, as B1.3 Atrium Void on the floor marked Yes in the Room Schedule; removed from every other floor it passes through (row 11, taken out of A5/A6/A7, so IPMS 1 and BUA AAA drop by that amount).', False),
 ('', False), ('Changes to the template (orange cells)', True),
 ('Asset sheets row 11 (new): 01.3 Atrium Void removed on other floors. Row 29 (new): B1.3 Atrium Void, inside B1 Vertical Circulation.', False),
 ('Asset sheets rows 19–21: A5/A6/A7 no longer subtract group H (balconies and terraces lie outside the external skin) and now subtract row 11.', False),
 ('Asset sheets row 82 renamed H1.2 Covered Terraces; row 84 added as H1.4 Uncovered Terraces.', False),
 ('Asset sheets row 101 (BUA AAA): A to G less A1, plus H1.1 and H1.2, for every column.', False),
 ('Summary: only the assets in use are kept. Rows 4–18 and 25 read each asset sheet\'s Total column with INDEX/MATCH.', False),
 ('', False), ('Check against the report', True)]
for a in A:
    t = a['totals']
    rb = f"report BUA {a['report_bua']:.2f} m², " if a.get('report_bua') else ''
    lines.append((f"{a['asset']}: {rb}external skin {t['skin']:.2f} m², AAA BUA {t['bua']:.2f} m², IPMS 1 {t['ipms']:.2f} m².", False))
for i, (t, b) in enumerate(lines, 1):
    c = nt.cell(row=i, column=1, value=t); c.font = Font(name='Aptos', bold=b, size=13 if i == 1 else 11)
    c.alignment = Alignment(wrap_text=True, vertical='top')
nt.column_dimensions['A'].width = 120
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('saved', OUT, [a['asset'] for a in A], 'room rows', last - 2)
