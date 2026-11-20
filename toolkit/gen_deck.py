"""gen_deck.py - the collected Annex D deck (Slides artifact type) for one or more assets.
usage: python gen_deck.py deck_config.json root_dir
Writes root_dir/project/deck.json and root_dir/project/slides/*.html.
deck_config.json:
 {"title": "Annex D: <Project> Built-Up Area", "client": "Sample Developer", "project": "SAMPLE Sample City",
  "report": "PRE-TENDER ESTIMATE REPORT – REV00 – ANNEX D", "date": "23rd OCTOBER 2026",
  "header": "Sample Developer – SAMPLE Sample City – ANNEX D – BUILT-UP AREA", "annex": "D",
  "logo": "<asset url>", "chevron": "<asset url>",
  "assets": [{"json": "crown.json", "images": {"comp": url, "compare": url, "ipms": url, "plan_GF": url, ...}}],
  "points": ["point to confirm", ...],
  "explore": {"ui_image": "<asset url of shot_ui.py jpg>", "zips": ["AAA-Area-Model-Crown-Villa.zip"]},   (optional; adds "Explore the model yourself")
  "rics": "<asset url of the white Regulated by RICS badge>"}   (optional; back cover badge)
The deck ends: ... schedule, points to confirm, Explore the model yourself, back cover.
Image urls are the asset-store urls returned by the Artifact upload (publish with asset:true)."""
import sys, json, os, html as H
cfg = json.load(open(sys.argv[1])); root = sys.argv[2]
base = os.path.dirname(os.path.abspath(sys.argv[1]))
AX = cfg.get('annex', 'D')
A = [(json.load(open(os.path.join(base, a['json']))), a.get('images', {})) for a in cfg['assets']]
os.makedirs(f'{root}/project/slides', exist_ok=True)
CH = cfg['chevron']; E = H.escape
f2 = lambda v: f'{v:,.2f}'
def frame(pg):
    return f'''<div style="position:absolute;left:64px;top:36px;width:1560px;height:44px;background:linear-gradient(90deg, #14264A, #2C4D80);display:flex;align-items:center;padding:0 18px"><p style="font-size:24px;font-weight:700;color:#FFFFFF;white-space:nowrap">| {E(cfg['header'])}</p></div>
<div style="position:absolute;left:1640px;top:32px;width:3px;height:40px;background:#14264A"></div>
<p style="position:absolute;left:1650px;top:28px;width:220px;font-size:24px;line-height:1.25;text-align:center;color:#1E3A66">AAA<br>CONSULTANCY</p>
<div style="position:absolute;left:64px;bottom:36px;width:1792px;height:40px;background:linear-gradient(90deg, #14264A, #2C4D80);display:flex;align-items:center;justify-content:end;padding:0 28px"><p style="font-size:24px;font-weight:700;color:#FFFFFF">PAGE | {AX}-{pg}</p></div>'''
SL, ORDER = {}, []
def add(sid, title, body, aside, gap=36):
    pg = len(ORDER)  # cover is page 0
    num = f'{AX}.{pg}'
    SL[sid] = f'''<section id="{sid}" data-transition="fade" style="background:#FFFFFF;color:#1B2233;font-family:'DM Sans', Arial, sans-serif;padding:124px 128px 160px;display:flex;flex-direction:column;gap:{gap}px">
<img src="{CH}" alt="" style="position:absolute;right:0;top:0;width:391px;height:1080px;object-fit:cover;filter:brightness(0);opacity:0.05">
<div style="display:flex;flex-direction:column;gap:6px">
<p style="font-size:24px;font-weight:700;letter-spacing:1px;color:#1E3A66">{AX}   BASIS FOR MEASUREMENT OF BUILT-UP AREA</p>
<h2 style="font-size:52px;font-weight:700;line-height:1.1;color:#1B2233">{num}   {E(title)}</h2>
</div>
{body}
{frame(pg)}
<aside>{E(aside)}</aside>
</section>
'''
    ORDER.append(sid)
P = lambda t, s=26: f'<p style="font-size:{s}px;line-height:1.45;color:#4A5568">{t}</p>'
def panel(src, alt, w=680, h=664):
    if not src: return f'<div style="background:#F2F4F7;width:{w}px;height:{h}px"></div>'
    return f'<div style="background:#F2F4F7;width:{w}px;height:{h}px;display:flex;align-items:center;justify-content:center"><img src="{src}" alt="{E(alt)}" style="width:{w-20}px;height:{h-24}px;object-fit:contain"></div>'
def table(head, rows, w, fs=22, al=None):
    al = al or 'l' + 'r' * (len(head) - 1)   # per column: l, r or c
    A_ = {'l': 'text-align:left', 'r': 'text-align:right', 'c': 'text-align:center'}
    h = ''.join(f'<th style="color:#FFFFFF;{A_[al[i]]}">{x}</th>' for i, x in enumerate(head))
    b = ''
    for bg, *cells in rows:
        st = f' style="background:{bg}"' if bg else ''
        b += f'<tr{st}>' + ''.join(f'<td style="{A_[al[i]]}">{c}</td>' for i, c in enumerate(cells)) + '</tr>'
    return f'<table style="width:{w}px;font-size:{fs}px;color:#1B2233;border:1px solid #1E3A66"><tr style="background:#1E3A66">{h}</tr>{b}</table>'
def eq(rows, label, total, color):
    r = ''.join(f'<div style="display:flex;justify-content:space-between;align-items:baseline;border-bottom:1px solid #D5DAE3;padding:12px 0"><p style="font-size:26px;color:#2B3446">{a}</p><p style="font-size:32px;font-weight:700;color:#1B2233">{b}</p></div>' for a, b in rows)
    return f'<div style="background:#F2F4F7;padding:28px 36px;display:flex;flex-direction:column">{r}<div style="display:flex;justify-content:space-between;align-items:baseline;padding:18px 0 4px"><p style="font-size:30px;font-weight:700;color:{color}">{label}</p><p style="font-size:50px;font-weight:700;color:{color}">{total}</p></div></div>'
OPEN_TXT = {'bal': 'H1.1 Balconies', 'cov': 'H1.2 Covered terrace', 'shel': 'H1.3 Sheltered area', 'unc': 'H1.4 Uncovered terrace',
            'perg': 'H1.4 Terrace under pergola', 'nb': 'A1 Notional boundary area', 'plant': 'Plant roof, not counted', 'void': 'B1.3 Atrium void, removed'}
BUA_K, IPMS_K = {'bal', 'cov'}, {'bal', 'cov', 'unc', 'perg', 'shel', 'nb'}
CODE_LABEL = {'B1.1':'Stairs','B1.2':'Elevators','B1.3':'Atrium void','B2.1':'Technical shafts','B2.2':'Voids','C1.1':'Mechanical rooms','C1.2':'Electrical rooms','C1.3':'Garbage rooms','C1.4':'Storage','D1.1':'Toilets and bathrooms','D1.2':'Cleaners’ cupboards','E1.1':'Corridors and lobbies','F1.1':'Primary rooms','F1.2':'Offices','F1.3':'Retail','F1.4':'Clinics','G1.1':'Seating areas','G1.2':'Child-minding','G2.1':'Car parking'}
GRP = {'A': 'Walls and columns', 'B': 'Vertical penetrations', 'C': 'Technical', 'D': 'Sanitary', 'E': 'Circulation', 'F': 'Primary', 'G': 'Secondary'}
T = {k: round(sum(a['totals'][k] for a, _ in A), 2) for k in ('skin', 'bua', 'ipms')}
multi = len(A) > 1

# ---- cover
SL['cover'] = f'''<section id="cover" data-transition="fade" style="background:linear-gradient(135deg, #0F1157 0%, #0B0A40 100%);color:#F4F5FB;font-family:'DM Sans', Arial, sans-serif;padding:128px;display:flex;flex-direction:column">
<img src="{CH}" alt="" style="position:absolute;right:0;top:0;width:391px;height:1080px;object-fit:cover">
<img src="{cfg['logo']}" alt="AAA Consultancy" style="position:absolute;left:120px;top:104px;width:370px;height:126px;object-fit:contain">
<p style="position:absolute;left:128px;top:330px;width:1200px;font-size:32px;font-weight:700;letter-spacing:6px;color:#9FB0E6">ANNEX {AX}</p>
<h1 style="position:absolute;left:124px;top:384px;width:1260px;font-size:96px;font-weight:700;line-height:1.08;color:#F4F5FB">BASIS FOR MEASUREMENT OF BUILT-UP AREA</h1>
<div style="position:absolute;left:128px;top:704px;width:1200px;display:flex;flex-direction:column;gap:14px">
<p style="font-size:28px;color:#F4F5FB">CLIENT: {E(cfg['client'])}</p><p style="font-size:28px;color:#F4F5FB">PROJECT: {E(cfg['project'])}</p>
<p style="font-size:28px;color:#F4F5FB">TITLE: {E(cfg['report'])}</p><p style="font-size:28px;color:#F4F5FB">DATE: {E(cfg['date'])}</p>
</div>
<aside>Annex {AX}: how the built-up area of each asset was measured from the drawings, and how it compares with IPMS 1.</aside>
</section>
'''
ORDER.append('cover')
# ---- basis
src = ''.join(f'<li><b>{E(a["asset"])}:</b> {E(a["meta"].get("sources", ""))}</li>' for a, _ in A)
cards = ''.join(f'<div style="background:#F2F4F7;border-left:10px solid {c};padding:24px 32px;display:flex;flex-direction:column;gap:6px"><p style="font-size:24px;font-weight:700;letter-spacing:1px;color:#5A6577">{l}</p><p style="font-size:56px;font-weight:700;line-height:1.05;color:#1B2233">{f2(v)} m²</p></div>'
                for l, v, c in (('EXTERNAL SKIN AREA', T['skin'], '#2A8562'), ('AAA BUA', T['bua'], '#3F6FBF'), ('IPMS 1', T['ipms'], '#D98A1C')))
add('basis', 'Basis of measurement', f'''<div style="flex:1;display:flex;gap:64px;align-items:start">
<ul style="flex:1;font-size:26px;line-height:1.45;color:#2B3446;display:flex;flex-direction:column;gap:18px">
<li>This annex sets out how the built-up area of {"each asset" if multi else E(A[0][0]["asset"])} was measured, and how it compares with IPMS 1.</li>{src}
<li>The measurement follows AAA Consultancy’s built-up area (BUA) methodology, aligned in principle with IPMS 1 and adjusted to normal practice in Egypt. Enclosed areas to the outer face of the external walls; rooms to the inner face of their walls.</li></ul>
<div style="width:560px;display:flex;flex-direction:column;gap:20px">{cards}</div></div>''', 'Totals ' + ('for all assets.' if multi else 'for the asset.'))
# ---- inclusions
inc = [('', '1', 'Areas from external façade skin', 'Included', 'Included'), ('', 'A–G', 'Walls, columns, vertical penetrations, technical, sanitary, circulation, primary and secondary areas', 'Included', 'Included'),
       ('', 'B1.3', 'Atrium voids: once, on one floor; removed on the other floors', 'Once', 'Once'),
       ('', 'H1.1–H1.2', 'Balconies and covered terraces', 'Included', 'Included'),
       ('#FBEFD9', 'H1.4', 'Uncovered terraces, and terraces under an open pergola', '<span style="color:#8A4B00">Excluded</span>', 'Included'),
       ('#FBEFD9', 'H1.3 / A1', 'Sheltered areas, notional boundary areas', '<span style="color:#8A4B00">Excluded</span>', 'Included'),
       ('', '–', 'Plant roofs, maintenance access only', 'Not counted', 'Not counted')]
add('inclusions', 'Inclusions and exclusions', '<div style="flex:1;display:flex;flex-direction:column;gap:20px">' +
    table(['#', 'Component', 'AAA BUA', 'IPMS 1'], [(bg, a, b, c, d) for bg, a, b, c, d in inc], 1664, 24, 'llcc') +
    P('The shaded rows are the only areas where AAA BUA and IPMS 1 differ.') + '</div>', 'Rules applied to every asset.')
# ---- per asset
for ai, (a, im) in enumerate(A):
    nm = a['asset']; pre = f'a{ai+1}'
    lvrows = [('', E(l['name']), f"{l['z0']:+.2f}", f2(a['sched'][l['id']]['skin'])) for l in reversed(a['levels'])]
    lvrows.append(('#E9EEF5', '<b>Total</b>', '', f"<b>{f2(a['totals']['skin'])}</b>"))
    add(f'{pre}-over', f'{nm} in 3D', f'''<div style="flex:1;display:flex;gap:56px;align-items:start">{panel(im.get('comp'), f'3D view of {nm}, coloured by component', 640, 656)}
<div style="flex:1;display:flex;flex-direction:column;gap:24px">{P(f'{len(a["levels"])} levels, shown separated. Coloured by area component A to H.', 28)}
{table(['Level', 'FFL', 'Ext. skin (m²)'], lvrows, 968, 24)}</div></div>''', f'{nm}: levels and external skin area.')
    for l in a['levels']:
        k = l['id']; s = a['sched'][k]
        rows = [('', 'A2', 'External walls', f2(s['a2']), 'Yes', 'Yes'), ('', 'A6', 'Internal walls, shafts', f2(s['a6']), 'Yes', 'Yes')]
        for code in sorted(s['comp']):
            lab = (a.get('code_labels') or {}).get(code, CODE_LABEL.get(code, GRP.get(code[0], code)))
            rows.append(('', code, E(lab), f2(s['comp'][code]), 'Yes', 'Yes'))
        rows.append(('#E9EEF5', '', '<b>External skin area</b>' + (' (net of void)' if s['void_removed'] else ''), f"<b>{f2(s['skin'])}</b>", 'Yes', 'Yes'))
        for key, n, ar in s['opens']:
            bua, ip = key in BUA_K, key in IPMS_K
            bg = '#FBEFD9' if (ip and not bua) else ''
            rows.append((bg, OPEN_TXT[key].split(' ')[0], E(n), f2(ar), 'Yes' if bua else '<span style="color:#8A4B00">No</span>', 'Yes' if ip else 'No'))
        rows += [('#DCE7F6', '', '<b>AAA BUA</b>', f"<b>{f2(s['bua'])}</b>", '', ''), ('#FBE3BD', '', '<b>IPMS 1</b>', f"<b>{f2(s['ipms'])}</b>", '', '')]
        fs = 22 if len(rows) <= 14 else 18
        add(f'{pre}-{k.lower()}', f"{nm}: {l['name']} (FFL {l['z0']:+.2f})", f'''<div style="flex:1;display:flex;gap:48px;align-items:start">{panel(im.get(f'plan_{k}'), f"{nm} {l['name']} plan coloured by component", 520, 664)}
<div style="flex:1">{table(['#', 'Component', 'm²', 'AAA BUA', 'IPMS 1'], rows, 1096, fs, 'llrcc')}</div></div>''', f"{nm} {l['name']}: components and open areas.", gap=32)
    t = a['totals']
    add_b = round(t['bua'] - t['skin'], 2); add_i = round(t['ipms'] - t['bua'], 2)
    add(f'{pre}-build', f'{nm}: from skin to BUA and IPMS 1', f'''<div style="flex:1;display:flex;gap:56px;align-items:start">{panel(im.get('compare'), f'{nm}: inside the skin green, added by BUA blue, added by IPMS 1 only amber')}
<div style="flex:1;display:flex;flex-direction:column;gap:24px">
{eq([('External skin area' + (' (net of atrium voids)' if t['void_removed'] else ''), f2(t['skin']) + ' m²'), ('+ Balconies and covered terraces', f2(add_b) + ' m²')], 'AAA BUA', f2(t['bua']) + ' m²', '#2D5AA6')}
{eq([('AAA BUA', f2(t['bua']) + ' m²'), ('+ Uncovered, pergola, sheltered and notional areas', f2(add_i) + ' m²')], 'IPMS 1', f2(t['ipms']) + ' m²', '#B26A00')}
</div></div>''', f'{nm}: AAA BUA {f2(t["bua"])} m², IPMS 1 {f2(t["ipms"])} m².', gap=32)
# ---- schedule across assets
rows = []
for a, _ in A:
    t = a['totals']; rb = a.get('report_bua')
    rows.append(('', E(a['asset']), f2(t['skin']), f2(t['bua']), f2(t['ipms']), f2(rb) if rb else '–',
                  (f"{t['bua'] - rb:+,.2f} ({(t['bua'] / rb - 1) * 100:+.1f}%)" if rb else '–')))
if multi: rows.append(('#E9EEF5', '<b>Total</b>', f"<b>{f2(T['skin'])}</b>", f"<b>{f2(T['bua'])}</b>", f"<b>{f2(T['ipms'])}</b>", '', ''))
add('schedule', 'Area schedule', '<div style="flex:1;display:flex;flex-direction:column;gap:20px">' +
    table(['Asset', 'Ext. skin (m²)', 'AAA BUA (m²)', 'IPMS 1 (m²)', 'Report BUA (m²)', 'AAA BUA vs report'], rows, 1664, 24) +
    P('Room-by-room areas are in the AAA Areas Calculations workbook, Room Schedule sheet.') + '</div>', 'All assets, with the check against the report figures.')
pts = ''.join(f'<li>{E(p)}</li>' for p in cfg.get('points', []))
if pts:
    add('check', 'Points to confirm', f'<ul style="font-size:26px;line-height:1.45;color:#2B3446;display:flex;flex-direction:column;gap:14px">{pts}</ul>', 'Assumptions to confirm with the designer before issue.')
ex = cfg.get('explore')
if ex:
    zips = ex['zips']
    zt = (f'<b>{E(zips[0])}</b>, sent with this report.' if len(zips) == 1 else
          'the file for each asset, sent with this report: ' + ', '.join(f'<b>{E(z)}</b>' for z in zips) + '.')
    lv_ids = ' / '.join(['All levels'] + [l['id'] for l in A[0][0]['levels']]) if not multi else 'Level buttons'
    step = lambda n, t: f'<div style="display:flex;gap:18px"><p style="font-size:22px;font-weight:700;color:#B26A00;width:36px">{n}</p><p style="flex:1;font-size:22px;line-height:1.35;color:#2B3446">{t}</p></div>'
    ctrl = [('Drag', 'Rotate'), ('Right-drag', 'Pan (two fingers on a tablet)'), ('Scroll', 'Zoom (pinch on a tablet)'),
            ('Compare … IPMS 1', 'Colour by measure: skin, AAA BUA, IPMS 1 or components'), (lv_ids, 'Show one floor, with its areas'),
            ('Separate floors, Plan view', 'Stacked or separated; 3D or plan'), ('Measurement lines', 'External skin outline on or off'),
            ('Click an area or row', 'Highlight it and show its area')]
    rows = ''.join(f'<tr><td style="font-weight:700;width:44%">{E(a)}</td><td>{E(b)}</td></tr>' for a, b in ctrl)
    add('explore', 'Explore the model yourself', f'''<div style="flex:1;display:flex;gap:52px;align-items:start">
<div style="width:980px;display:flex;flex-direction:column;gap:14px">
<div style="background:#F2F4F7;border:1px solid #D5DAE3;width:980px;height:601px;display:flex;align-items:center;justify-content:center"><img src="{ex['ui_image']}" alt="The area model: one floor in 3D with every space labelled by name, code and area, the component ledger on the right and the measure buttons along the top" style="width:978px;height:599px;object-fit:contain"></div>
{P('The model is for illustration of the area measurement method; areas are taken from the Areas Calculations workbook.', 22)}</div>
<div style="flex:1;display:flex;flex-direction:column;gap:22px">
<h3 style="font-size:30px;font-weight:700;color:#1E3A66">Open the file</h3>
<div style="display:flex;flex-direction:column;gap:14px">{step('01', 'Unzip ' + zt)}{step('02', 'Double-click the .html file. It opens in Chrome, Edge or Safari; no internet, account or installation is needed.')}{step('03', 'Pick a floor or an area measure, or click any part of the model or any row of the table.')}</div>
<h3 style="font-size:30px;font-weight:700;color:#1E3A66">Controls</h3>
<table style="width:632px;font-size:20px;color:#1B2233">{rows}</table>
</div></div>''', 'How to open and use the offline model sent with this report.', gap=32)
SL['back'] = f'''<section id="back" data-transition="fade" style="background:linear-gradient(135deg, #0F1157 0%, #0B0A40 100%);color:#F4F5FB;font-family:'DM Sans', Arial, sans-serif;padding:128px;display:flex;flex-direction:column">
<img src="{CH}" alt="" style="position:absolute;right:0;top:0;width:391px;height:1080px;object-fit:cover">
<img src="{cfg['logo']}" alt="AAA Consultancy" style="position:absolute;left:120px;top:360px;width:470px;height:160px;object-fit:contain;object-position:left center">
<p style="position:absolute;left:128px;top:560px;width:1100px;font-size:30px;line-height:1.45;color:#C9D2EE">This annex accompanies the report it is issued with and should be read together with it. Areas are measured from the drawings listed in {AX}.1; the 3D model illustrates the measurement method.</p>
{f'<img src="{cfg["rics"]}" alt="Regulated by RICS" style="position:absolute;left:128px;top:880px;width:240px;height:96px;object-fit:contain;object-position:left center">' if cfg.get('rics') else ''}
<aside>Back cover.</aside>
</section>
'''
ORDER.append('back')
for sid, h in SL.items(): open(f'{root}/project/slides/{sid}.html', 'w').write(h)
deck = {"v": 4, "lists": "css", "title": cfg['title'], "cover": "cover", "order": ORDER,
        "faces": {"dm-sans": {"family": "DM Sans", "href": "https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap"}}, "designSystems": []}
json.dump(deck, open(f'{root}/project/deck.json', 'w'))
print(len(ORDER), 'slides:', ORDER)
