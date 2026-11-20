"""asset_lib.py - geometry -> area schedule for one asset (AAA BUA and IPMS 1).

An asset spec is a Python dict (see sample_villa_example.py):
  asset   : sheet / display name, e.g. 'Sample Villa'
  meta    : {eyebrow, note, scope, report_label, sources, group{}, desc{}, open_labels{}, bua_text, ipms_text}
  levels  : [{id:'GF', name:'Ground floor', label:'Ground', z0:0.30, h:3.25, roof:False}, ...]  bottom -> top
            label = the column header used in the workbook (row 3) and the Room Schedule 'Level' column
  skin    : {level_id: polygon}  outer face of the external walls (gross, voids NOT deducted)
  internal: {level_id: polygon}  inner face of the external walls (omit a level to leave it blank)
  rooms   : {level_id: [(code, name, ref, polygon), ...]}   codes from the AAA template (B1.1, C1.4, D1.1, E1.1, F1.1 ...)
            Atrium void: code 'B1.3', name, ref, polygon, and a 5th item True on the ONE floor where it is counted,
            False on every other floor it passes through (there it is removed from the floor area).
  opens   : [(key, name, level_id, polygon, note[, extra dict]), ...] outside the external skin
            key: bal H1.1 | cov H1.2 | shel H1.3 | unc H1.4 | perg H1.4 (under open pergola) | nb A1 | plant (not counted)
            extra (optional): {'ztop': 3.42, 'grade': True}  -> model only
  code_labels: {code: label} for the model ledger
  report_bua : optional figure from the report, for the check
Polygons: [(x, y), ...] in metres, X east, Y south (plan frame), any orientation.
"""
import json

def R(x0, y0, x1, y1): return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]

def area(p):
    s = 0
    for i in range(len(p)):
        a, b = p[i], p[(i + 1) % len(p)]
        s += a[0] * b[1] - b[0] * a[1]
    return abs(s) / 2

IN_BUA = {'cov': True, 'bal': True, 'unc': False, 'perg': False, 'shel': False, 'nb': False, 'plant': False}
IN_IPMS = {'cov': True, 'bal': True, 'unc': True, 'perg': True, 'shel': True, 'nb': True, 'plant': False}
OPEN_CODE = {'bal': 'H1.1', 'cov': 'H1.2', 'shel': 'H1.3', 'unc': 'H1.4', 'perg': 'H1.4', 'nb': 'A1', 'plant': 'Plant roof'}
VOID = 'B1.3'

def _rooms(spec, k):
    out = []
    for r in spec['rooms'].get(k, []):
        code, name, ref, poly = r[:4]
        counted = r[4] if len(r) > 4 else True
        out.append(dict(code=code, name=name, ref=ref, poly=[list(p) for p in poly], area=round(area(poly), 2),
                        void=code == VOID, counted=bool(counted)))
    return out

def build(spec):
    """Return the JSON-able data dict used by the workbook builder, the 3D model and the deck."""
    lv_out, rooms, opens, sched = [], {}, [], {}
    for L in spec['levels']:
        k = L['id']; lv_out.append(dict(L))
        rs = _rooms(spec, k)
        rooms[k] = [r for r in rs if not (r['void'] and not r['counted'])]
        removed = [r for r in rs if r['void'] and not r['counted']]
        sk = round(area(spec['skin'][k]), 2)
        it = round(area(spec['internal'][k]), 2) if spec.get('internal', {}).get(k) else None
        void_rm = round(sum(r['area'] for r in removed), 2)
        void_all = round(sum(r['area'] for r in rs if r['void']), 2)
        rsum = sum(r['area'] for r in rs if not r['void'])
        comp = {}
        for r in rs:
            if r['void'] and not r['counted']: continue
            comp[r['code']] = round(comp.get(r['code'], 0) + r['area'], 2)
        a2 = round(sk - it, 2) if it is not None else 0.0
        a6 = round((it if it is not None else sk) - rsum - void_all, 2)
        ops = []
        for o in spec['opens']:
            key, name, lvl, poly, note = o[:5]
            extra = o[5] if len(o) > 5 else {}
            if lvl != k: continue
            d = dict(key=key, name=name, level=k, poly=[list(p) for p in poly], area=round(area(poly), 2), note=note, **extra)
            opens.append(d); ops.append([key, name, d['area']])
        for r in removed:
            opens.append(dict(key='void', name=r['name'] + ' (removed)', level=k, poly=r['poly'], area=r['area'],
                              note='Atrium void, counted on one floor only; removed from this floor.', grade=False))
            ops.append(['void', r['name'] + ' (removed)', r['area']])
        net = round(sk - void_rm, 2)
        bua = net + sum(a for key, n, a in ops if IN_BUA.get(key))
        ipms = net + sum(a for key, n, a in ops if IN_IPMS.get(key))
        sched[k] = dict(skin=net, skin_gross=sk, internal=it, void_removed=void_rm, a2=a2, a6=a6, comp=comp, opens=ops,
                        bua=round(bua, 2), ipms=round(ipms, 2))
    tot = {t: round(sum(sched[k][t] for k in sched), 2) for t in ('skin', 'bua', 'ipms', 'void_removed')}
    return dict(asset=spec['asset'], meta=spec.get('meta', {}), levels=lv_out,
                skin={k: [list(p) for p in v] for k, v in spec['skin'].items()},
                internal={k: [list(p) for p in v] for k, v in spec.get('internal', {}).items()},
                rooms=rooms, opens=opens, sched=sched, totals=tot, code_labels=spec.get('code_labels', {}),
                report_bua=spec.get('report_bua'))

def dump(spec, path):
    d = build(spec)
    json.dump(d, open(path, 'w'), indent=1)
    print(d['asset'], d['totals'], 'report BUA', d['report_bua'])
    for k, s in d['sched'].items():
        print(f"  {k}: skin {s['skin']} (gross {s['skin_gross']}, void removed {s['void_removed']})  A2 {s['a2']}  A6 {s['a6']}  BUA {s['bua']}  IPMS {s['ipms']}  comp {s['comp']}")
    return d
