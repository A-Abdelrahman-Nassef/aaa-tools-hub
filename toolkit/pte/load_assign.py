"""load_assign.py - which BOQ goes where, for the hub's 06-06 loading step.

assign(master, register_rows, boq_paths, divmap=None) -> dict with
  loads    {tab: {block index: [{"num", "rows", "package", "file", "label"}]}}  for load_master.plan
  boqs     one entry per BOQ: file, asset, package, tab, block, items, headings, qty, divisions (label -> number)
  divmap   {"<package>|<BOQ division label>": {"num": n or None, "why": reason}}  (the review screen edits these)
  problems things that stop the build; notes things worth reading
"""
import re
from pathlib import Path

import boq_items as B
import load_master as M
import names as N


def _key(package, label):
    return f'{package}|{label}'


def assign(master, register_rows, boq_paths, divmap=None):
    divmap = dict(divmap or {})
    problems, notes = [], []
    reg = {N.E.clean(r['file']).upper(): r for r in register_rows}
    boqs, wanted_tabs = [], []
    for p in boq_paths:
        stem = N.E.clean(Path(p).name)
        r = reg.get(stem.upper())
        if not r:
            problems.append(f'{Path(p).name}: not in the register (run 06-03 again with this BOQ)')
            continue
        boqs.append(dict(file=Path(p).name, path=str(p), asset=r['asset'], package=r['package']))
        if r['asset'] not in wanted_tabs:
            wanted_tabs.append(r['asset'])
    try:
        L = M.detect(master, wanted_tabs)
    except M.LoadError as e:
        return dict(problems=problems + [str(e)], notes=notes, boqs=boqs, divmap=divmap, loads={}, layout=None)
    for tab in wanted_tabs:
        if not L[tab]['unloaded']:
            problems.append(f'{tab}: already has items loaded; start from the master built by 06-04')
    loads = {}
    seen = set()
    for b in boqs:
        Lt = L[b['asset']]
        names = [x['name'].lower() for x in Lt['blocks']]
        if Lt['blocks'][0]['header'] is None:
            bi = 0
        elif b['package'].lower() in names:
            bi = names.index(b['package'].lower())
        else:
            problems.append(f"{b['file']}: {b['asset']} has no package block \"{b['package']}\" "
                            f"(blocks: {', '.join(x['name'] for x in Lt['blocks'])})")
            continue
        if (b['asset'], bi) in seen and Lt['blocks'][0]['header'] is not None:
            problems.append(f"{b['file']}: a second BOQ for {b['asset']} / {b['package']}")
        seen.add((b['asset'], bi))
        x = B.read_boq(b['path'])
        b.update(tab=b['asset'], block=bi, items=x['counts']['items'], headings=x['counts']['headings'],
                 qty=x['counts']['qty'], title=x['title'], notes=x['notes'], divisions={})
        master_nums = {d['num'] for d in Lt['blocks'][bi]['divs']}
        for d in x['divisions']:
            k = _key(b['package'], d['label'])
            if k not in divmap:
                n, why = B.map_division(d['num'], d['name'], b['package'])
                divmap[k] = dict(num=n, why=why)
            n = divmap[k]['num']
            b['divisions'][d['label']] = n
            items = sum(1 for r in d['rows'] if r['type'] == 'item')
            if n is None:
                problems.append(f"{b['file']}: {d['label']} ({items} items) has no master division; choose one")
                continue
            if n not in master_nums:
                problems.append(f"{b['file']}: Division {n:02d} is not in the master")
                continue
            loads.setdefault(b['asset'], {}).setdefault(bi, []).append(
                dict(num=n, rows=d['rows'], package=b['package'], file=b['file'], label=d['label']))
        for n_ in x['notes']:
            notes.append(f"{b['file']}: {n_}")
    for tab, by_block in loads.items():               # two BOQ parts in one master division: a code heading each
        for bi, parts in by_block.items():
            nums = [p['num'] for p in parts]
            for p in parts:
                if nums.count(p['num']) > 1:
                    p['rows'] = [dict(type='sec', desc=f"{p['package']} - {p['label']}")] + p['rows']
    missing = []
    for tab in wanted_tabs:
        for bi, blk in enumerate(L[tab]['blocks']):
            if not any(b.get('tab') == tab and b.get('block') == bi for b in boqs):
                missing.append(f"{tab} / {blk['name'] or 'single block'}")
    if missing:
        notes.append('No BOQ for: ' + ', '.join(missing) + ' (left as template placeholders).')
    return dict(problems=problems, notes=notes, boqs=boqs, divmap=divmap, loads=loads, layout=L)
