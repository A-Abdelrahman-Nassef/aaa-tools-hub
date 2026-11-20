"""run_step.py - run one script-only PTE Phase 1 step for a hub job (no Claude).

    python run_step.py <module> <job dir> [--check]

<job dir>/job.json   {"inputs": [{"name", "kind", "path"}...], "options": {...}}
<job dir>/review.json (asset-names build only) {"rows": [{"file", "asset", "package"}...]}

Writes <job dir>/out/manifest.json (or check/manifest.json with --check):
  {"stage", "module", "summary", "columns", "rows", "checks": [{"check", "ok", "detail"}], "notes",
   "publish": [{"file": <path in out>, "kind": <project file kind>}]}
Only files listed in "publish" go back to the project's files, and only when the step passed its checks.
Exit code 1 when nothing usable came out.
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

EXCEL = ('.xlsx', '.xlsm', '.xltx', '.xltm')


def stage_inputs(job, src):
    """Copy the chosen project files into <job>/in; a previous register is renamed so it is never read as a BOQ."""
    shutil.rmtree(src, ignore_errors=True); src.mkdir(parents=True)
    notes, staged = [], []
    for i, f in enumerate(job['inputs'], 1):
        name = f['name']
        if f['kind'] == 'register':
            name = f'File_Names - previous {i}{Path(name).suffix}'
        dst = src / name
        if dst.exists():
            notes.append(f'{name} was chosen twice ({f["kind"]} and an earlier kind): used the first one.')
            continue
        shutil.copy2(f['path'], dst)
        staged.append(dict(f, staged=str(dst)))
    return staged, notes


# ---------------- 06-01 copy BOQs to unlocked files ----------------
def step_unlock(job, d, out):
    import unlock_excel as U
    staged, notes = stage_inputs(job, d / 'in')
    rows, checks, publish = [], [], []
    tmp = tempfile.mkdtemp()
    files, converted = [], set()
    for s in staged:
        p = s['staged']
        if p.lower().endswith('.zip'):
            files += U.collect([p], tmp)
        elif p.lower().endswith('.xls'):        # legacy .xls: Excel saves it as .xlsx first, then it is unlocked
            import excel
            try:
                x = excel.xls_to_xlsx(p, os.path.join(tmp, 'converted'))
                files.append(x); converted.add(x)
            except Exception as e:
                rows.append([s['name'], 'could not convert from .xls', '', 'no'])
                checks.append(dict(check=f"{s['name']} opens in Excel", ok=False,
                                   detail='needs a password to open, or the file is damaged: ask the designer for an '
                                          f'unprotected copy ({str(e).splitlines()[0][:120]})'))
        elif p.lower().endswith(EXCEL):
            files.append(p)
        else:
            notes.append(f"{s['name']}: not an Excel file, ignored.")
    dest = out / 'Unlocked BOQs'
    dest.mkdir(parents=True, exist_ok=True)
    unlocked = already = 0
    for f in files:
        name = os.path.basename(f)
        n = U.unlock(f, str(dest / name))
        if n is None:
            rows.append([name, 'password to open: cannot unlock', '', 'no'])
            checks.append(dict(check=f'{name} opens', ok=False, detail='Encrypted with a password to open. Ask the designer for it.'))
            (dest / name).unlink(missing_ok=True)
            continue
        try:
            same = U.verify(f, str(dest / name))
        except Exception as e:
            same = False; notes.append(f'{name}: check error {e}')
        status = 'unlocked' if n else 'had no protection (copied as is)'
        if f in converted:
            status = 'converted from .xls by Excel, ' + ('then unlocked' if n else 'had no protection')
        unlocked += bool(n); already += not n
        rows.append([name, status, n, 'yes' if same else 'NO'])
        checks.append(dict(check=f'{name} identical to the original', ok=bool(same),
                           detail='values, styles, merges, images; no sheet still protected'))
        if same:
            publish.append(dict(file=f'Unlocked BOQs/{name}', kind='boq-unlocked'))
        else:
            (dest / name).unlink(missing_ok=True)
    if len(publish) > 1:
        shutil.make_archive(str(out / 'Unlocked BOQs'), 'zip', out, 'Unlocked BOQs')
    summary = f'{len(publish)} of {len(files)} file(s) ready: {unlocked} unlocked, {already} had no protection' + \
        (f'; {len(converted)} converted from .xls' if converted else '')
    return dict(summary=summary, columns=['File', 'Result', 'Protection tags removed', 'Identical to original'],
                rows=rows, checks=checks, notes=notes, publish=publish)


# ---------------- 06-03 asset / package register ----------------
def step_names_check(job, d, out):
    import names as N
    staged, notes = stage_inputs(job, d / 'in')
    prev, new = [], []
    for s in staged:
        p = Path(s['staged'])
        if s['kind'] == 'register':
            got = N.read_register(p)
            if not got:
                notes.append(f"{s['name']}: not a File_Names register (no 'Assets & Packages' sheet), ignored.")
            prev += [dict(r, how='previous register', source=s['name']) for r in got]
        elif p.suffix.lower() in N.E.EXTS:
            a, pk, how = N.guess(p.name, N.sheet_names(p))
            new.append(dict(file=N.E.clean(p.name), asset=a, package=pk, how=how, source=s['name']))
    N.merge_qualifiers(new)
    N.number_duplicates(prev + new)
    known = {r['file'].upper() for r in prev}
    dup = [r for r in new if r['file'].upper() in known]
    new = sorted((r for r in new if r['file'].upper() not in known), key=lambda r: (r['asset'].upper(), r['package'].upper()))
    if dup:
        notes.append(f'{len(dup)} file(s) already in the previous register, kept as they were there: '
                     + ', '.join(r['file'] for r in dup))
    rows = prev + new
    probs = N.problems(rows)
    checks = [dict(check=p, ok=False, detail='fix on the review screen') for p in probs]
    summary = (f"{len({r['asset'] for r in rows})} assets, {len(rows)} files "
               f"({len(prev)} from the previous register, {len(new)} new)")
    return dict(summary=summary, items=rows, checks=checks, notes=notes + N.remarks(rows))


def step_names_build(job, d, out):
    import excel, names as N
    review = json.loads((d / 'review.json').read_text(encoding='utf-8'))
    rows = N.number_duplicates(review['rows'])
    probs = N.problems(rows)
    if probs:
        raise SystemExit('review has problems: ' + '; '.join(probs))
    g = N.build(rows, str(out))
    rc = excel.recalc(str(out / 'File_Names.xlsx'))
    checks = [dict(check='File_Names.xlsx recalculates in Excel with no formula errors', ok=rc['errors'] == 0,
                   detail='; '.join(rc['samples']) or 'Matrix COUNTIFS formulas')]
    table = [[a, ', '.join(g.codes[a]), len(g.codes[a])] for a in g.order]
    publish = [dict(file='File_Names.xlsx', kind='register'), dict(file='File_Names.txt', kind='register')] \
        if rc['errors'] == 0 else []
    return dict(summary=f'{len(g.order)} assets, {len(g.rows())} files', columns=['Asset', 'Packages', 'No.'],
                rows=table, checks=checks, notes=N.remarks(rows), publish=publish)


# ---------------- 06-02 stack divisions into one sheet ----------------
def step_divisions(job, d, out):
    import combine_divisions as C, verify_combine as V
    staged, notes = stage_inputs(job, d / 'in')
    files = C.collect([str(d / 'in')])
    dest = out / 'ALL DIVISIONS'
    work = Path(tempfile.mkdtemp())
    made, srcs, rc_by, rows, checks = {}, [], {}, [], []
    no_div = []
    for f in files:
        stem = Path(f).stem
        try:
            o, status, report, rc, src = C.combine(f, str(dest), str(work))
        except Exception as e:
            checks.append(dict(check=f'{stem} combined', ok=False, detail=f'{type(e).__name__}: {e}'))
            continue
        if not o:
            no_div.append(stem)
            continue
        made[stem] = o; srcs.append(src); rc_by[stem] = rc
    if no_div:
        notes.append(f'{len(no_div)} file(s) have no DIVISION tabs, so there is nothing to stack '
                     f'(single-sheet BOQs go straight to the next step): ' + ', '.join(no_div))
    vrows, total, vnotes = V.verify(srcs, str(dest)) if srcs else ([], {}, [])
    publish = []
    for r in vrows:
        stem = r['file']; rc = rc_by.get(stem, {'errors': 0, 'samples': []})
        ok = not r.get('missing') and r['bad'] == 0
        checks.append(dict(check=f'{stem}: every cell matches the source', ok=ok,
                           detail=f"{r.get('values', 0)} values, {r.get('formulas', 0)} formulas, {r['bad']} mismatches"))
        checks.append(dict(check=f'{stem}: recalculates with no formula errors', ok=rc['errors'] == 0,
                           detail='; '.join(rc['samples'])))
        rows.append([stem, r.get('divs', ''), r.get('items', ''), r.get('formulas', ''), r['bad'], rc['errors']])
        if ok and rc['errors'] == 0:
            publish.append(dict(file=f'ALL DIVISIONS/{Path(made[stem]).name}', kind='boq-divisions'))
    if len(publish) > 1:
        shutil.make_archive(str(out / 'ALL DIVISIONS'), 'zip', out, 'ALL DIVISIONS')
    shutil.rmtree(work, ignore_errors=True)
    summary = f'{len(publish)} of {len(files)} BOQ(s) stacked' + (f', {len(no_div)} with no DIVISION tabs' if no_div else '')
    return dict(summary=summary, columns=['File', 'Divisions', 'Items', 'Formulas', 'Mismatches', 'Formula errors'],
                rows=rows, checks=checks, notes=notes + vnotes, publish=publish,
                nothing_to_do=not files or len(no_div) == len(files))


# ---------------- 06-04 name the assets in the master ----------------
def _master_inputs(job, d):
    staged, notes = stage_inputs(job, d / 'in')
    masters = [s for s in staged if s['kind'] == 'master']
    regs = [s for s in staged if s['kind'] == 'register' and s['name'].lower().endswith('.xlsx')]
    if len(masters) != 1:
        raise SystemExit(f'tick exactly one PTE master ({len(masters)} ticked)')
    if len(regs) > 1:
        raise SystemExit('tick at most one register (File_Names.xlsx)')
    return masters[0], (regs[0] if regs else None), notes


def step_master_check(job, d, out):
    sys.path.insert(0, str(HERE / 'name_assets'))
    import expanded as X, names as N
    master, reg, notes = _master_inputs(job, d)
    try:
        L = X.detect(master['staged'])
    except X.LayoutError as e:
        raise SystemExit(str(e))
    assets = []
    if reg:
        rows = N.read_register(reg['staged'])
        g = N.E.Groups()
        for r in rows:
            g.add(r['asset'], r['package'])
        assets = [dict(name=a, packages=list(g.codes[a])) for a in g.order]
    else:
        notes.append('No register ticked: type the assets and their packages on the review screen.')
    if len(assets) > len(L['tabs']):
        notes.append(f"The register has {len(assets)} assets but the master only {len(L['tabs'])} asset tabs.")
    checks = [dict(check=p, ok=False, detail='the build stops on this') for p in L['problems']]
    other = [s for s in L['sheets'] if s not in L['tabs']]
    summary = (f"master with {len(L['tabs'])} asset tabs; {len(assets)} assets from the register"
               if reg else f"master with {len(L['tabs'])} asset tabs")
    return dict(summary=summary, items=assets, checks=checks, notes=notes, defects=L['defects'],
                master=master['name'], tabs=L['tabs'], other_sheets=other,
                options=dict(delete_spare=True, amount='pte'))


def step_master_build(job, d, out):
    sys.path.insert(0, str(HERE / 'name_assets'))
    import excel, expanded as X
    master, _, notes = _master_inputs(job, d)
    review = json.loads((d / 'review.json').read_text(encoding='utf-8'))
    opt = review['options']
    # one block of divisions per asset tab, no package blocks (Ahmed, 18 Nov 2026): each item's BOQ code goes
    # into column B (Client Code) when 06-06 loads it. expanded.plan still supports package blocks if wanted.
    assets = [dict(name=a['name'], packages=[]) for a in review['assets']]
    L = X.detect(master['staged'])
    if L['problems']:
        raise SystemExit('; '.join(L['problems']))
    probs = X.check_names(assets, [s for s in L['sheets'] if s not in L['tabs']])
    if probs:
        raise SystemExit('; '.join(probs))
    ops, expect = X.plan(L, assets, delete_spare=opt.get('delete_spare', True), amount=opt.get('amount', 'pte'))
    dst = out / master['name']
    excel.run_ops(master['staged'], ops, out=str(dst))
    checks = []
    fails = X.verify_static(str(dst), expect)
    checks.append(dict(check='Links and formulas: every summary row, total, weight and elemental group',
                       ok=not fails, detail='; '.join(fails[:6]) or f'{len(assets)} assets'))
    fails, n = X.verify_numbers(str(dst), expect)
    checks.append(dict(check='Numbers end to end: amounts planted in a throwaway copy reach every total',
                       ok=not fails, detail='; '.join(fails[:6]) or f'{n} values recalculated by Excel and read back'))
    r = __import__('subprocess').run([sys.executable, str(HERE / 'name_assets' / 'validate_xlsx.py'), str(dst)],
                                     capture_output=True, text=True)
    checks.append(dict(check='Excel will open it (structure check from the 06-04 skill)', ok=r.returncode == 0,
                       detail=(r.stdout.strip().splitlines() or [''])[-1]))
    amt = expect['amount_col']
    codes = {a['name']: a.get('packages', []) for a in review['assets']}
    rows = [[a['name'], ', '.join(codes.get(a['name'], [])), a['total'],
             f"Asset Summary row {expect['asum_rows'][k]}", f"columns {expect['elem_cols'][k]}.."]
            for k, a in enumerate(expect['assets'])]
    notes = notes + [f'Found in the template: {x}' for x in L['defects']]
    if expect['deleted']:
        notes.append(f"Deleted the spare tabs {expect['deleted'][0]} to {expect['deleted'][-1]} with their summary rows, "
                     'Schedule rows and elemental columns.')
    notes.append(f'The Elemental Summary reads column {amt} '
                 f"({'Pre-Tender Estimate, AAA quantities' if amt == expect['cols']['pte'] else 'Estimated Amount, Designer quantities'}).")
    notes.append('Project Summary names now link to Schedule of Assets, like the Asset Summary.')
    ok = all(c['ok'] for c in checks)
    return dict(summary=f"{len(assets)} assets named, one block of divisions each",
                columns=['Asset (tab)', 'BOQ codes (column B at loading)', 'Total Amounts row', 'Summary', 'Elemental'],
                rows=rows,
                checks=checks, notes=notes,
                publish=[dict(file=master['name'], kind='master-named')] if ok else [])


# ---------------- 06-06 load the BOQ items into the named master ----------------
def _load_inputs(job, d):
    staged, notes = stage_inputs(job, d / 'in')
    masters = [s for s in staged if s['kind'] == 'master-named']
    regs = [s for s in staged if s['kind'] == 'register' and s['name'].lower().endswith('.xlsx')]
    boqs = [s for s in staged if s['kind'] in ('boq-unlocked', 'boq') and s['name'].lower().endswith(EXCEL)]
    if len(masters) != 1:
        raise SystemExit(f'tick exactly one named PTE master ({len(masters)} ticked)')
    if len(regs) != 1:
        raise SystemExit('tick the register (File_Names.xlsx)')
    if not boqs:
        raise SystemExit('tick the BOQs to load')
    return masters[0], regs[0], boqs, notes


def step_load_check(job, d, out):
    import load_assign as A, names as N
    master, reg, boqs, notes = _load_inputs(job, d)
    a = A.assign(master['staged'], N.read_register(reg['staged']), [b['staged'] for b in boqs])
    items = [dict(file=b['file'], asset=b['asset'], package=b['package'], items=b.get('items'),
                  headings=b.get('headings'), qty=b.get('qty'), title=b.get('title', '')) for b in a['boqs']]
    counts = {}
    for b in a['boqs']:
        for label, n in b.get('divisions', {}).items():
            k = f"{b['package']}|{label}"
            counts[k] = counts.get(k, 0) + 1
    divmap = [dict(key=k, package=k.split('|', 1)[0], label=k.split('|', 1)[1], num=v['num'], why=v['why'],
                   boqs=counts.get(k, 0)) for k, v in a['divmap'].items()]
    master_divs = []
    if a['layout']:
        first = next(iter(a['layout'].values()))
        master_divs = [dict(num=x['num'], text=str(x['text'])) for x in first['blocks'][0]['divs']]
    checks = [dict(check=p, ok=False, detail='fix before building') for p in a['problems']]
    return dict(summary=f"{len(items)} BOQs, {sum(x['items'] or 0 for x in items)} items into "
                        f"{len({x['asset'] for x in items})} asset tabs",
                items=items, divmap=divmap, master_divs=master_divs, checks=checks, notes=notes + a['notes'],
                master=master['name'])


def step_load_build(job, d, out):
    import excel, load_assign as A, load_master as M, names as N
    master, reg, boqs, notes = _load_inputs(job, d)
    review = json.loads((d / 'review.json').read_text(encoding='utf-8'))
    divmap = {k: dict(num=v, why='approved on the review screen') for k, v in review['divmap'].items()}
    a = A.assign(master['staged'], N.read_register(reg['staged']), [b['staged'] for b in boqs], divmap)
    if a['problems']:
        raise SystemExit('; '.join(a['problems']))
    ops, expect = M.plan(a['layout'], a['loads'])
    dst = out / master['name']
    excel.run_ops(master['staged'], ops, out=str(dst))
    tabs = list(expect['tabs'])
    links = M.asset_summary_links(str(dst), tabs, {t: dict(cols=e['cols'], total=e['total'])
                                                   for t, e in expect['tabs'].items()})
    checks = []
    fails, n = M.verify_static(str(dst), expect, links)
    if len(links) != len(tabs):
        fails.append(f'Asset Summary links found for {len(links)} of {len(tabs)} tabs')
    checks.append(dict(check='Every loaded row matches its BOQ row (Ser., description, unit, quantity) with its formulas',
                       ok=not fails, detail='; '.join(fails[:5]) or f'{n} items read back'))
    fails, k = M.verify_numbers(str(dst), expect, links)
    checks.append(dict(check='Numbers: all quantities priced at 1 in a throwaway copy; package, tab and Asset Summary '
                             'totals equal the BOQ quantities, no formula errors',
                       ok=not fails, detail='; '.join(fails[:5]) or f'{k} totals recalculated by Excel and read back'))
    total_items = sum(b['items'] for b in a['boqs'])
    if n != total_items:
        checks.append(dict(check='Item count', ok=False, detail=f'{n} loaded, {total_items} read from the BOQs'))
    rows = [[b['tab'], b['package'], b['items'], b['headings'], f"{b['qty']:,.2f}", b['file']]
            for b in sorted(a['boqs'], key=lambda b: (tabs.index(b['tab']), b['block']))]
    ok = all(c['ok'] for c in checks)
    changed = [f"{k}: {v['num']}" for k, v in divmap.items()]
    return dict(summary=f"{total_items} items from {len(a['boqs'])} BOQs loaded into {len(tabs)} asset tabs",
                columns=['Asset (tab)', 'Code (column B)', 'Priced lines', 'Text lines', 'Designer QTY total', 'BOQ'],
                rows=rows,
                checks=checks, notes=notes + a['notes'] + ['Division mapping used: ' + '; '.join(changed)],
                publish=[dict(file=master['name'], kind='master-loaded')] if ok else [])


# ---------------- 06-05 link repeated items to one rate ----------------
def step_link(job, d, out):
    import excel, link_rates as R
    staged, notes = stage_inputs(job, d / 'in')
    masters = [s for s in staged if s['kind'] == 'master-loaded']
    if len(masters) != 1:
        raise SystemExit(f'tick exactly one loaded PTE master ({len(masters)} ticked)')
    src = masters[0]
    assets = R.read_items(src['staged'])
    probs = R.problems(assets)
    if probs:
        raise SystemExit('; '.join(probs))
    links = R.match(assets)
    dst = out / src['name']
    excel.run_ops(src['staged'], R.plan(assets, links), out=str(dst))
    fails, n = R.verify(str(dst), assets, links)
    R.report(str(out / 'Linking report.xlsx'), assets, links)
    checks = [dict(check='Every link points at its source (an item above it in the same asset, or an earlier asset) '
                         'and shows its rate (unique test rates in a throwaway copy, recalculated by Excel); '
                         'no formula errors',
                   ok=not fails, detail='; '.join(fails[:5]) or f'{n} links checked')]
    rows = []
    for k, a in enumerate(assets):
        b = sum(1 for l in links if l['tab'] == a['tab'] and l['colour'] == 'blue')
        o = sum(1 for l in links if l['tab'] == a['tab'] and l['colour'] == 'orange')
        own = sum(1 for l in links if l['tab'] == a['tab'] and l.get('same_asset'))
        srcs = sorted({l['src_tab'] for l in links if l['tab'] == a['tab'] and not l.get('same_asset')})
        rows.append([k + 1, a['tab'], len(a['items']), b, o, len(a['items']) - b - o, own, ', '.join(srcs)])
    blue = sum(1 for l in links if l['colour'] == 'blue')
    oranges = [l for l in links if l['colour'] == 'orange']
    notes += [f"ORANGE {l['tab']} row {l['row']}: \"{l['desc'][:60]}\" <- {l['src_tab']} row {l['src_row']} ({l['why']})"
              for l in oranges[:40]]
    if len(oranges) > 40:
        notes.append(f'... and {len(oranges) - 40} more oranges: see the Linking report.')
    own = sum(1 for l in links if l.get('same_asset'))
    return dict(summary=f'{blue} BLUE and {len(oranges)} ORANGE links over {len(assets)} assets '
                        f'({own} within the same asset)',
                columns=['Order', 'Asset', 'Items', 'Blue', 'Orange', 'Not linked', 'Within the asset',
                         'Linked to earlier assets'], rows=rows,
                checks=checks, notes=notes,
                publish=[dict(file=src['name'], kind='master-linked')] if not fails else [])


STEPS = {'boq-unlock': step_unlock, 'divisions': step_divisions, 'asset-names': step_names_build,
         'name-assets': step_master_build, 'load-boq': step_load_build, 'link-rates': step_link}
CHECKS = {'asset-names': step_names_check, 'name-assets': step_master_check, 'load-boq': step_load_check}


def main():
    module, jobdir = sys.argv[1], Path(sys.argv[2])
    check = '--check' in sys.argv[3:]
    job = json.loads((jobdir / 'job.json').read_text(encoding='utf-8'))
    out = jobdir / ('check' if check else 'out')
    shutil.rmtree(out, ignore_errors=True); out.mkdir(parents=True)
    res = (CHECKS if check else STEPS)[module](job, jobdir, out)
    res.update(stage='check' if check else 'build', module=module)
    res.setdefault('publish', [])
    (out / 'manifest.json').write_text(json.dumps(res, indent=1, ensure_ascii=False, default=str), encoding='utf-8')
    for c in res.get('checks', []):
        print(('PASS ' if c['ok'] else 'FAIL ') + c['check'] + (f" - {c['detail']}" if c.get('detail') else ''))
    for n in res.get('notes', []):
        print('NOTE ' + n)
    if not check and not res['publish'] and not res.get('nothing_to_do'):
        print(f"failed - {res['summary']}")
        sys.exit(1)
    print(f"{'check' if check else 'done'} - {res['summary']}")


if __name__ == '__main__':
    main()
