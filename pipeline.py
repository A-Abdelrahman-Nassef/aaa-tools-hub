"""pipeline.py - Areas Calculation module, run outside chat (Stage A).

usage:
  python pipeline.py <job_dir> --check-only                      build data, checks, review items and 3D previews
  python pipeline.py <job_dir> [--review review.json] [--project "Name"] [--rev R01] [--force] [--no-offline] [--no-test]
                          [--no-annex] [--client "Name"] [--letter D] [--annex-template file.pptx]

<job_dir>/spec/ holds one spec per asset, written in Claude Code from the drawings:
  *.py   a Python file defining SPEC = dict(...)   (same shape as toolkit/sample_villa_example.py)
  *.json the same dict as JSON (rooms as [code, name, ref, [[x,y],...], counted?])

--check-only  (the review screen reads these; nothing for the client is built)
  <job_dir>/check/data/<asset>.json, check/checks.json, check/review_items.json, check/preview/<asset>.html,
  check/manifest.json. Exit 0 even when checks fail (the review screen shows them); exit 1 only if a spec can't be read.

full run  (after the review is approved)
  review.json overrides, per spec file: {"name": "...", "atrium_level": "GF"}; then
  1. spec -> asset JSON (toolkit/asset_lib.build)            -> out/data/<asset>.json
  2. checks (A2 >= 0, A6 >= 0, skin vs report BUA within +-2%; A6 share is a warning) -> out/checks.json
     stops if a hard check fails, unless --force
  3. collected workbook (toolkit/build_workbook.py)           -> out/<Project> - AAA Areas Calculations - <rev>.xlsx
  4. 3D model per asset (toolkit/make_model.py)               -> out/models/<Asset> Area Model.html
  5. offline client copy per asset (toolkit/make_offline.py)  -> out/offline/AAA-Area-Model-<Asset>.html/.zip
  6. offline test (toolkit/offline_test.py, Playwright)
  7. 3D still per asset (toolkit/render_stills.py on the offline copy) -> out/stills/<asset>/compare.png
  8. Annex D PowerPoint (toolkit/build_annex_pptx.py, AAA template from templates/)
                                                               -> out/<Project> - Annex D - Areas Calculation - <rev>.pptx
  Writes out/manifest.json listing every output and the review it was built from.
"""
import argparse, json, os, re, runpy, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TK = ROOT / 'toolkit'
sys.path.insert(0, str(TK))
from asset_lib import VOID, build  # noqa: E402

A6_SHARE_MAX = 0.12     # A6 = internal partitions + small leftovers; Sample Villa (delivered) runs 6-10%.
                        # Above 12% a room is probably missing. Warning only.
BUA_TOL = 0.02          # external skin within +-2% of the report BUA

# Judgement calls the reviewer must confirm (agreed measurement rules, 00 Areas Calculation skill)
CALLS = {
    'perg': 'Under an open-slat pergola: treated as an uncovered terrace (IPMS 1 only, not in AAA BUA).',
    'cov': 'Treated as a covered terrace (e.g. under a cantilever or the slab above): counted in AAA BUA.',
    'plant': 'Plant / equipment roof with maintenance access only: not counted in either measure.',
    'shel': 'Sheltered area: IPMS 1 only, not in AAA BUA.',
    'nb': 'Notional boundary area (A1): excluded from AAA BUA, although report 2.11 ticks row A as counted.',
}


def find_template():
    """The AAA PowerPoint template dropped into templates/ (first .pptx/.potx, ignoring Office lock files)."""
    ts = sorted(p for p in (ROOT / 'templates').glob('*') if p.suffix.lower() in ('.pptx', '.potx') and not p.name.startswith('~$'))
    return ts[0] if ts else None


def slugify(s):
    return re.sub(r'[^A-Za-z0-9]+', '-', s).strip('-')


def load_spec(p: Path) -> dict:
    if p.suffix == '.py':
        g = runpy.run_path(str(p), run_name='spec')
        if 'SPEC' not in g:
            raise SystemExit(f'{p.name}: no SPEC dict defined')
        return g['SPEC']
    return json.loads(p.read_text(encoding='utf-8'))


def void_levels(spec) -> list:
    """Levels that have an atrium void (B1.3) room, bottom -> top, and the level currently marked as counted."""
    lv, counted = [], []
    for L in spec['levels']:
        rs = [r for r in spec['rooms'].get(L['id'], []) if r[0] == VOID]
        if rs:
            lv.append(L['id'])
            if any((r[4] if len(r) > 4 else True) for r in rs):
                counted.append(L['id'])
    return lv, counted


def apply_review(spec: dict, rv: dict) -> dict:
    if rv.get('name'):
        spec['asset'] = rv['name'].strip()
    lvl = rv.get('atrium_level')
    if lvl:
        rooms = {}
        for k, rs in spec['rooms'].items():
            rooms[k] = [tuple(r[:4]) + ((k == lvl),) if r[0] == VOID else r for r in rs]
        spec['rooms'] = rooms
    return spec


def checks(d: dict) -> list:
    out = []
    for k, s in d['sched'].items():
        base = s['internal'] if s['internal'] is not None else s['skin_gross']
        out.append(dict(asset=d['asset'], level=k, check='A2 external walls ≥ 0', value=s['a2'], unit='m²', ok=s['a2'] >= -0.005))
        out.append(dict(asset=d['asset'], level=k, check='A6 internal walls and leftovers ≥ 0', value=s['a6'], unit='m²', ok=s['a6'] >= -0.005))
        share = s['a6'] / base if base else 0
        out.append(dict(asset=d['asset'], level=k, check=f'A6 at most {A6_SHARE_MAX:.0%} of the internal area (else a room may be missing)',
                        value=round(share * 100, 2), unit='%', ok=share <= A6_SHARE_MAX, warn_only=True))
    rb = d.get('report_bua')
    if rb:
        diff = (d['totals']['skin'] - rb) / rb
        out.append(dict(asset=d['asset'], level='all', check=f'External skin within ±{BUA_TOL:.0%} of the report BUA ({rb:.2f} m²)',
                        value=round(diff * 100, 2), unit='%', ok=abs(diff) <= BUA_TOL))
    else:
        out.append(dict(asset=d['asset'], level='all', check='Report BUA given (for the ±2% skin check)', value=None, unit='',
                        ok=False, warn_only=True))
    return out


def review_items(d: dict, spec_file: str, spec: dict) -> dict:
    vl, counted = void_levels(spec)
    calls = [dict(key=f"{slugify(d['asset'])}-{i}", kind=o['key'], level=o['level'], name=o['name'], area=o['area'],
                  rule=CALLS[o['key']], note=o.get('note', ''))
             for i, o in enumerate(d['opens']) if o['key'] in CALLS]
    names = {L['id']: L['name'] for L in d['levels']}
    return dict(asset=d['asset'], spec=spec_file, levels=[dict(id=L['id'], name=L['name'], label=L['label']) for L in d['levels']],
                sched={k: dict(skin=s['skin_gross'], internal=s['internal'], a2=s['a2'], a6=s['a6'], bua=s['bua'], ipms=s['ipms'],
                               void_removed=s['void_removed']) for k, s in d['sched'].items()},
                totals=d['totals'], report_bua=d.get('report_bua'),
                atrium=dict(levels=[dict(id=k, name=names.get(k, k)) for k in vl], counted=counted) if vl else None,
                calls=calls)


def run(cmd: list, label: str):
    print(f'\n== {label}\n   ' + ' '.join(f'"{c}"' if ' ' in str(c) else str(c) for c in cmd), flush=True)
    env = dict(os.environ, PYTHONUTF8='1')
    r = subprocess.run([str(c) for c in cmd], cwd=TK, env=env)
    if r.returncode:
        raise SystemExit(f'{label} failed (exit {r.returncode})')


def wjson(p: Path, obj):
    p.write_text(json.dumps(obj, indent=1, ensure_ascii=False), encoding='utf-8')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('job')
    ap.add_argument('--check-only', action='store_true')
    ap.add_argument('--review', default=None, help='review.json approved on the hub review screen')
    ap.add_argument('--project', default=None)
    ap.add_argument('--rev', default='R01')
    ap.add_argument('--force', action='store_true', help='carry on even if a hard check fails')
    ap.add_argument('--no-offline', action='store_true')
    ap.add_argument('--no-test', action='store_true')
    ap.add_argument('--no-annex', action='store_true')
    ap.add_argument('--client', default='')
    ap.add_argument('--letter', default='D', help='annex letter')
    ap.add_argument('--annex-template', default=None, help='PowerPoint template (default: the one in templates/)')
    ap.add_argument('--template', default=str(TK / 'areas_template.xlsx'))
    ap.add_argument('--node-modules', default=str(ROOT / 'node_modules'))
    a = ap.parse_args()

    job = Path(a.job).resolve()
    specs = sorted(p for p in (job / 'spec').glob('*') if p.suffix in ('.py', '.json'))
    if not specs:
        raise SystemExit(f'no specs in {job / "spec"}')
    review = json.loads(Path(a.review).read_text(encoding='utf-8')) if a.review else {}
    out = job / ('check' if a.check_only else 'out')
    (out / 'data').mkdir(parents=True, exist_ok=True)
    project = a.project or job.name
    manifest = dict(project=project, rev=a.rev, stage='check' if a.check_only else 'build', assets=[], files={},
                    review=review or None)

    # 1-2. spec -> asset JSON + checks
    datas, allchecks, items, seen = [], [], [], set()
    for p in specs:
        try:
            spec = load_spec(p)
            spec = apply_review(spec, review.get('assets', {}).get(p.name, {}))
            d = build(spec)
        except SystemExit:
            raise
        except Exception as e:
            raise SystemExit(f'{p.name}: spec could not be read - {type(e).__name__}: {e}')
        slug = slugify(d['asset'])
        if slug in seen:
            raise SystemExit(f'two assets are named "{d["asset"]}" - asset names must differ')
        seen.add(slug)
        f = out / 'data' / f'{slug}.json'
        wjson(f, d)
        datas.append((d, f, slug))
        allchecks += checks(d)
        items.append(review_items(d, p.name, spec))
        t = d['totals']
        manifest['assets'].append(dict(asset=d['asset'], slug=slug, spec=p.name, data=str(f.relative_to(job)),
                                       skin=t['skin'], bua=t['bua'], ipms=t['ipms'], report_bua=d.get('report_bua')))
        print(f"{d['asset']}: skin {t['skin']:.2f}  BUA {t['bua']:.2f}  IPMS 1 {t['ipms']:.2f}  report BUA {d.get('report_bua')}")
    wjson(out / 'checks.json', allchecks)
    for c in allchecks:
        if not c['ok']:
            tag = 'WARNING     ' if c.get('warn_only') else 'CHECK FAILED'
            print(f"  {tag}  {c['asset']} {c['level']}: {c['check']} = {c['value']}{c['unit']}")
    bad = [c for c in allchecks if not c['ok'] and not c.get('warn_only')]
    manifest['checks_ok'] = not bad
    py = sys.executable

    if a.check_only:
        wjson(out / 'review_items.json', items)
        (out / 'preview').mkdir(exist_ok=True)
        for d, f, slug in datas:
            m = out / 'preview' / f'{slug}.html'
            run([py, 'make_model.py', 'model_template.txt', f, m], f"preview {d['asset']}")
            manifest['files'].setdefault('preview', []).append(str(m.relative_to(job)))
        wjson(out / 'manifest.json', manifest)
        print(f'\ncheck done - {len(bad)} hard check(s) failed')
        return

    if bad and not a.force:
        wjson(out / 'manifest.json', manifest)
        raise SystemExit(f'{len(bad)} check(s) failed - fix the spec or approve with a reason')

    # 3. workbook (never overwrite an existing revision)
    wb = out / f'{project} - AAA Areas Calculations - {a.rev}.xlsx'
    if wb.exists():
        raise SystemExit(f'{wb.name} already exists - pass a new --rev')
    run([py, 'build_workbook.py', a.template, wb] + [f for _, f, _ in datas], 'workbook')
    manifest['files']['workbook'] = str(wb.relative_to(job))

    # 4. models
    (out / 'models').mkdir(exist_ok=True)
    for d, f, slug in datas:
        m = out / 'models' / f"{d['asset']} Area Model.html"
        run([py, 'make_model.py', 'model_template.txt', f, m], f"model {d['asset']}")
        manifest['files'].setdefault('models', []).append(str(m.relative_to(job)))

    # 5-6. offline copies + test
    if not a.no_offline:
        nm = Path(a.node_modules)
        if not (nm / 'three' / 'build' / 'three.min.js').exists():
            print(f'\n!! skipping offline copies: {nm} missing - run setup.ps1 (npm install)')
        else:
            for d, f, slug in datas:
                m = out / 'models' / f"{d['asset']} Area Model.html"
                run([py, 'make_offline.py', m, d['asset'], out / 'offline', nm], f"offline {d['asset']}")
                h, z = out / 'offline' / f'AAA-Area-Model-{slug}.html', out / 'offline' / f'AAA-Area-Model-{slug}.zip'
                manifest['files'].setdefault('offline', []).append(dict(html=str(h.relative_to(job)), zip=str(z.relative_to(job))))
                if not a.no_test:
                    run([py, 'offline_test.py', h], f"offline test {d['asset']}")

    # 7-8. stills + Annex D PowerPoint (same asset JSONs as the workbook, so the numbers match it)
    if not a.no_annex:
        stills = out / 'stills'
        for d, f, slug in datas:
            h = out / 'offline' / f'AAA-Area-Model-{slug}.html'
            if h.exists():
                run([py, 'render_stills.py', h, f, stills / slug, 'compare'], f"still {d['asset']}")
            else:
                print(f"\n!! no offline model for {d['asset']}: its annex slide has no 3D still")
        tpl = Path(a.annex_template) if a.annex_template else find_template()
        if not tpl:
            print(f'\n!! no PowerPoint template in {ROOT / "templates"}: the annex uses the plain AAA navy layout')
        annex = out / f'{project} - Annex {a.letter} - Areas Calculation - {a.rev}.pptx'
        run([py, 'build_annex_pptx.py', '--out', annex, '--project', project, '--client', a.client, '--rev', a.rev,
             '--letter', a.letter, '--template', tpl or '', '--workbook', wb.name, '--stills', stills]
            + [f for _, f, _ in datas], 'annex')
        manifest['files']['annex'] = str(annex.relative_to(job))

    wjson(out / 'manifest.json', manifest)
    print(f'\ndone - {out / "manifest.json"}')


if __name__ == '__main__':
    main()
