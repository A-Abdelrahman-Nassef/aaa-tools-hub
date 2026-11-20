"""naming.py - output file names (Ahmed, 19 Nov 2026):

    <Date> - <Client> - <Project> - <Task> - Rev.<nn> - <Initials><ext>
    2026.11.19 - Sample Developer - VT - PTE & Tender Evaluation - Rev.00 - AAN.xlsx

* Date      the day the file is produced (local time), yyyy.mm.dd
* Client    the project's client (left out when empty)
* Project   the project's short name for files (projects.short_name), else its name
* Task      what the file is (see TASKS below and the Areas names in rename_areas)
* Rev.nn    per project and task: Rev.00 for the first file issued, then up by one for every new file
            (the PTE master's 04, 05 and 06 files share one counter, "PTE & Tender Evaluation")
* Initials  users.initials of whoever ran the step (jobs.run_by, else the job's creator), else the initials
            of the user's name

BOQ files keep their own names: the register matches BOQs by file name.
Every issued name is recorded in the `issued` table, which is where the Rev counter comes from.
"""
import json, re
from datetime import datetime
from pathlib import Path

from . import db

# step module -> {file in the step's out folder (or suffix pattern): task}
TASK_MASTER = 'PTE & Tender Evaluation'
TASKS = {
    'boq-unlock': {'Unlocked BOQs.zip': 'Unlocked BOQs'},
    'divisions': {'*.zip': 'BOQs - All Divisions'},
    'asset-names': {'File_Names.xlsx': 'Asset & Package Register', 'File_Names.txt': 'Asset & Package Register'},
    'name-assets': {'master': TASK_MASTER},
    'load-boq': {'master': TASK_MASTER},
    'link-rates': {'master': TASK_MASTER, 'Linking report.xlsx': 'PTE Linking Report'},
}
MASTER_KINDS = ('master-named', 'master-loaded', 'master-linked')
BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')


def clean(s):
    return re.sub(r'\s+', ' ', BAD.sub('-', str(s or ''))).strip(' .-')


def initials_of(u):
    if not u:
        return ''
    if (u['initials'] or '').strip():
        return clean(u['initials']).upper()
    return ''.join(w[0] for w in re.findall(r'[A-Za-z]+', u['name'] or '')).upper()


def parts(job):
    """Date, client, project and initials for a job's files."""
    p = db.q('select * from projects where id=?', (job['project_id'],), one=True)
    uid = job['run_by'] if 'run_by' in job.keys() and job['run_by'] else job['created_by']
    u = db.q('select * from users where id=?', (uid,), one=True)
    return dict(date=datetime.now().strftime('%Y.%m.%d'), client=clean(p['client']),
                project=clean(p['short_name'] or p['name']), initials=initials_of(u) or 'XX')


def next_rev(pid, task):
    r = db.q('select max(rev) m from issued where project_id=? and task=?', (pid, task), one=True)
    return 0 if r['m'] is None else r['m'] + 1


def file_name(pt, task, rev, ext):
    bits = [pt['date'], pt['client'], pt['project'], clean(task), f'Rev.{rev:02d}', pt['initials']]
    return ' - '.join(b for b in bits if b) + ext


def issue(job, task, files):
    """Rename these files (Paths, all for one task) to the next revision's name; returns {old Path: new Path}.
    Files of one task from one run share a Rev (e.g. the register's .xlsx and .txt)."""
    if not files:
        return {}
    pt = parts(job)
    rev = next_rev(job['project_id'], task)
    out = {}
    for f in files:
        new = f.with_name(file_name(pt, task, rev, f.suffix))
        if new != f:
            if new.exists():
                new.unlink()
            f.rename(new)
        out[f] = new
        db.x('insert into issued(project_id, task, rev, name, job_id, created) values(?,?,?,?,?,?)',
             (job['project_id'], task, rev, new.name, job['id'], db.now()))
    return out


def rename_step(job, out, man):
    """After a PTE step built: rename its deliverables (not BOQs) and update manifest.json. Returns the manifest."""
    rules = TASKS.get(job['module'], {})
    by_task = {}
    publish = man.get('publish', [])
    for p in publish:                                          # the master travels from step to step
        if p['kind'] in MASTER_KINDS and 'master' in rules:
            by_task.setdefault(rules['master'], []).append(out / p['file'])
    for pat, task in rules.items():
        if pat == 'master':
            continue
        hits = sorted(out.glob(pat)) if '*' in pat else ([out / pat] if (out / pat).is_file() else [])
        by_task.setdefault(task, []).extend(h for h in hits if h not in by_task.get(task, []))
    moved = {}
    for task, files in by_task.items():
        moved.update(issue(job, task, [f for f in files if f.is_file()]))
    rel = {str(k.relative_to(out)).replace('\\', '/'): str(v.relative_to(out)).replace('\\', '/') for k, v in moved.items()}
    for p in publish:
        p['file'] = rel.get(p['file'], p['file'])
    man['renamed'] = rel
    (out / 'manifest.json').write_text(json.dumps(man, indent=1, ensure_ascii=False), encoding='utf-8')
    return man


def rename_areas(job, d):
    """After an Areas build: the collected workbook, the Annex D deck, and each asset's 3D model and offline copy."""
    m = d / 'out' / 'manifest.json'
    if not m.exists():
        return
    man = json.loads(m.read_text(encoding='utf-8'))
    f = man.get('files', {})

    def one(rel, task):
        if not rel or not (d / rel).is_file():
            return rel
        new = issue(job, task, [d / rel])[d / rel]
        return str(new.relative_to(d))
    f['workbook'] = one(f.get('workbook'), 'AAA Areas Calculations')
    f['annex'] = one(f.get('annex'), 'Annex D - Areas Calculation')
    assets = [a['asset'] for a in man.get('assets', [])]
    models = f.get('models', [])
    for i, rel in enumerate(models):
        a = assets[i] if i < len(assets) else Path(rel).stem
        models[i] = one(rel, f'Area Model - {a}')
    for i, off in enumerate(f.get('offline', [])):
        a = assets[i] if i < len(assets) else ''
        task = f'Area Model - {a} - Offline'
        pair = [d / off[k] for k in ('html', 'zip') if off.get(k) and (d / off[k]).is_file()]
        moved = issue(job, task, pair)
        for k in ('html', 'zip'):
            if off.get(k) and (d / off[k]) in moved:
                off[k] = str(moved[d / off[k]].relative_to(d))
    m.write_text(json.dumps(man, indent=1, ensure_ascii=False), encoding='utf-8')
