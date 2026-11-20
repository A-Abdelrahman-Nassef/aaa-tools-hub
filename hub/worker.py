"""worker.py - background job queue. One task at a time; runs the module's scripts (no Claude).

Job status flow (areas module):
  awaiting_spec -> checking -> review | spec_error      (check: pipeline --check-only, with review_draft.json)
  review -> queued -> running -> done | failed          (build: pipeline with the approved review.json)
Step modules (PTE Phase 1, toolkit/pte/run_step.py):
  setup -> queued -> running -> done | failed           (no review)
  setup -> checking -> review | failed -> queued -> running -> done | failed   (review=True, e.g. asset-names)
  On done, the files the step lists under "publish" are added to the project's files.
"""
import json, os, re, shutil, subprocess, sys, threading
from pathlib import Path
from . import db
from .modules import MODULES

ROOT = db.ROOT
JOBS = db.DATA / 'jobs'
_stop = threading.Event()


def job_dir(jid: int) -> Path:
    return JOBS / str(jid)


def _pipeline(job, args, logname):
    d = job_dir(job['id'])
    cmd = [sys.executable, str(ROOT / 'pipeline.py'), str(d)] + args
    with open(d / logname, 'w', encoding='utf-8') as log:
        log.write(' '.join(cmd) + '\n\n'); log.flush()
        r = subprocess.run(cmd, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT, env={**os.environ, 'PYTHONUTF8': '1'})
    tail = (d / logname).read_text(encoding='utf-8', errors='ignore').strip().splitlines()[-1:]
    return r.returncode, (tail[0] if tail else '')


def _check(job):
    jid = job['id']
    d = job_dir(jid)
    args = ['--check-only']
    if (d / 'review_draft.json').exists():
        args += ['--review', str(d / 'review_draft.json')]
    code, last = _pipeline(job, args, 'check.log')
    if code:
        db.x("update jobs set status='spec_error', message=? where id=?", (last, jid))
    else:
        db.x("update jobs set status='review', message=? where id=?", (last.replace('check done - ', '').strip(), jid))


def _build(job):
    jid = job['id']
    d = job_dir(jid)
    proj = db.q('select name, client from projects where id=?', (job['project_id'],), one=True)
    pname = re.sub(r'[<>:"/\\|?*]+', '-', proj['name']).strip(' .-') or f'Project {job["project_id"]}'   # used in file names
    db.x("update jobs set status='running', started=?, message=null where id=?", (db.now(), jid))
    review = json.loads((d / 'review.json').read_text(encoding='utf-8'))
    args = ['--project', pname, '--client', proj['client'] or '', '--rev', job['rev'], '--review', str(d / 'review.json')]
    if review.get('override_reason'):
        args.append('--force')
    code, last = _pipeline(job, args, 'run.log')
    msg = last if code else None
    m = d / 'out' / 'manifest.json'
    if not code and m.exists():
        from . import naming
        naming.rename_areas(job, d)
        man = json.loads(m.read_text(encoding='utf-8'))
        msg = '; '.join(f"{a['asset']}: BUA {a['bua']:.2f}, IPMS 1 {a['ipms']:.2f}" for a in man['assets'])
    db.x('update jobs set status=?, finished=?, message=? where id=?', ('failed' if code else 'done', db.now(), msg, jid))
    # Stage A: the Claude step was done by hand in Claude Code (subscription), so the run itself costs nothing.
    db.log_cost(job['created_by'], job['project_id'], jid, job['module'], 'claude-code (manual)', 0.0,
                note='Stage A: spec written in Claude Code; scripts only')


def _step(job, check):
    """PTE Phase 1 step module: run_step.py, then (build) put the published files into the project."""
    jid = job['id']
    d = job_dir(jid)
    if not check:
        db.x("update jobs set status='running', started=?, message=null where id=?", (db.now(), jid))
    cmd = [sys.executable, str(ROOT / 'toolkit' / 'pte' / 'run_step.py'), job['module'], str(d)] + (['--check'] if check else [])
    logname = 'check.log' if check else 'run.log'
    with open(d / logname, 'w', encoding='utf-8') as log:
        log.write(' '.join(cmd) + '\n\n'); log.flush()
        r = subprocess.run(cmd, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT, env={**os.environ, 'PYTHONUTF8': '1'})
    lines = (d / logname).read_text(encoding='utf-8', errors='ignore').strip().splitlines()
    last = lines[-1] if lines else ''
    msg = re.sub(r'^(done|check|failed) - ', '', last)
    if check:
        db.x('update jobs set status=?, message=? where id=?', ('failed' if r.returncode else 'review', msg, jid))
        return
    if not r.returncode:
        man = json.loads((d / 'out' / 'manifest.json').read_text(encoding='utf-8'))
        from . import naming             # Date - Client - Project - Task - Rev.nn - Initials (not BOQs)
        man = naming.rename_step(job, d / 'out', man)
        n = publish(job, man)
        if n:
            msg += f'; {n} file(s) added to the project'
    db.x('update jobs set status=?, finished=?, message=? where id=?', ('failed' if r.returncode else 'done', db.now(), msg, jid))
    db.log_cost(job['created_by'], job['project_id'], jid, job['module'], 'scripts', 0.0, note='PTE step: scripts only, no Claude')
    from . import chain                  # "Update later steps": start the next out-of-date steps (stops on failure)
    started = chain.advance(jid)
    if started:
        db.x('update jobs set message=? where id=?', (msg + f"; next: job {', '.join(map(str, started))}", jid))


def publish(job, man):
    """Copy a step's results into the project's files (replacing the same name and kind)."""
    pid, n = job['project_id'], 0
    renamed = set((man.get('renamed') or {}).values())     # outputs naming.py gave a Date / Rev name this run
    done = set()                                            # project file ids written by this run
    for p in man.get('publish', []):
        src = job_dir(job['id']) / 'out' / p['file']
        if not src.is_file():
            continue
        dest = db.DATA / 'projects' / str(pid) / 'files' / p['kind'] / src.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        rel = str(dest.relative_to(db.DATA))
        old = db.q('select id, path from files where project_id=? and kind=? and name=?', (pid, p['kind'], src.name), one=True)
        if not old and p['file'] in renamed:
            # only a renamed output (new date / Rev in its name) replaces this job's previous file of that kind;
            # BOQs keep their names, so each one is its own project file (20 Nov 2026: 8 BOQs had become 1)
            old = next((f for f in db.q('select id, name, path from files where project_id=? and kind=? and job_id=?',
                                         (pid, p['kind'], job['id']))
                        if Path(f['name']).suffix == src.suffix and f['id'] not in done), None)
            if old and db.DATA / old['path'] != dest:
                (db.DATA / old['path']).unlink(missing_ok=True)
        if old:
            db.x('update files set name=?, path=?, size=?, uploaded_by=?, created=?, job_id=? where id=?',
                 (src.name, rel, src.stat().st_size, job['created_by'], db.now(), job['id'], old['id']))
            done.add(old['id'])
        else:
            done.add(db.x('insert into files(project_id,kind,name,path,size,uploaded_by,created,job_id) values(?,?,?,?,?,?,?,?)',
                          (pid, p['kind'], src.name, rel, src.stat().st_size, job['created_by'], db.now(), job['id'])))
        n += 1
    return n


def _run(job):
    if MODULES.get(job['module'], {}).get('runner') == 'step':
        _step(job, job['status'] == 'checking')
    else:
        (_check if job['status'] == 'checking' else _build)(job)


def loop():
    db.x("update jobs set status='queued' where status='running'")   # a crash mid-run: run it again
    while not _stop.is_set():
        job = db.q("select * from jobs where status in ('checking','queued') order by id limit 1", one=True)
        if job:
            try:
                _run(job)
            except Exception as e:  # keep the queue alive
                db.x("update jobs set status='failed', finished=?, message=? where id=?", (db.now(), f'worker error: {e}', job['id']))
        else:
            _stop.wait(1)


def start():
    t = threading.Thread(target=loop, name='hub-worker', daemon=True)
    t.start()
    return t


def stop():
    _stop.set()
