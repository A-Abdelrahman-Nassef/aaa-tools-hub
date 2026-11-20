"""chain.py - which PTE steps are out of date, and re-running them in order.

Steps are chained through project files: a step job reads project files (job.json "inputs") and its results
replace project files of its output kind (worker.publish). So a finished step is out of date when, after it
finished:
  * a file it read was replaced (a newer upload, or a later step run published it again) or removed, or
  * a new file of a kind it read was added (e.g. more BOQs), or
  * an earlier step that produced one of its files is itself out of date.
Only the latest job of each step module in a project is judged.

"Update later steps" (start_chain) re-runs the step, or only the ones after it when the step itself is current.
Each re-run job carries jobs.chain=1; when it finishes, advance() starts the next out-of-date steps whose earlier
steps are all current, so 04 runs before 05 and 05 before 06. Review steps stop on their review screen as usual
(pre-filled from the previous approved review); approving them lets the chain carry on.
"""
import json, shutil

from . import db
from .modules import MODULES

BUSY = ('checking', 'queued', 'running')


def _dir(jid):
    from . import worker
    return worker.job_dir(jid)


def _jread(p, default=None):
    return json.loads(p.read_text(encoding='utf-8')) if p.exists() else default


def is_step(module):
    return MODULES.get(module, {}).get('runner') == 'step'


def latest_steps(pid):
    """Latest job of each step module in the project: {module: job row}."""
    out = {}
    for j in db.q('select * from jobs where project_id=? order by id', (pid,)):
        if is_step(j['module']):
            out[j['module']] = j
    return out


def status(pid):
    """{job id: dict(state, reasons)} for the latest step jobs. state: 'current' | 'stale' (its own files
    changed) | 'waits' (an earlier step it depends on is out of date); busy / unfinished jobs are 'n/a'."""
    files = db.q('select * from files where project_id=?', (pid,))
    by_id = {f['id']: f for f in files}
    jobs = latest_steps(pid)
    res, used = {}, {}
    for m, j in jobs.items():
        inp = (_jread(_dir(j['id']) / 'job.json') or {}).get('inputs', [])
        used[j['id']] = inp
        if j['status'] != 'done' or not j['finished'] or not inp:
            res[j['id']] = dict(state='n/a', reasons=[])
            continue
        why = []
        ids = {i['file_id'] for i in inp}
        for i in inp:
            f = by_id.get(i['file_id'])
            if not f:
                why.append(f"{i['name']} was removed from the project")
            elif f['created'] > j['finished'] and f['job_id'] != j['id']:
                src = f"job {f['job_id']}" if f['job_id'] else 'an upload'
                why.append(f"{f['name']} ({_label(f['kind'])}) changed on {f['created'][:16]} by {src}")
        kinds = {i['kind'] for i in inp}
        for f in files:
            if f['kind'] in kinds and f['id'] not in ids and f['created'] > j['finished'] and f['job_id'] != j['id']:
                why.append(f"new file {f['name']} ({_label(f['kind'])}) added on {f['created'][:16]}")
        res[j['id']] = dict(state='stale' if why else 'current', reasons=why)
    # out of date through an earlier step: a file it read was produced by a step job that is out of date
    by_job = {j['id']: j for j in jobs.values()}
    producer = {}
    for jid, inp in used.items():
        producer[jid] = {by_id[i['file_id']]['job_id'] for i in inp if i['file_id'] in by_id} - {None, jid}
    def upstream_bad(jid, seen=()):
        for p in producer.get(jid, ()):
            if p in seen or p not in by_job:
                continue
            if res.get(p, {}).get('state') in ('stale', 'waits') or by_job[p]['status'] in BUSY + ('review', 'setup'):
                return p
            q = upstream_bad(p, seen + (jid,))
            if q:
                return q
        return None
    for jid in list(res):
        if res[jid]['state'] in ('current', 'stale'):
            p = upstream_bad(jid)
            if p:
                res[jid]['state'] = 'waits' if res[jid]['state'] == 'current' else 'stale'
                res[jid]['reasons'].append(f"the earlier step {MODULES[by_job[p]['module']]['name']} (job {p}) "
                                           f"is out of date or not finished")
                res[jid]['waits_on'] = p
    return res


def downstream_pending(pid, jid):
    """Later step jobs that are out of date (directly or through earlier steps)."""
    st = status(pid)
    return [k for k, v in st.items() if k != jid and v['state'] in ('stale', 'waits')]


def _label(kind):
    from .modules import FILE_KINDS
    return FILE_KINDS.get(kind, kind)


def default_inputs(j):
    """Project files a step takes by default: from each input group, all files of the first kind the project has."""
    m = MODULES[j['module']]
    files = db.q('select * from files where project_id=? order by name', (j['project_id'],))
    have = {f['kind'] for f in files}
    out = []
    for grp in m['inputs']:
        k = next((k for k in grp if k in have), None)
        if k:
            out += [f for f in files if f['kind'] == k]
    return out


def start_step(j, files, chain=False):
    """Start a step job on these project files (as the "Run" button does). A finished job becomes the next
    revision; its last approved review is kept as review_prev.json to pre-fill the new review."""
    d = _dir(j['id'])
    m = MODULES[j['module']]
    inputs = [dict(file_id=f['id'], name=f['name'], kind=f['kind'], path=str((db.DATA / f['path']).resolve()))
              for f in files]
    (d / 'job.json').write_text(json.dumps(dict(inputs=inputs, options={}), indent=1, ensure_ascii=False),
                                encoding='utf-8')
    if (d / 'review.json').exists():
        shutil.copy2(d / 'review.json', d / 'review_prev.json')
    for x in ('review_draft.json', 'review.json'):
        (d / x).unlink(missing_ok=True)
    shutil.rmtree(d / 'check', ignore_errors=True)
    rev = j['rev']
    prev = _jread(d / 'out' / 'manifest.json')
    if prev and prev.get('stage') == 'build':      # a new run of a finished job is the next revision
        rev = f"R{int(rev[1:]) + 1:02d}"
    db.x('update jobs set status=?, rev=?, message=?, started=null, finished=null, chain=? where id=?',
         ('checking' if m.get('review') else 'queued', rev, 'updating after an earlier step changed' if chain else None,
          int(chain), j['id']))


def start_chain(j):
    """The "Update later steps" button: re-run this step if it is out of date, else start the steps after it."""
    st = status(j['project_id']).get(j['id'], {})
    if st.get('state') == 'stale':
        files = default_inputs(j)
        if not files:
            raise ValueError('the project has no files this step can take')
        start_step(j, files, chain=True)
        return [j['id']]
    return advance(j['id'], force=True)


def advance(jid, force=False):
    """After a chained job finished: start the out-of-date later steps whose earlier steps are all current.
    Returns the started job ids."""
    j = db.q('select * from jobs where id=?', (jid,), one=True)
    if not j or (not j['chain'] and not force):
        return []
    db.x('update jobs set chain=0 where id=?', (jid,))
    if j['status'] != 'done':
        return []
    started = []
    for k, v in status(j['project_id']).items():
        if k == jid or v['state'] != 'stale' or v.get('waits_on'):
            continue
        jj = db.q('select * from jobs where id=?', (k,), one=True)
        if jj['status'] in BUSY:
            continue
        files = default_inputs(jj)
        if files:
            start_step(jj, files, chain=True)
            started.append(k)
    return started
