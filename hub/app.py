"""app.py - AAA Tools Hub web app (Stage A: runs on Ahmed's laptop).

run:  .\\.venv\\Scripts\\python.exe -m uvicorn hub.app:app --port 8765
"""
import json, mimetypes, re, secrets, shutil
from datetime import datetime
from contextlib import asynccontextmanager
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from . import chain, db, naming, worker
from .modules import ALLOW_PY_SPECS, FILE_KINDS, MODULES, UPLOAD_KINDS

HERE = Path(__file__).resolve().parent
MAX_UPLOAD = 200 * 1024 * 1024   # 200 MB per file


def secret_key() -> str:
    f = db.DATA / 'secret.key'
    db.DATA.mkdir(parents=True, exist_ok=True)
    if not f.exists():
        f.write_text(secrets.token_hex(32), encoding='utf-8')
    return f.read_text(encoding='utf-8').strip()


@asynccontextmanager
async def lifespan(_):
    db.init()
    worker.start()
    yield
    worker.stop()


app = FastAPI(title='AAA Tools Hub', lifespan=lifespan)
app.add_middleware(SessionMiddleware, secret_key=secret_key(), session_cookie='aaa_hub', same_site='lax', max_age=12 * 3600)
app.mount('/static', StaticFiles(directory=HERE / 'static'), name='static')
T = Jinja2Templates(directory=HERE / 'templates')


def job_name(j):
    """A job shows its module's current name; the title stored with old jobs is the name at the time."""
    m = MODULES.get(j['module'])
    return m['name'] if m else j['title']


T.env.globals.update(pending_users=lambda: db.q("select count(*) n from users where status='pending'", one=True)['n'],
                     MODULES=MODULES, FILE_KINDS=FILE_KINDS, UPLOAD_KINDS=UPLOAD_KINDS, job_name=job_name,
                     num=lambda n: f'{n or 0:,}')


class LoginRequired(Exception):
    pass


class AwaitingApproval(Exception):
    pass


@app.exception_handler(LoginRequired)
async def _login(request, exc):
    return RedirectResponse('/login', 303)


@app.exception_handler(AwaitingApproval)
async def _waiting(request, exc):
    return RedirectResponse('/waiting', 303)


def user(request: Request):
    """The signed-in, active user. Checked on every request, so approval, deactivation and admin changes apply at once:
    a pending account only sees the waiting page; a deactivated or rejected one is signed out."""
    uid = request.session.get('uid')
    u = db.q('select * from users where id=?', (uid,), one=True) if uid else None
    if not u:
        raise LoginRequired()
    if u['status'] == 'pending':
        raise AwaitingApproval()
    if u['status'] != 'active':
        request.session.clear()
        raise LoginRequired()
    return u


def admin(request: Request):
    """Admin-only pages and actions: blocked on the server for members, not only hidden in the menu."""
    u = user(request)
    if not u['is_admin']:
        raise HTTPException(403, 'Admins only')
    return u


def page(request, name, **ctx):
    return T.TemplateResponse(request, name, ctx)


def safe_name(n: str) -> str:
    n = Path(n or 'file').name
    n = re.sub(r'[^\w.\- ()&,]+', '_', n).strip(' .')      # commas are kept: BOQ names like "-96,97-A3-02"
    return n or 'file'


def inside(base: Path, rel: str) -> Path:
    p = (base / rel).resolve()
    if base.resolve() not in p.parents and p != base.resolve():
        raise HTTPException(404)
    if not p.is_file():
        raise HTTPException(404)
    return p


async def save_upload(up: UploadFile, dest: Path) -> int:
    dest.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with open(dest, 'wb') as f:
        while chunk := await up.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD:
                f.close(); dest.unlink(missing_ok=True)
                raise HTTPException(413, f'{up.filename} is over 200 MB')
            f.write(chunk)
    return size


def can_see(u, pid):
    """Project access: admins see every project; a member only the projects they belong to (project_members),
    and never an archived one."""
    if u['is_admin']:
        return True
    return bool(db.q("""select 1 from project_members m join projects p on p.id=m.project_id
                        where m.project_id=? and m.user_id=? and p.archived=0""", (pid, u['id']), one=True))


ARCHIVE_OK = ('/archive', '/restore', '/delete')


@app.middleware('http')
async def archived_read_only(request: Request, call_next):
    """An archived project is read-only for everyone: no uploads, new jobs, runs or approvals until it is restored."""
    if request.method == 'POST' and not request.url.path.endswith(ARCHIVE_OK):
        m = re.match(r'^/(projects|jobs)/(\d+)/', request.url.path)
        if m:
            if m.group(1) == 'projects':
                pid = int(m.group(2))
            else:
                j = db.q('select project_id from jobs where id=?', (int(m.group(2)),), one=True)
                pid = j['project_id'] if j else 0
            if db.q('select 1 from projects where id=? and archived=1', (pid,), one=True):
                return PlainTextResponse('This project is archived (read-only). An admin can restore it on its page.', 409)
    return await call_next(request)


def project_or_404(pid, u):
    """The project, or 404 (not 403) for one this user may not see, so its existence is not given away."""
    p = db.q('select * from projects where id=?', (pid,), one=True)
    if not p or not can_see(u, pid):
        raise HTTPException(404)
    return p


def job_or_404(jid, u):
    j = db.q('select * from jobs where id=?', (jid,), one=True)
    if not j or not can_see(u, j['project_id']):
        raise HTTPException(404)
    return j


# ---------------- login ----------------
@app.get('/login', response_class=HTMLResponse)
def login_form(request: Request):
    return page(request, 'login.html', error=None)


@app.post('/login')
def login(request: Request, email: str = Form(...), password: str = Form(...)):
    u = db.q('select * from users where lower(email)=lower(?)', (email.strip(),), one=True)
    if not u or not db.check_pw(password, u['pw_hash']):
        return page(request, 'login.html', error='Email or password is wrong.')
    if u['status'] == 'inactive':
        return page(request, 'login.html', error='This account has been deactivated. Ask an admin.')
    if u['status'] == 'rejected':
        return page(request, 'login.html', error='This account request was not approved. Ask an admin.')
    request.session.clear(); request.session['uid'] = u['id']
    return RedirectResponse('/waiting' if u['status'] == 'pending' else '/projects', 303)


@app.get('/logout')
def logout(request: Request):
    request.session.clear()
    return RedirectResponse('/login', 303)


# ---------------- sign-up ----------------
PW_MIN = 10


def signup_domain():
    return (db.setting('signup_domain', '') or '').strip().lstrip('@').lower()


def pw_problem(pw, again):
    if len(pw) < PW_MIN:
        return f'The password must be at least {PW_MIN} characters.'
    if pw != again:
        return 'The two passwords are not the same.'
    return None


@app.get('/signup', response_class=HTMLResponse)
def signup_form(request: Request):
    return page(request, 'signup.html', error=None, f={}, domain=signup_domain())


@app.post('/signup')
def signup(request: Request, name: str = Form(''), email: str = Form(''), password: str = Form(''),
           password2: str = Form('')):
    """A new account starts as pending: it can sign in, but only sees the waiting page until an admin approves it."""
    dom = signup_domain()
    name, email = re.sub(r'\s+', ' ', name).strip(), email.strip().lower()
    f = dict(name=name, email=email)
    err = None
    if not dom:
        err = 'Sign-up is closed. Ask an admin for an account.'
    elif len(name) < 3:
        err = 'Enter your full name.'
    elif not re.fullmatch(r'[^@\s]+@[^@\s]+', email):
        err = 'Enter a valid email address.'
    elif not email.endswith('@' + dom):
        err = f'Only {dom} work emails can sign up. Use your @{dom} address, or ask an admin for an account.'
    else:
        err = pw_problem(password, password2)
    if not err:
        old = db.q('select status from users where lower(email)=?', (email,), one=True)
        if old:
            err = {'pending': 'This email has already signed up and is waiting for approval. Sign in to see its status.',
                   'rejected': 'This email\'s request was not approved. Ask an admin.'}.get(
                old['status'], 'This email already has an account. Sign in instead.')
    if err:
        return page(request, 'signup.html', error=err, f=f, domain=dom)
    uid = db.x("insert into users(email,name,pw_hash,is_admin,created,status) values(?,?,?,0,?,'pending')",
               (email, name, db.hash_pw(password), db.now()))
    request.session.clear(); request.session['uid'] = uid
    return RedirectResponse('/waiting', 303)


@app.get('/waiting', response_class=HTMLResponse)
def waiting(request: Request):
    uid = request.session.get('uid')
    u = db.q('select * from users where id=?', (uid,), one=True) if uid else None
    if not u:
        return RedirectResponse('/login', 303)
    if u['status'] == 'active':
        return RedirectResponse('/projects', 303)
    if u['status'] != 'pending':
        request.session.clear()
        return RedirectResponse('/login', 303)
    return page(request, 'waiting.html', who=u)


# ---------------- users (admins) ----------------
@app.get('/users', response_class=HTMLResponse)
def users_page(request: Request):
    u = admin(request)
    rows = db.q("""select * from users order by case status when 'pending' then 0 when 'active' then 1
                   when 'inactive' then 2 else 3 end, is_admin desc, lower(name)""")
    return page(request, 'users.html', u=u, rows=rows, domain=signup_domain(),
                msg=request.query_params.get('msg'), err=request.query_params.get('err'))


def back(msg=None, err=None):
    from urllib.parse import urlencode
    q = urlencode({k: v for k, v in (('msg', msg), ('err', err)) if v})
    return RedirectResponse('/users' + ('?' + q if q else ''), 303)


def active_admins(exclude=None):
    return db.q("select count(*) n from users where is_admin=1 and status='active' and id!=?",
                (exclude or 0,), one=True)['n']


@app.post('/users/{uid}/action')
def user_action(request: Request, uid: int, action: str = Form(...)):
    """approve / reject (pending), make_admin / remove_admin, deactivate / reactivate."""
    me = admin(request)
    t = db.q('select * from users where id=?', (uid,), one=True)
    if not t:
        raise HTTPException(404)
    lose_admin = action in ('remove_admin', 'deactivate') and t['is_admin'] and t['status'] == 'active'
    if lose_admin and active_admins(exclude=uid) == 0:
        return back(err='The hub needs at least one active admin: make someone else admin first.')
    if uid == me['id'] and action in ('remove_admin', 'deactivate'):
        return back(err='You cannot remove your own admin role or deactivate yourself; ask another admin.')
    allowed = {'approve': ('pending', 'rejected'), 'reject': ('pending',), 'deactivate': ('active',),
               'reactivate': ('inactive',), 'make_admin': ('active',), 'remove_admin': ('active',)}
    if action not in allowed or t['status'] not in allowed[action]:
        return back(err=f"Cannot {action.replace('_', ' ')} {t['name']} ({t['status']}).")
    if action == 'approve':
        db.x("update users set status='active' where id=?", (uid,)); msg = f"{t['name']} approved: they can sign in now."
    elif action == 'reject':
        db.x("update users set status='rejected' where id=?", (uid,)); msg = f"{t['name']}'s request rejected."
    elif action == 'deactivate':
        db.x("update users set status='inactive' where id=?", (uid,)); msg = f"{t['name']} deactivated (signed out at once)."
    elif action == 'reactivate':
        db.x("update users set status='active' where id=?", (uid,)); msg = f"{t['name']} reactivated."
    elif action == 'make_admin':
        db.x('update users set is_admin=1 where id=?', (uid,)); msg = f"{t['name']} is now an admin."
    else:
        db.x('update users set is_admin=0 where id=?', (uid,)); msg = f"{t['name']} is now a member."
    return back(msg=msg)


@app.post('/users/{uid}/edit')
def user_edit(request: Request, uid: int, name: str = Form(''), email: str = Form(''), initials: str = Form(''),
              password: str = Form(''), password2: str = Form('')):
    """Name, email (= sign-in), initials for file names, and an optional new password."""
    admin(request)
    t = db.q('select * from users where id=?', (uid,), one=True)
    if not t:
        raise HTTPException(404)
    name, email = re.sub(r'\s+', ' ', name).strip() or t['name'], email.strip().lower() or t['email']
    if not re.fullmatch(r'[^@\s]+@[^@\s]+', email):
        return back(err=f'"{email}" is not an email address.')
    if db.q('select 1 from users where lower(email)=? and id!=?', (email, uid), one=True):
        return back(err=f'{email} is already used by another account.')
    ini = re.sub(r'[^A-Za-z]', '', initials).upper()[:6] or None
    db.x('update users set name=?, email=?, initials=? where id=?', (name, email, ini, uid))
    if password:
        p = pw_problem(password, password2)
        if p:
            return back(err=f"{name}: details saved, password not changed. {p}")
        db.x('update users set pw_hash=? where id=?', (db.hash_pw(password), uid))
        return back(msg=f"{name}: details saved and password changed. Give them the new password in person.")
    return back(msg=f'{name}: details saved.')


@app.post('/users/new')
def user_new(request: Request, name: str = Form(''), email: str = Form(''), initials: str = Form(''),
             role: str = Form('member'), password: str = Form(''), password2: str = Form('')):
    """An admin adds an account directly (active at once, any email domain)."""
    admin(request)
    name, email = re.sub(r'\s+', ' ', name).strip(), email.strip().lower()
    if len(name) < 3 or not re.fullmatch(r'[^@\s]+@[^@\s]+', email):
        return back(err='Give a full name and an email address.')
    if db.q('select 1 from users where lower(email)=?', (email,), one=True):
        return back(err=f'{email} already has an account.')
    p = pw_problem(password, password2)
    if p:
        return back(err=p)
    db.x("insert into users(email,name,pw_hash,is_admin,created,status,initials) values(?,?,?,?,?,'active',?)",
         (email, name, db.hash_pw(password), int(role == 'admin'), db.now(),
          re.sub(r'[^A-Za-z]', '', initials).upper()[:6] or None))
    return back(msg=f'{name} added as {"admin" if role == "admin" else "member"}. Give them the password in person.')


@app.post('/users/domain')
def user_domain(request: Request, domain: str = Form('')):
    """The work-email domain that may sign up; empty closes sign-up."""
    admin(request)
    d = domain.strip().lstrip('@').lower()
    if d and not re.fullmatch(r'[a-z0-9.-]+\.[a-z]{2,}', d):
        return back(err=f'"{domain}" is not a domain like aaa-consultancy.com.')
    db.set_setting('signup_domain', d)
    return back(msg=f'Sign-up open to @{d} emails.' if d else 'Sign-up closed.')


@app.get('/')
def home(request: Request):
    u = user(request)
    return RedirectResponse('/projects', 303)


# ---------------- projects ----------------
@app.get('/projects', response_class=HTMLResponse)
def projects(request: Request):
    u = user(request)
    archived = bool(u['is_admin']) and request.query_params.get('archived') == '1'   # admins: the archived list
    rows = db.q("""select p.*, (select count(*) from files f where f.project_id=p.id) nfiles,
                   (select count(*) from jobs j where j.project_id=p.id) njobs,
                   (select coalesce(sum(c.usd),0) from api_calls c where c.project_id=p.id and c.sample=0) ai_usd,
                   (select coalesce(sum(c.usd),0) from api_calls c where c.project_id=p.id and c.sample=1) sample_usd
                   from projects p where (? or exists (select 1 from project_members m
                     where m.project_id=p.id and m.user_id=?)) and p.archived=?
                   order by p.id desc""", (int(bool(u['is_admin'])), u['id'], int(archived)))   # AI cost: logged, live
    n_arch = db.q('select count(*) n from projects where archived=1', one=True)['n'] if u['is_admin'] else 0
    return page(request, 'projects.html', u=u, projects=rows, archived=archived, n_arch=n_arch)


@app.post('/projects')
def new_project(request: Request, name: str = Form(...), client: str = Form(''), short_name: str = Form('')):
    u = user(request)
    pid = db.x('insert into projects(name,client,short_name,created_by,created) values(?,?,?,?,?)',
               (name.strip(), client.strip(), naming.clean(short_name) or None, u['id'], db.now()))
    db.x('insert or ignore into project_members(project_id,user_id,added_by,added) values(?,?,?,?)',
         (pid, u['id'], u['id'], db.now()))
    return RedirectResponse(f'/projects/{pid}', 303)


@app.post('/projects/{pid}/naming')
def project_naming(request: Request, pid: int, short_name: str = Form(''), client: str = Form('')):
    """The client and short project name that go into output file names."""
    u = user(request)
    project_or_404(pid, u)
    db.x('update projects set short_name=?, client=? where id=?', (naming.clean(short_name) or None, client.strip(), pid))
    return RedirectResponse(f'/projects/{pid}', 303)


@app.post('/projects/{pid}/details')
def project_details(request: Request, pid: int, name: str = Form(''), client: str = Form(''), short_name: str = Form('')):
    """Project name, client and short name for files; anyone on the project (and admins) can change them.
    Files already made keep their names; the next output of each step uses the new ones."""
    u = user(request)
    p = project_or_404(pid, u)
    name = re.sub(r'\s+', ' ', name).strip()
    if len(name) < 2:
        raise HTTPException(400, 'Give the project a name')
    db.x('update projects set name=?, client=?, short_name=? where id=?',
         (name[:120], re.sub(r'\s+', ' ', client).strip()[:120], naming.clean(short_name) or None, pid))
    return RedirectResponse(f'/projects/{pid}?saved=1', 303)


def project_size(pid):
    """What deleting a project removes: its files, jobs and job folders, with the total size on disk."""
    jobs = [r['id'] for r in db.q('select id from jobs where project_id=?', (pid,))]
    nfiles = db.q('select count(*) n from files where project_id=?', (pid,), one=True)['n']
    size = 0
    for d in [db.DATA / 'projects' / str(pid)] + [worker.job_dir(j) for j in jobs]:
        if d.is_dir():
            size += sum(f.stat().st_size for f in d.rglob('*') if f.is_file())
    return dict(jobs=len(jobs), files=nfiles, mb=size / 1048576)


def busy_jobs(pid):
    return db.q(f"select id from jobs where project_id=? and status in ({','.join('?' * len(BUSY))})", (pid, *BUSY))


@app.post('/projects/{pid}/archive')
def project_archive(request: Request, pid: int):
    """Hide a finished project from every list and make it read-only; nothing is deleted."""
    me = admin(request)
    project_or_404(pid, me)
    if busy_jobs(pid):
        raise HTTPException(409, 'A job of this project is running: archive it when it has finished')
    db.x('update projects set archived=1 where id=?', (pid,))
    return RedirectResponse(f'/projects/{pid}', 303)


@app.post('/projects/{pid}/restore')
def project_restore(request: Request, pid: int):
    me = admin(request)
    project_or_404(pid, me)
    db.x('update projects set archived=0 where id=?', (pid,))
    return RedirectResponse(f'/projects/{pid}', 303)


@app.post('/projects/{pid}/delete')
def project_delete(request: Request, pid: int, confirm_name: str = Form('')):
    """Delete a project for good: its files, jobs, job folders and members. The name must be typed to confirm.
    Logged API costs stay in the cost history (their project shows as deleted)."""
    me = admin(request)
    p = project_or_404(pid, me)
    if confirm_name.strip() != p['name']:
        raise HTTPException(400, f'Type the project name exactly ({p["name"]}) to delete it')
    if busy_jobs(pid):
        raise HTTPException(409, 'A job of this project is running: delete it when it has finished')
    jobs = [r['id'] for r in db.q('select id from jobs where project_id=?', (pid,))]
    for d in [db.DATA / 'projects' / str(pid)] + [worker.job_dir(j) for j in jobs]:
        if d.is_dir() and db.DATA.resolve() in d.resolve().parents:
            shutil.rmtree(d, ignore_errors=True)
    for t in ('files', 'project_members', 'issued'):
        db.x(f'delete from {t} where project_id=?', (pid,))
    db.x('delete from jobs where project_id=?', (pid,))
    db.x('delete from projects where id=?', (pid,))
    return RedirectResponse('/projects?deleted=' + re.sub(r'[^\w .()-]', '', p['name'])[:60], 303)


@app.post('/projects/{pid}/members')
def project_members(request: Request, pid: int, action: str = Form(...), user_id: int = Form(...)):
    """Admins add or remove people on a project; members see only the projects they belong to."""
    me = admin(request)
    project_or_404(pid, me)
    t = db.q("select * from users where id=?", (user_id,), one=True)
    if not t:
        raise HTTPException(404)
    if action == 'add':
        db.x('insert or ignore into project_members(project_id,user_id,added_by,added) values(?,?,?,?)',
             (pid, user_id, me['id'], db.now()))
    elif action == 'remove':
        db.x('delete from project_members where project_id=? and user_id=?', (pid, user_id))
    return RedirectResponse(f'/projects/{pid}#members', 303)


@app.get('/account', response_class=HTMLResponse)
def account(request: Request):
    u = user(request)
    return page(request, 'account.html', u=u, example=naming.initials_of(u))


@app.post('/account')
def account_save(request: Request, initials: str = Form('')):
    """Your initials, used at the end of output file names (e.g. AAN)."""
    u = user(request)
    v = re.sub(r'[^A-Za-z]', '', initials).upper()[:6]
    db.x('update users set initials=? where id=?', (v or None, u['id']))
    return RedirectResponse('/account?saved=1', 303)


@app.get('/projects/{pid}', response_class=HTMLResponse)
def project(request: Request, pid: int):
    u = user(request)
    p = project_or_404(pid, u)
    order = list(FILE_KINDS)
    allf = db.q('select * from files where project_id=?', (pid,))
    # Files are browsed by folder: one folder per file kind (as stored, files/<kind>/). ?folder=<kind> opens one.
    folders = {}
    for f in allf:
        g = folders.setdefault(f['kind'], dict(kind=f['kind'], label=FILE_KINDS.get(f['kind'], f['kind']),
                                               count=0, size=0, last=''))
        g['count'] += 1; g['size'] += f['size'] or 0; g['last'] = max(g['last'], f['created'])
    folders = sorted(folders.values(), key=lambda g: order.index(g['kind']) if g['kind'] in order else 99)
    folder = request.query_params.get('folder', '')
    if folder not in [g['kind'] for g in folders]:
        folder = ''
    files = sorted((f for f in allf if f['kind'] == folder), key=lambda f: f['name'].lower()) if folder else []
    jobs = db.q('select * from jobs where project_id=? order by id desc', (pid,))
    return page(request, 'project.html', u=u, p=p, files=files, folders=folders, folder=folder, nfiles=len(allf),
                jobs=jobs, chainst=chain.status(pid), size=project_size(pid) if u['is_admin'] else None,
                creator=(db.q('select name from users where id=?', (p['created_by'],), one=True) or {'name': ''})['name'],
                members=db.q("""select u.id, u.name, u.email, u.is_admin from project_members m join users u on u.id=m.user_id
                                where m.project_id=? order by lower(u.name)""", (pid,)),
                others=db.q("""select id, name, email, is_admin from users where status='active' and id not in
                               (select user_id from project_members where project_id=?) order by lower(name)""", (pid,))
                       if u['is_admin'] else [],
                name_example=naming.file_name(dict(date=datetime.now().strftime('%Y.%m.%d'), client=naming.clean(p['client']),
                                                   project=naming.clean(p['short_name'] or p['name']),
                                                   initials=naming.initials_of(u) or 'XX'),
                                              naming.TASK_MASTER, 0, '.xlsx'))


@app.post('/projects/{pid}/files')
async def upload_files(request: Request, pid: int, kind: str = Form('drawing'), files: List[UploadFile] = File(...)):
    u = user(request)
    project_or_404(pid, u)
    if kind not in UPLOAD_KINDS:
        raise HTTPException(400, 'Unknown file kind')
    for up in files:
        if not up.filename:
            continue
        name = safe_name(up.filename)
        old = db.q('select id,path from files where project_id=? and name=? and kind=?', (pid, name, kind), one=True)
        dest = db.DATA / old['path'] if old else db.DATA / 'projects' / str(pid) / 'files' / kind / name
        size = await save_upload(up, dest)
        if old:   # same name and kind uploaded again: replace it
            db.x('update files set size=?, uploaded_by=?, created=?, job_id=null where id=?', (size, u['id'], db.now(), old['id']))
        else:
            db.x('insert into files(project_id,kind,name,path,size,uploaded_by,created) values(?,?,?,?,?,?,?)',
                 (pid, kind, name, str(dest.relative_to(db.DATA)), size, u['id'], db.now()))
    return RedirectResponse(f'/projects/{pid}?folder={kind}#files', 303)    # open the folder just uploaded to


@app.get('/projects/{pid}/files/{fid}')
def get_file(request: Request, pid: int, fid: int):
    u = user(request)
    project_or_404(pid, u)
    f = db.q('select * from files where id=? and project_id=?', (fid, pid), one=True)
    if not f:
        raise HTTPException(404)
    return FileResponse(inside(db.DATA, f['path']), filename=f['name'])


# ---------------- jobs ----------------
@app.post('/projects/{pid}/jobs')
def new_job(request: Request, pid: int, module: str = Form(...), title: str = Form('')):
    u = user(request)
    project_or_404(pid, u)
    if module not in MODULES or not MODULES[module].get('ready'):
        raise HTTPException(400, 'This module is not available yet')
    n = db.q('select count(*) n from jobs where project_id=? and module=?', (pid, module), one=True)['n']
    rev = f'R{n + 1:02d}'
    step = MODULES[module].get('runner') == 'step'
    jid = db.x('insert into jobs(project_id,module,title,status,rev,created_by,created) values(?,?,?,?,?,?,?)',
               (pid, module, title.strip() or MODULES[module]['name'], 'setup' if step else 'awaiting_spec', rev, u['id'], db.now()))
    (worker.job_dir(jid) / ('in' if step else 'spec')).mkdir(parents=True, exist_ok=True)
    return RedirectResponse(f'/jobs/{jid}', 303)


def job_files(jid):
    d = worker.job_dir(jid)
    spec = sorted(p.name for p in (d / 'spec').glob('*') if p.is_file()) if (d / 'spec').exists() else []
    out = sorted(str(p.relative_to(d / 'out')).replace('\\', '/') for p in (d / 'out').rglob('*')
                 if p.is_file() and p.suffix in ('.xlsx', '.pptx', '.html', '.zip', '.png', '.json')) if (d / 'out').exists() else []
    return spec, out


BUSY = ('checking', 'queued', 'running')


def jread(p: Path, default=None):
    return json.loads(p.read_text(encoding='utf-8')) if p.exists() else default


def recheck(jid):
    """Spec or review edits change the numbers: throw the old check away and queue a fresh one."""
    d = worker.job_dir(jid)
    shutil.rmtree(d / 'check', ignore_errors=True)
    has_spec = any((d / 'spec').glob('*')) if (d / 'spec').exists() else False
    db.x('update jobs set status=?, message=null where id=?', ('checking' if has_spec else 'awaiting_spec', jid))


def not_busy(j):
    if j['status'] in BUSY:
        raise HTTPException(409, 'The job is busy - wait for it to finish')


@app.get('/jobs/{jid}', response_class=HTMLResponse)
def job(request: Request, jid: int):
    u = user(request)
    j = job_or_404(jid, u)
    p = project_or_404(j['project_id'], u)
    if MODULES[j['module']].get('runner') == 'step':
        return step_job_page(request, u, j, p)
    d = worker.job_dir(jid)
    spec, out = job_files(jid)
    built = jread(d / 'out' / 'manifest.json')
    if built and built.get('stage') != 'build':
        built = None
    review = jread(d / 'review.json')
    if review:
        r = db.q('select name from users where id=?', (review.get('approved_by'),), one=True)
        review['approver'] = r['name'] if r else '?'
    return page(request, 'job.html', u=u, j=j, p=p, m=MODULES[j['module']], spec=spec, out=out, allow_py=py_ok(u),
                check=jread(d / 'check' / 'manifest.json'), checks=jread(d / 'check' / 'checks.json', []),
                items=jread(d / 'check' / 'review_items.json', []), draft=jread(d / 'review_draft.json', {}),
                built=built, built_checks=jread(d / 'out' / 'checks.json', []), review=review, busy=j['status'] in BUSY)


def py_ok(u):
    """A .py spec is run by the pipeline on this laptop, so only admins may upload one; others upload JSON."""
    return ALLOW_PY_SPECS and bool(u['is_admin'])


@app.post('/jobs/{jid}/spec')
async def upload_spec(request: Request, jid: int, files: List[UploadFile] = File(...)):
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    ok = tuple(t for t in MODULES[j['module']]['spec_types'] if py_ok(u) or t != '.py')
    for up in files:
        name = safe_name(up.filename)
        if Path(name).suffix.lower() not in ok:
            raise HTTPException(400, f'{name}: spec must be {" or ".join(ok)}')
        await save_upload(up, worker.job_dir(jid) / 'spec' / name)
    recheck(jid)
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/spec/{name}/delete')
def delete_spec(request: Request, jid: int, name: str):
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    inside(worker.job_dir(jid) / 'spec', safe_name(name)).unlink()
    recheck(jid)
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/recheck')
def recheck_job(request: Request, jid: int):
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    recheck(jid)
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/edits')
async def review_edits(request: Request, jid: int):
    """Asset names and the atrium counting floor: saved as review_draft.json, then re-checked (numbers may change)."""
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    d = worker.job_dir(jid)
    items = jread(d / 'check' / 'review_items.json', [])
    form = await request.form()
    draft = {'assets': {}}
    for it in items:
        s = it['spec']
        a = {}
        name = re.sub(r"[\[\]:*?/\\']+", ' ', str(form.get(f'name::{s}', ''))).strip()[:31]   # Excel sheet name rules
        if name:
            a['name'] = name
        lvl = str(form.get(f'atrium::{s}', ''))
        if it.get('atrium') and lvl in [L['id'] for L in it['atrium']['levels']]:
            a['atrium_level'] = lvl
        if a:
            draft['assets'][s] = a
    (d / 'review_draft.json').write_text(json.dumps(draft, indent=1, ensure_ascii=False), encoding='utf-8')
    recheck(jid)
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/approve')
async def approve(request: Request, jid: int):
    u = user(request)
    j = job_or_404(jid, u)
    if j['status'] != 'review':
        raise HTTPException(409, 'Only a checked job waiting for review can be approved')
    d = worker.job_dir(jid)
    items = jread(d / 'check' / 'review_items.json', [])
    checks = jread(d / 'check' / 'checks.json', [])
    form = await request.form()
    need = ['names', 'a1'] + [f"atrium::{it['spec']}" for it in items if it.get('atrium')] + \
           [c['key'] for it in items for c in it['calls']]
    missing = [k for k in need if form.get(f'ok::{k}') != 'on']
    if missing:
        raise HTTPException(400, f'Tick every point before approving ({len(missing)} left)')
    hard = [c for c in checks if not c['ok'] and not c.get('warn_only')]
    reason = str(form.get('override_reason', '')).strip()
    if hard and len(reason) < 10:
        raise HTTPException(400, f'{len(hard)} check(s) failed: give a reason (10+ characters) to approve anyway')
    # Spending limit: Stage A runs cost nothing, but the gate is in place for Stage B (API calls).
    if db.budget_left() < 0:
        raise HTTPException(402, 'Monthly spending limit reached - ask an admin to raise it on the Costs page')
    draft = jread(d / 'review_draft.json', {})
    confirmed = ['Asset names: ' + ', '.join(it['asset'] for it in items),
                 'Notional boundary areas (A1) are excluded from AAA BUA, although report 2.11 counts row A.']
    for it in items:
        if it.get('atrium'):
            lv = ', '.join(it['atrium']['counted']) or 'none'
            confirmed.append(f"{it['asset']}: atrium void counted on {lv}, removed from the other floors.")
        for c in it['calls']:
            confirmed.append(f"{it['asset']} {c['level']}: {c['name']} ({c['area']:.2f} m²). {c['rule']}")
    review = dict(assets=draft.get('assets', {}), confirmed=confirmed, override_reason=reason if hard else None,
                  failed_checks=[f"{c['asset']} {c['level']}: {c['check']} = {c['value']}{c['unit']}" for c in hard],
                  approved_by=u['id'], approved_at=db.now())
    rev = j['rev']
    prev = jread(d / 'out' / 'manifest.json')
    if prev and prev.get('stage') == 'build':      # outputs of this revision may already be with someone: next revision
        rev = f"R{int(rev[1:]) + 1:02d}"
    (d / 'review.json').write_text(json.dumps(review, indent=1, ensure_ascii=False), encoding='utf-8')
    shutil.rmtree(d / 'out', ignore_errors=True)
    db.x("update jobs set status='queued', rev=?, message=null, started=null, finished=null, run_by=? where id=?",
         (rev, u['id'], jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.get('/jobs/{jid}/log', response_class=PlainTextResponse)
def job_log(request: Request, jid: int, which: str = 'run'):
    u = user(request)
    job_or_404(jid, u)
    f = worker.job_dir(jid) / ('check.log' if which == 'check' else 'run.log')
    return f.read_text(encoding='utf-8', errors='ignore') if f.exists() else 'Nothing yet.'


@app.get('/jobs/{jid}/check/{rel:path}')
def job_check_file(request: Request, jid: int, rel: str):
    u = user(request)
    job_or_404(jid, u)
    rel = rel.replace('\\', '/')
    if rel.startswith('check/'):       # the pipeline stores preview paths relative to the job ("check\preview\x.html")
        rel = rel[len('check/'):]
    f = inside(worker.job_dir(jid) / 'check', rel)
    return FileResponse(f, media_type=mimetypes.guess_type(f.name)[0])


@app.get('/jobs/{jid}/out/{rel:path}')
def job_out(request: Request, jid: int, rel: str, view: int = 0):
    u = user(request)
    job_or_404(jid, u)
    f = inside(worker.job_dir(jid) / 'out', rel)
    if view and f.suffix in ('.html', '.png', '.json'):
        return FileResponse(f, media_type=mimetypes.guess_type(f.name)[0])
    return FileResponse(f, filename=f.name)


# ---------------- step modules (PTE Phase 1) ----------------
def step_candidates(j):
    """Project files a step can take, grouped by kind, and the ones ticked by default."""
    m = MODULES[j['module']]
    kinds = [k for grp in m['inputs'] for k in grp]
    files = db.q('select * from files where project_id=? order by name', (j['project_id'],))
    groups = [(k, FILE_KINDS.get(k, k), [f for f in files if f['kind'] == k]) for k in kinds]
    have = {k for k, _, fs in groups if fs}
    default = set()
    for grp in m['inputs']:
        k = next((k for k in grp if k in have), None)
        if k:
            default |= {f['id'] for f in files if f['kind'] == k}
    return [g for g in groups if g[2]], default


def clean_label(v):
    """Asset / package names become Excel sheet names later: no []:*?/\\ and at most 31 characters."""
    v = re.sub(r"[\[\]:*?/\\']+", ' ', str(v or ''))
    return re.sub(r'\s+', ' ', v).strip()[:31]


def name_problems(rows):
    out = []
    for r in rows:
        if not r['asset']:
            out.append(f"{r['file']}: no asset")
        if not r['package']:
            out.append(f"{r['file']}: no package")
    return out


WHOLE = 'Whole'      # same as toolkit/pte/names.WHOLE: the code of a BOQ that covers a whole asset


def is_whole(package):
    return re.sub(r' -\d+$', '', (package or '').strip()).lower() == WHOLE.lower()


def with_scope(rows):
    """Register review rows carry scope: 'whole' (one BOQ for the whole asset, code "Whole") or 'package'."""
    return [dict(r, scope=r.get('scope') or ('whole' if is_whole(r['package']) else 'package')) for r in rows]


def scope_warnings(rows):
    """Not blocking: an asset with a whole-asset BOQ and package BOQs may price the same work twice."""
    out, by = [], {}
    for r in rows:
        by.setdefault(r['asset'], []).append(r)
    for a, rs in by.items():
        whole = [r for r in rs if r['scope'] == 'whole']
        if whole and len(whole) < len(rs):
            out.append(f"{a}: a whole-asset BOQ and package BOQs ("
                       f"{', '.join(r['package'] or '?' for r in rs if r['scope'] != 'whole')}). "
                       f"Check they do not price the same work twice.")
    return out


def number_duplicates(rows):
    """Same as toolkit/pte/names.number_duplicates: two BOQs of one asset with one code -> "Arch -1", "Arch -2"."""
    groups = {}
    for r in rows:
        if r['package']:
            groups.setdefault((r['asset'].upper(), r['package'].upper()), []).append(r)
    for rs in groups.values():
        if len(rs) > 1:
            base = rs[0]['package']
            for i, r in enumerate(rs, 1):
                r['package'] = f'{base} -{i}'
    return rows


def master_problems(assets, check):
    """06-04 review: names become tab names and are matched across five sheets."""
    out, seen = [], set()
    other = {s.lower() for s in check.get('other_sheets', [])}
    if not assets:
        out.append('No assets.')
    if len(assets) > len(check.get('tabs', [])):
        out.append(f"{len(assets)} assets, but the master has only {len(check.get('tabs', []))} asset tabs.")
    for a in assets:
        n = a['name']
        if n.lower() in seen:
            out.append(f'"{n}" is used twice.')
        if n.lower() in other:
            out.append(f'"{n}" is already the name of another sheet.')
        seen.add(n.lower())
    return out


def step_job_page(request, u, j, p):
    d = worker.job_dir(j['id'])
    groups, default = step_candidates(j)
    job_json = jread(d / 'job.json')
    chosen = {f.get('file_id') for f in job_json['inputs']} if job_json else default
    check = jread(d / 'check' / 'manifest.json')
    prev = jread(d / 'review_prev.json', {})       # the last approved review, when the step is run again
    rows = None
    if check and check.get('items') is not None and j['module'] == 'name-assets' and j['status'] == 'review':
        draft = jread(d / 'review_draft.json', {})
        rows = draft.get('assets', check['items'])  # names always come from the (possibly edited) register
        check = dict(check, options=draft.get('options') or prev.get('options') or check['options'])
        return page(request, 'job_step.html', u=u, j=j, p=p, m=MODULES[j['module']], groups=groups, chosen=chosen,
                    check=check, rows=rows, problems=master_problems(rows, check), built=None, extra=[],
                    review=None, busy=j['status'] in BUSY)
    if check and check.get('divmap') is not None and j['module'] == 'load-boq' and j['status'] == 'review':
        draft = jread(d / 'review_draft.json', {}).get('divmap') or prev.get('divmap', {})   # previous mapping kept
        dm = [dict(x, num=draft.get(x['key'], x['num'])) for x in check['divmap']]
        nums = {m['num'] for m in check.get('master_divs', [])}
        probs = [f"{x['package']} / {x['label']}: choose a master division" for x in dm if x['num'] is None] + \
                [f"{x['package']} / {x['label']}: Division {x['num']} is not in the master" for x in dm
                 if x['num'] is not None and nums and x['num'] not in nums]
        return page(request, 'job_step.html', u=u, j=j, p=p, m=MODULES[j['module']], groups=groups, chosen=chosen,
                    check=dict(check, divmap=dm), rows=check['items'], problems=probs, built=None, extra=[],
                    review=None, busy=j['status'] in BUSY)
    if check and check.get('items') is not None and j['module'] == 'asset-names':
        rows = with_scope(jread(d / 'review_draft.json', {}).get('rows') or check['items'])
    built = jread(d / 'out' / 'manifest.json')
    extra = []
    if built:
        listed = {x['file'] for x in built.get('publish', [])}
        extra = [x.name for x in sorted((d / 'out').iterdir())        # zips and reports not added to the project
                 if x.suffix in ('.zip', '.xlsx') and x.name not in listed]
    st = chain.status(j['project_id'])
    order = list(MODULES)
    later = [x for x in (db.q('select * from jobs where id=?', (k,), one=True) for k, v in st.items()
                         if k != j['id'] and v['state'] in ('stale', 'waits'))
             if order.index(x['module']) > order.index(j['module'])]       # steps after this one, in step order
    later.sort(key=lambda x: order.index(x['module']))
    return page(request, 'job_step.html', u=u, j=j, p=p, m=MODULES[j['module']], groups=groups, chosen=chosen,
                check=check, rows=rows, problems=name_problems(rows) if rows else [], built=built, extra=extra,
                warnings=scope_warnings(rows) if rows and j['module'] == 'asset-names' else [],
                review=jread(d / 'review.json'), busy=j['status'] in BUSY,
                chainst=st.get(j['id'], dict(state='n/a', reasons=[])), later=later)


@app.post('/jobs/{jid}/run')
async def step_run(request: Request, jid: int):
    """Save the chosen input files and start the step (or its check, for steps with a review)."""
    u = user(request)
    j = job_or_404(jid, u)
    m = MODULES[j['module']]
    if m.get('runner') != 'step' or j['status'] != 'setup':
        raise HTTPException(409, 'This job is not waiting for its input files')
    form = await request.form()
    ids = {int(k[3:]) for k in form if k.startswith('f::') and k[3:].isdigit()}
    kinds = {k for grp in m['inputs'] for k in grp}
    files = [f for f in db.q('select * from files where project_id=?', (j['project_id'],)) if f['id'] in ids and f['kind'] in kinds]
    if not files:
        raise HTTPException(400, 'Tick at least one file')
    chain.start_step(j, files)
    db.x('update jobs set run_by=? where id=?', (u['id'], jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/chain')
def step_chain(request: Request, jid: int):
    """"Update later steps": re-run this step if out of date, then the out-of-date steps after it, in order."""
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    if MODULES[j['module']].get('runner') != 'step' or j['status'] != 'done':
        raise HTTPException(409, 'Only a finished step can update the steps after it')
    try:
        started = chain.start_chain(j)
        for k in started:
            db.x('update jobs set run_by=? where id=?', (u['id'], k))
    except ValueError as e:
        raise HTTPException(400, str(e))
    first = started[0] if started else jid
    return RedirectResponse(f'/jobs/{first}', 303)


@app.post('/jobs/{jid}/names')
async def step_names(request: Request, jid: int):
    """Asset / package review: save the edits, or approve and build the register."""
    u = user(request)
    j = job_or_404(jid, u)
    d = worker.job_dir(jid)
    if j['module'] == 'asset-names' and j['status'] == 'done' and can_reopen(d):
        j = reopen_register(j, d)      # edits sent from a page opened before the register was finished: keep them
    if j['module'] != 'asset-names' or j['status'] != 'review':
        raise HTTPException(409, f"This register is {j['status']} and cannot take edits now; reload the page")
    check = jread(d / 'check' / 'manifest.json', {})
    base = with_scope(jread(d / 'review_draft.json', {}).get('rows') or check.get('items', []))
    form = await request.form()
    at = {str(form.get(k)): k.split('::', 1)[1] for k in form if k.startswith('file::')}   # BOQ file -> its form row
    rows = []
    for n, r in enumerate(base):
        i = at.get(r['file'], n)
        scope = 'whole' if form.get(f'scope::{i}', r['scope']) == 'whole' else 'package'
        typed = form.get(f'package::{i}')
        pkg = clean_label(typed if typed is not None else r['package'])
        if scope == 'whole':
            pkg = WHOLE if not is_whole(pkg) else pkg          # keeps "Whole -1" numbering
        elif is_whole(pkg):
            if r['scope'] == 'package':        # "Whole" typed as the code of a package BOQ: the whole-asset choice
                scope = 'whole'
            else:                              # switched from whole to package: the code is still to be typed
                pkg = ''
        rows.append(dict(r, asset=clean_label(form.get(f'asset::{i}', r['asset'])), package=pkg, scope=scope))
    rows = number_duplicates(rows)
    draft = dict(jread(d / 'review_draft.json', {}), rows=rows)     # keeps done_message while editing a register
    (d / 'review_draft.json').write_text(json.dumps(draft, indent=1, ensure_ascii=False), encoding='utf-8')
    if form.get('action') != 'approve':
        return RedirectResponse(f'/jobs/{jid}#review', 303)
    probs = name_problems(rows)
    if probs:
        raise HTTPException(400, 'Fix these first: ' + '; '.join(probs))
    review = dict(rows=[dict(file=r['file'], asset=r['asset'], package=r['package'], scope=r['scope']) for r in rows],
                  approved_by=u['id'], approved_at=db.now())
    (d / 'review.json').write_text(json.dumps(review, indent=1, ensure_ascii=False), encoding='utf-8')
    rev = j['rev']
    prev = jread(d / 'out' / 'manifest.json')
    if prev and prev.get('stage') == 'build':      # an edited register is the next revision
        rev = f"R{int(rev[1:]) + 1:02d}"
    db.x("update jobs set status='queued', rev=?, message=null, started=null, finished=null, run_by=? where id=?",
         (rev, u['id'], jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/names/edit')
def step_names_edit(request: Request, jid: int):
    """Reopen a finished register for editing, starting from the approved choices; approving builds the next revision."""
    u = user(request)
    j = job_or_404(jid, u)
    d = worker.job_dir(jid)
    if j['module'] != 'asset-names' or j['status'] != 'done' or not can_reopen(d):
        raise HTTPException(409, 'Only a finished register can be edited')
    reopen_register(j, d)
    return RedirectResponse(f'/jobs/{jid}#review', 303)


def can_reopen(d):
    return (d / 'review.json').exists() and (d / 'check' / 'manifest.json').exists()


def reopen_register(j, d):
    """Back to the review screen, starting from the approved choices; returns the updated job row."""
    how = {x['file']: x.get('how', '') for x in jread(d / 'check' / 'manifest.json', {}).get('items', [])}
    rows = with_scope([dict(r, how=how.get(r['file'], '')) for r in jread(d / 'review.json')['rows']])
    (d / 'review_draft.json').write_text(json.dumps(dict(rows=rows, done_message=j['message']), indent=1,
                                                    ensure_ascii=False), encoding='utf-8')
    db.x("update jobs set status='review', message=? where id=?", (f"editing the {j['rev']} register", j['id']))
    return db.q('select * from jobs where id=?', (j['id'],), one=True)


@app.post('/jobs/{jid}/names/cancel')
def step_names_cancel(request: Request, jid: int):
    """Stop editing a finished register: back to it as it was, nothing rebuilt."""
    u = user(request)
    j = job_or_404(jid, u)
    d = worker.job_dir(jid)
    built = jread(d / 'out' / 'manifest.json')
    if j['module'] != 'asset-names' or j['status'] != 'review' or not (built and built.get('stage') == 'build'):
        raise HTTPException(409, 'This register is not being edited')
    msg = jread(d / 'review_draft.json', {}).get('done_message') or built.get('summary')
    (d / 'review_draft.json').unlink(missing_ok=True)
    db.x("update jobs set status='done', message=? where id=?", (msg, jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/master')
async def step_master(request: Request, jid: int):
    """06-04 review: asset names, packages and options; save, or approve and build."""
    u = user(request)
    j = job_or_404(jid, u)
    if j['module'] != 'name-assets' or j['status'] != 'review':
        raise HTTPException(409, 'This job is not waiting for a review')
    d = worker.job_dir(jid)
    check = jread(d / 'check' / 'manifest.json', {})
    form = await request.form()
    assets = []
    for i in range(int(form.get('count', 0))):
        name = clean_label(form.get(f'name::{i}', ''))
        if not name:                                   # a blank name drops the row
            continue
        codes = check.get('items', [])[i]['packages'] if i < len(check.get('items', [])) else []
        assets.append(dict(name=name, packages=codes))       # codes shown only: 06-04 builds one block per asset
    options = dict(delete_spare=form.get('spare') != 'keep', amount='est' if form.get('amount') == 'est' else 'pte')
    (d / 'review_draft.json').write_text(json.dumps(dict(assets=assets, options=options), indent=1, ensure_ascii=False),
                                         encoding='utf-8')
    if form.get('action') != 'approve':
        return RedirectResponse(f'/jobs/{jid}#review', 303)
    probs = master_problems(assets, check)
    if probs:
        raise HTTPException(400, 'Fix these first: ' + '; '.join(probs))
    review = dict(assets=assets, options=options, approved_by=u['id'], approved_at=db.now())
    (d / 'review.json').write_text(json.dumps(review, indent=1, ensure_ascii=False), encoding='utf-8')
    db.x("update jobs set status='queued', message=null, started=null, finished=null, run_by=? where id=?",
         (u['id'], jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/load')
async def step_load(request: Request, jid: int):
    """06-06 review: the master division for each BOQ division; save, or approve and load."""
    u = user(request)
    j = job_or_404(jid, u)
    if j['module'] != 'load-boq' or j['status'] != 'review':
        raise HTTPException(409, 'This job is not waiting for a review')
    d = worker.job_dir(jid)
    check = jread(d / 'check' / 'manifest.json', {})
    if check.get('checks'):
        raise HTTPException(400, 'Fix the problems listed first (they need other input files)')
    form = await request.form()
    divmap = {}
    for i, x in enumerate(check.get('divmap', [])):
        v = str(form.get(f'div::{i}', '')).strip()
        divmap[x['key']] = int(v) if v.isdigit() else None
    (d / 'review_draft.json').write_text(json.dumps(dict(divmap=divmap), indent=1, ensure_ascii=False), encoding='utf-8')
    if form.get('action') != 'approve':
        return RedirectResponse(f'/jobs/{jid}#review', 303)
    nums = {m['num'] for m in check.get('master_divs', [])}
    bad = [k for k, v in divmap.items() if v is None or (nums and v not in nums)]
    if bad:
        raise HTTPException(400, 'Choose a master division for: ' + '; '.join(bad))
    review = dict(divmap=divmap, approved_by=u['id'], approved_at=db.now())
    (d / 'review.json').write_text(json.dumps(review, indent=1, ensure_ascii=False), encoding='utf-8')
    db.x("update jobs set status='queued', message=null, started=null, finished=null, run_by=? where id=?",
         (u['id'], jid))
    return RedirectResponse(f'/jobs/{jid}', 303)


@app.post('/jobs/{jid}/setup')
def step_setup(request: Request, jid: int):
    """Back to choosing the input files (the last run's results stay until the next run finishes)."""
    u = user(request)
    j = job_or_404(jid, u)
    not_busy(j)
    if MODULES[j['module']].get('runner') != 'step':
        raise HTTPException(409)
    d = worker.job_dir(jid)
    shutil.rmtree(d / 'check', ignore_errors=True)
    (d / 'review_draft.json').unlink(missing_ok=True)
    db.x("update jobs set status='setup', message=null where id=?", (jid,))
    return RedirectResponse(f'/jobs/{jid}', 303)


# ---------------- costs ----------------
@app.get('/costs', response_class=HTMLResponse)
def costs(request: Request):
    """Claude API cost tracker, admins only: totals by person, task, project and month, plus every call."""
    u = admin(request)
    f = cost_filters(request)
    rows = cost_rows(f)
    groups = {}
    for title, key in (('By person', 'uname'), ('By task', 'task'), ('By project', 'pname'), ('By month', 'month')):
        g = {}
        for r in rows:
            t = g.setdefault(r[key] or '(none)', dict(name=r[key] or '(none)', calls=0, input=0, output=0,
                                                      cache_read=0, cache_write=0, usd=0.0))
            t['calls'] += 1
            t['input'] += r['input_tokens']; t['output'] += r['output_tokens']
            t['cache_read'] += r['cache_read_tokens']; t['cache_write'] += r['cache_write_tokens']
            t['usd'] += r['usd']
        order = sorted(g.values(), key=lambda t: t['name'], reverse=True) if key == 'month' else \
            sorted(g.values(), key=lambda t: -t['usd'])
        groups[title] = order
    total = dict(calls=len(rows), usd=sum(r['usd'] for r in rows),
                 tokens=sum(r['input_tokens'] + r['output_tokens'] + r['cache_read_tokens'] + r['cache_write_tokens']
                            for r in rows))
    users = db.q('select id, name from users order by name')
    qs = '&'.join(f'{k}={v}' for k, v in f.items() if v)
    return page(request, 'costs.html', u=u, rows=rows, groups=groups, total=total, f=f, users=users, qs=qs,
                samples=sum(1 for r in rows if r['sample']), spend=db.month_spend(),
                limit=float(db.setting('monthly_limit_usd', '0')), month=db.now()[:7])


def cost_filters(request: Request) -> dict:
    """date_from / date_to (YYYY-MM-DD, both days included) and user id, from the query string."""
    p = request.query_params
    d = lambda v: v if re.fullmatch(r'\d{4}-\d{2}-\d{2}', v or '') else ''
    uid = p.get('user', '')
    return dict(date_from=d(p.get('date_from')), date_to=d(p.get('date_to')), user=uid if uid.isdigit() else '')


def cost_rows(f: dict):
    where, args = [], []
    if f['date_from']:
        where.append('substr(c.ts,1,10) >= ?'); args.append(f['date_from'])
    if f['date_to']:
        where.append('substr(c.ts,1,10) <= ?'); args.append(f['date_to'])
    if f['user']:
        where.append('c.user_id = ?'); args.append(int(f['user']))
    rows = db.q(f"""select c.*, u.name uname, p.name pname from api_calls c left join users u on u.id=c.user_id
                    left join projects p on p.id=c.project_id {'where ' + ' and '.join(where) if where else ''}
                    order by c.ts desc, c.id desc""", args)
    out = []
    for r in rows:
        r = dict(r)
        r['task'] = MODULES[r['module']]['name'] if r['module'] in MODULES else (r['module'] or '')
        r['month'] = r['ts'][:7]
        out.append(r)
    return out


@app.get('/costs.csv')
def costs_csv(request: Request):
    admin(request)
    import csv, io
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(['Date/time (UTC)', 'User', 'Project', 'Task / module', 'Model', 'Input tokens', 'Output tokens',
                'Cache read tokens', 'Cache write tokens', 'Cost (USD)', 'Job ID', 'Sample', 'Note'])
    for r in cost_rows(cost_filters(request)):
        w.writerow([r['ts'], r['uname'] or '', r['pname'] or '', r['task'], r['model'], r['input_tokens'],
                    r['output_tokens'], r['cache_read_tokens'], r['cache_write_tokens'], f"{r['usd']:.6f}",
                    r['job_id'] or '', 'yes' if r['sample'] else '', r['note'] or ''])
    name = f"claude-api-costs-{db.now()[:10]}.csv"
    return PlainTextResponse('﻿' + buf.getvalue(), media_type='text/csv; charset=utf-8',   # BOM: Excel reads UTF-8
                             headers={'Content-Disposition': f'attachment; filename="{name}"'})


@app.post('/costs/samples')
def cost_samples(request: Request, action: str = Form(...)):
    """Sample entries to try the page before the API is connected; one button adds them, one deletes them."""
    admin(request)
    from . import costs as costs_mod
    if action == 'add':
        costs_mod.add_samples()
    elif action == 'delete':
        costs_mod.delete_samples()
    return RedirectResponse('/costs', 303)


@app.post('/costs/limit')
def set_limit(request: Request, limit: float = Form(...)):
    admin(request)
    if limit < 0:
        raise HTTPException(400, 'Limit must be 0 or more')
    db.set_setting('monthly_limit_usd', f'{limit:.2f}')
    return RedirectResponse('/costs', 303)
