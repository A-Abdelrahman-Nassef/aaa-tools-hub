"""db.py - SQLite store for the hub (users, projects, files, jobs, cost log, settings)."""
import hashlib, hmac, os, secrets, sqlite3, threading
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(os.environ.get('HUB_DATA', ROOT / 'data'))
DB_PATH = DATA / 'hub.db'
_lock = threading.Lock()

SCHEMA = """
create table if not exists users (id integer primary key, email text unique not null, name text not null,
  pw_hash text not null, is_admin integer not null default 0, created text not null);
create table if not exists projects (id integer primary key, name text not null, client text, created_by integer,
  created text not null);
create table if not exists files (id integer primary key, project_id integer not null, kind text not null, name text not null,
  path text not null, size integer, uploaded_by integer, created text not null);
create table if not exists jobs (id integer primary key, project_id integer not null, module text not null, title text,
  status text not null, rev text not null default 'R01', message text, created_by integer, created text not null,
  started text, finished text);
create table if not exists cost_log (id integer primary key, ts text not null, user_id integer, project_id integer,
  job_id integer, module text, source text not null, input_tokens integer default 0, output_tokens integer default 0,
  usd real not null default 0, note text);
create table if not exists settings (key text primary key, value text);
create table if not exists api_calls (id integer primary key, ts text not null, user_id integer, project_id integer,
  job_id integer, module text, model text not null, input_tokens integer not null default 0,
  output_tokens integer not null default 0, cache_read_tokens integer not null default 0,
  cache_write_tokens integer not null default 0, usd real not null default 0, sample integer not null default 0,
  note text);
"""


def now():
    return datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')


def connect():
    DATA.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DB_PATH, check_same_thread=False)
    c.row_factory = sqlite3.Row
    c.execute('pragma foreign_keys=on')
    return c


def init():
    with connect() as c:
        c.executescript(SCHEMA)
        c.execute("insert or ignore into settings(key,value) values('monthly_limit_usd','10')")
        cols = [r[1] for r in c.execute('pragma table_info(files)')]
        if 'job_id' not in cols:      # step 6: files a step module put back into the project
            c.execute('alter table files add column job_id integer')
        jcols = [r[1] for r in c.execute('pragma table_info(jobs)')]
        if 'chain' not in jcols:      # chain.py: re-run later steps after an earlier one changed
            c.execute('alter table jobs add column chain integer not null default 0')
        if 'run_by' not in jcols:     # naming.py: whose initials go in the output file names
            c.execute('alter table jobs add column run_by integer')
        ucols = [r[1] for r in c.execute('pragma table_info(users)')]
        if 'initials' not in ucols:
            c.execute('alter table users add column initials text')
        if 'status' not in ucols:     # sign-up: pending -> active (approved) | rejected; active -> inactive (deactivated)
            c.execute("alter table users add column status text not null default 'active'")
        c.execute("insert or ignore into settings(key,value) values('signup_domain','aaa-consultancy.com')")
        # project access: members see only their projects (admins see all); the creator is always a member
        c.execute("""create table if not exists project_members (project_id integer not null, user_id integer not null,
                     added_by integer, added text not null, primary key (project_id, user_id))""")
        c.execute("""insert or ignore into project_members(project_id, user_id, added_by, added)
                     select id, created_by, created_by, created from projects where created_by is not null""")
        pcols = [r[1] for r in c.execute('pragma table_info(projects)')]
        if 'short_name' not in pcols:
            c.execute('alter table projects add column short_name text')
        if 'archived' not in pcols:   # archived: hidden from lists, read-only, admins can restore or delete it
            c.execute('alter table projects add column archived integer not null default 0')
        c.execute("""create table if not exists issued (id integer primary key, project_id integer not null,
                     task text not null, rev integer not null, name text not null, job_id integer, created text not null)""")


def q(sql, args=(), one=False):
    with _lock, connect() as c:
        rows = c.execute(sql, args).fetchall()
    return (rows[0] if rows else None) if one else rows


def x(sql, args=()):
    with _lock, connect() as c:
        cur = c.execute(sql, args)
        return cur.lastrowid


# ---------- passwords (stdlib PBKDF2, no extra package) ----------
def hash_pw(pw: str) -> str:
    salt = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac('sha256', pw.encode(), salt.encode(), 310_000).hex()
    return f'pbkdf2${salt}${h}'


def check_pw(pw: str, stored: str) -> bool:
    try:
        _, salt, h = stored.split('$')
    except ValueError:
        return False
    return hmac.compare_digest(hashlib.pbkdf2_hmac('sha256', pw.encode(), salt.encode(), 310_000).hex(), h)


# ---------- settings and costs ----------
def setting(key, default=None):
    r = q('select value from settings where key=?', (key,), one=True)
    return r['value'] if r else default


def set_setting(key, value):
    x('insert into settings(key,value) values(?,?) on conflict(key) do update set value=excluded.value', (key, str(value)))


def month_spend() -> float:
    """This month's Claude API spend (costs.py logs it in api_calls); sample entries are left out."""
    r = q("select coalesce(sum(usd),0) s from api_calls where sample=0 and substr(ts,1,7)=?", (now()[:7],), one=True)
    return float(r['s'])


def budget_left() -> float:
    return float(setting('monthly_limit_usd', '0')) - month_spend()


def log_cost(user_id, project_id, job_id, module, source, usd=0.0, input_tokens=0, output_tokens=0, note=None):
    x('insert into cost_log(ts,user_id,project_id,job_id,module,source,input_tokens,output_tokens,usd,note) values(?,?,?,?,?,?,?,?,?,?)',
      (now(), user_id, project_id, job_id, module, source, input_tokens, output_tokens, usd, note))
