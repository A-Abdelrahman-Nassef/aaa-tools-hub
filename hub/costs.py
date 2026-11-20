"""costs.py - Claude API cost tracker.

Every Claude API call is logged in the `api_calls` table with its tokens and its cost in USD.
Prices come from config/model_prices.json (USD per 1M tokens), read on every call, so Ahmed edits that
file when prices change; a call's cost is fixed when it is logged.

Stage B: after each API call, the hub calls
    log_api_call(user_id, project_id, job_id, module, response.model, response.usage)
or passes the token counts by hand.
"""
import json

from . import db

PRICES_FILE = db.ROOT / 'config' / 'model_prices.json'
SAMPLE_NOTE = 'sample entry - delete on the Costs page'


def prices() -> dict:
    return json.loads(PRICES_FILE.read_text(encoding='utf-8'))['models']


def price_call(model, input_tokens=0, output_tokens=0, cache_read_tokens=0, cache_write_tokens=0,
               cache_write_1h_tokens=0):
    """USD for one call, or None when the model is not in the prices file."""
    p = prices().get(model)
    if not p:
        return None
    return (input_tokens * p['input'] + output_tokens * p['output'] + cache_read_tokens * p['cache_read']
            + cache_write_tokens * p['cache_write']
            + cache_write_1h_tokens * p.get('cache_write_1h', p['cache_write'])) / 1_000_000


def log_api_call(user_id, project_id, job_id, module, model, usage=None, input_tokens=0, output_tokens=0,
                 cache_read_tokens=0, cache_write_tokens=0, cache_write_1h_tokens=0, ts=None, sample=False,
                 note=None):
    """Log one Claude API call. `usage` may be the SDK's response.usage; otherwise pass the counts.
    A model with no price is still logged (usd 0) with a note, so the Costs page shows it."""
    if usage is not None:
        input_tokens = getattr(usage, 'input_tokens', 0) or 0
        output_tokens = getattr(usage, 'output_tokens', 0) or 0
        cache_read_tokens = getattr(usage, 'cache_read_input_tokens', 0) or 0
        cache_write_tokens = getattr(usage, 'cache_creation_input_tokens', 0) or 0
        cc = getattr(usage, 'cache_creation', None)          # split by cache lifetime, when the API gives it
        if cc is not None and getattr(cc, 'ephemeral_1h_input_tokens', None) is not None:
            cache_write_1h_tokens = cc.ephemeral_1h_input_tokens or 0
            cache_write_tokens = getattr(cc, 'ephemeral_5m_input_tokens', 0) or 0
    usd = price_call(model, input_tokens, output_tokens, cache_read_tokens, cache_write_tokens, cache_write_1h_tokens)
    if usd is None:
        note = (note + '; ' if note else '') + f'no price for {model} in config/model_prices.json'
        usd = 0.0
    return db.x('insert into api_calls(ts,user_id,project_id,job_id,module,model,input_tokens,output_tokens,'
                'cache_read_tokens,cache_write_tokens,usd,sample,note) values(?,?,?,?,?,?,?,?,?,?,?,?,?)',
                (ts or db.now(), user_id, project_id, job_id, module, model, input_tokens, output_tokens,
                 cache_read_tokens, cache_write_tokens + cache_write_1h_tokens, round(usd, 6), int(sample), note))


def add_samples():
    """A few made-up calls to try the Costs page; flagged sample=1, left out of the spending limit."""
    users = [r['id'] for r in db.q('select id from users order by id')]
    projects = [r['id'] for r in db.q('select id from projects order by id')]
    jobs = {r['project_id']: r['id'] for r in db.q('select id, project_id from jobs order by id')}
    if not users:
        return 0
    rows = [  # (date, user index, project index, module, model, in, out, cache read, cache write)
        ('2026-09-14 09:12:00', 0, 0, 'areas', 'claude-opus-5-5', 182_000, 21_500, 0, 45_000),
        ('2026-09-14 09:40:00', 0, 0, 'areas', 'claude-opus-5-5', 12_400, 18_200, 45_000, 0),
        ('2026-10-02 11:05:00', 1, 1, 'areas', 'claude-sonnet-5-5', 96_000, 14_000, 0, 30_000),
        ('2026-10-21 14:30:00', 1, 2, 'name-assets', 'claude-sonnet-5-5', 54_000, 6_800, 30_000, 0),
        ('2026-11-03 10:00:00', 0, 2, 'load-boq', 'claude-opus-5-5', 240_000, 32_000, 0, 60_000),
        ('2026-11-03 10:25:00', 0, 2, 'link-rates', 'claude-haiku-5-5', 410_000, 9_500, 60_000, 0),
        ('2026-11-12 16:45:00', 1, 1, 'areas', 'claude-fable-5-2', 75_000, 12_000, 0, 0),
    ]
    for ts, ui, pi, module, model, i, o, cr, cw in rows:
        uid = users[ui % len(users)]
        pid = projects[pi % len(projects)] if projects else None
        log_api_call(uid, pid, jobs.get(pid), module, model, input_tokens=i, output_tokens=o, cache_read_tokens=cr,
                     cache_write_tokens=cw, ts=ts, sample=True, note=SAMPLE_NOTE)
    return len(rows)


def delete_samples():
    n = db.q('select count(*) n from api_calls where sample=1', one=True)['n']
    db.x('delete from api_calls where sample=1')
    return n
