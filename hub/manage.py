"""manage.py - hub admin from the command line.

  python -m hub.manage adduser <email> "<Name>" [--admin]     asks for the password (not echoed)
  python -m hub.manage passwd <email>
  python -m hub.manage users
  python -m hub.manage initials <email> AAN                    initials for that user's output file names
  python -m hub.manage samples add|delete                     sample Claude API cost entries for the Costs page
"""
import getpass, sys
from . import db


def main(argv):
    db.init()
    if not argv or argv[0] not in ('adduser', 'passwd', 'users', 'samples', 'initials'):
        print(__doc__); return 1
    cmd = argv[0]
    if cmd == 'initials' and len(argv) == 3:
        v = ''.join(ch for ch in argv[2] if ch.isalpha()).upper()[:6]
        n = db.q('select count(*) n from users where email=?', (argv[1].strip().lower(),), one=True)['n']
        db.x('update users set initials=? where email=?', (v or None, argv[1].strip().lower()))
        print(f'initials {v} set' if n else 'no such user'); return 0 if n else 1
    if cmd == 'samples':
        from . import costs
        if argv[1:] == ['add']:
            print('added', costs.add_samples(), 'sample entries')
        elif argv[1:] == ['delete']:
            print('deleted', costs.delete_samples(), 'sample entries')
        else:
            print(__doc__); return 1
        return 0
    if cmd == 'users':
        for u in db.q('select id,email,name,is_admin,status,created from users order by id'):
            print(u['id'], u['email'], u['name'], 'admin' if u['is_admin'] else 'member', u['status'], u['created'])
        return 0
    email = argv[1].strip().lower()
    pw = getpass.getpass('Password: ')
    if len(pw) < 10 or pw != getpass.getpass('Again: '):
        print('Passwords must match and be at least 10 characters.'); return 1
    if cmd == 'adduser':
        name = argv[2]
        db.x('insert into users(email,name,pw_hash,is_admin,created) values(?,?,?,?,?)',
             (email, name, db.hash_pw(pw), int('--admin' in argv), db.now()))
        print('added', email)
    else:
        db.x('update users set pw_hash=? where email=?', (db.hash_pw(pw), email)); print('password changed', email)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
