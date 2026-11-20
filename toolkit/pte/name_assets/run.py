#!/usr/bin/env python3
"""
Name the assets across a whole AAA tender evaluation workbook, in one run.

Two template families are handled; which one you have is read from the file, not
asked for:

  PACKAGED  the workbook ships "Asset #01 {no package}" and "Asset #02
            {with package}". Each asset gets its own set of packages, or none,
            and the run also updates Scheduale of Assets and Project Summary.
  FLAT      the older workbook with a single asset template and no packages.

Phases
  1  asset tabs + Asset Summary (+ Scheduale of Assets, Project Summary)
  2  the elemental NRM sheet - a column group per asset
  3  checks: will Excel open it, and are the numbers right

The input file is never modified; everything is written to --out.

Usage
  python3 run.py --file IN.xlsx --out OUT.xlsx \
      --assets "Building 01: Structure, Architecture, Plumbing, Electrical" \
               "Building 02: Structure, MEP" \
               "External Works"
  An asset with no colon has no packages. On a flat workbook package lists are
  rejected rather than quietly ignored.

  [--rows 15]        spare Asset Summary rows beyond the named assets (flat only)
  [--return-col K]   K = AAA Pre-Tender Estimate, J = AAA Estimated
  [--aggregate]      sumif (default when packaged) adds every row matching an NRM
                     code so all packages count; xlookup takes the first match
  [--skip-nrm]       leave the elemental sheet alone
  [--skip-checks]    only when you are running the checks yourself
"""
import argparse, os, shutil, subprocess, sys, tempfile, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))


def run(script, args, title):
    print('\n' + '=' * 72)
    print(title)
    print('=' * 72)
    p = subprocess.run([sys.executable, os.path.join(HERE, script)] + args,
                       capture_output=True, text=True)
    sys.stdout.write(p.stdout)
    if p.returncode != 0:
        sys.stdout.write(p.stderr)
    return p.returncode


def family(path):
    with zipfile.ZipFile(path) as z:
        wb = z.read('xl/workbook.xml').decode()
    return 'packaged' if '{with package}' in wb else 'flat'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--assets', nargs='+', required=True)
    ap.add_argument('--rows', type=int)
    ap.add_argument('--return-col', default='K')
    ap.add_argument('--aggregate', choices=['xlookup', 'sumif'])
    ap.add_argument('--skip-nrm', action='store_true')
    ap.add_argument('--skip-checks', action='store_true')
    a = ap.parse_args()

    if not os.path.exists(a.file):
        sys.exit('No such file: %s' % a.file)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)) or '.', exist_ok=True)

    fam = family(a.file)
    names = [s.split(':', 1)[0].strip() for s in a.assets]
    with_pkgs = [s for s in a.assets if ':' in s]
    print('template family : %s' % fam)
    print('assets          : %s' % ', '.join(names))
    if fam == 'flat' and with_pkgs:
        sys.exit('This workbook has no packaged template, but package lists were given '
                 'for: %s\nDrop the package lists, or start from the workbook that '
                 'ships "Asset #02 {with package}".'
                 % ', '.join(s.split(':', 1)[0].strip() for s in with_pkgs))

    aggregate = a.aggregate or ('sumif' if fam == 'packaged' else 'xlookup')
    tmp = tempfile.mkdtemp(prefix='_assets')
    stage = os.path.join(tmp, 'stage.xlsx')

    if fam == 'packaged':
        rc = run('build_r00.py', ['--file', a.file, '--out', stage, '--assets'] + a.assets,
                 'PHASE 1  asset tabs, Scheduale, Asset & Project Summary')
    else:
        args1 = ['--file', a.file, '--out', stage, '--names'] + names
        if a.rows:
            args1 += ['--rows', str(a.rows)]
        rc = run('build_assets.py', args1, 'PHASE 1  asset tabs and Asset Summary')
    if rc:
        sys.exit('Phase 1 failed; nothing written to %s' % a.out)

    if a.skip_nrm:
        shutil.copy(stage, a.out)
        print('\nElemental sheet left untouched (--skip-nrm).')
    else:
        args2 = ['--file', stage, '--out', a.out, '--return-col', a.return_col,
                 '--aggregate', aggregate, '--names'] + names
        if fam == 'packaged':
            # the header cell IS the asset name here: the sheet looks the BUA up in
            # the Scheduale of Assets by it, so a suffix would break the match
            args2 += ['--label-suffix', '']
        if run('nrm_assets.py', args2, 'PHASE 2  elemental NRM summary'):
            shutil.copy(stage, a.out)
            print('\nPhase 2 failed. %s holds the phase 1 result; the elemental sheet '
                  'is unchanged from the input.' % a.out)
            sys.exit(1)

    if a.skip_checks:
        print('\nChecks skipped (--skip-checks). Run validate_xlsx.py before sending '
              'this file to anyone.')
        return

    bad = 0
    # Excel opens far fewer files than LibreOffice does; this one is not optional.
    bad += run('validate_xlsx.py', [a.out], 'PHASE 3a  will Excel open it')
    base = subprocess.run([sys.executable, os.path.join(HERE, 'validate_xlsx.py'), a.file],
                          capture_output=True, text=True)
    print('\n(the same checks against the input file: %s)'
          % ('PASS' if base.returncode == 0 else
             'FAIL - so treat any matching failure above as pre-existing'))

    if fam == 'packaged':
        bad += run('verify_r00.py', ['--file', a.out], 'PHASE 3b  numbers, end to end')
    else:
        bad += run('verify_summary.py', ['--file', a.out, '--baseline', a.file],
                   'PHASE 3b  Asset Summary numbers')
        if not a.skip_nrm:
            bad += run('verify_nrm.py', ['--file', a.out, '--names'] + names,
                       'PHASE 3c  elemental sheet numbers')

    print('\n' + '=' * 72)
    print('RESULT: %s   ->  %s' % ('ALL CHECKS PASSED' if not bad else 'CHECKS FAILED', a.out))
    print('=' * 72)
    sys.exit(0 if not bad else 1)


if __name__ == '__main__':
    main()
