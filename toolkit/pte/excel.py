"""excel.py - recalculate or convert workbooks through Excel (Windows), via excel_com.ps1.

The skills recalculate with LibreOffice in the cloud sandbox; on the hub machine Excel does it.
"""
import json, os, subprocess
from pathlib import Path

PS1 = Path(__file__).resolve().parent / 'excel_com.ps1'


def _run(args, timeout):
    r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(PS1)] + args,
                       capture_output=True, text=True, timeout=timeout)
    line = (r.stdout.strip().splitlines() or [''])[-1]
    if r.returncode or not line.startswith('{'):
        raise RuntimeError(f'Excel step failed: {(r.stderr or r.stdout).strip()[:400]}')
    return json.loads(line)


def recalc(path, timeout=600):
    """Full recalculation in Excel, saved in place. Returns {'errors': n, 'samples': [...]}."""
    res = _run(['-Path', os.path.abspath(path)], timeout)
    s = res.get('samples') or []
    res['samples'] = [s] if isinstance(s, str) else list(s)
    return res


def run_ops(src, ops, out='', timeout=1800):
    """Run workbook operations in Excel (excel_ops.ps1). Saves to `out` when given; returns the 'read' values.

    ops: list of dicts, e.g. {'op': 'rename', 'sheet': 'Asset #01', 'to': 'VT'},
    {'op': 'column', 'sheet': s, 'col': 'C', 'from': 8, 'values': ['=SUBTOTAL(..)', '', ...]},
    {'op': 'read', 'sheet': s, 'cells': ['L105', 'K105']}.
    """
    import tempfile
    fd, path = tempfile.mkstemp(suffix='.json'); os.close(fd)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(ops, f, ensure_ascii=False)
        args = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(PS1.parent / 'excel_ops.ps1'),
                '-Src', os.path.abspath(src), '-Ops', path]
        if out:
            args += ['-Out', os.path.abspath(out)]
        r = subprocess.run(args, capture_output=True, text=True, timeout=timeout)
    finally:
        os.unlink(path)
    line = (r.stdout.strip().splitlines() or [''])[-1]
    res = json.loads(line) if line.startswith('{') else {'error': (r.stderr or r.stdout).strip()[:600]}
    if r.returncode or 'error' in res:
        raise RuntimeError(f"Excel operations failed: {res.get('error')}")
    return res.get('reads') or {}


def xls_to_xlsx(path, outdir, timeout=300):
    """Convert a legacy .xls to .xlsx in outdir; returns the new path."""
    os.makedirs(outdir, exist_ok=True)
    dst = os.path.abspath(os.path.join(outdir, Path(path).stem + '.xlsx'))
    _run(['-Path', os.path.abspath(path), '-SaveAs', dst], timeout)
    return dst
