"""make_offline.py - self-contained client copy of a model page (no internet, no claude.ai).
usage: python make_offline.py model.html "Asset Name" out_dir [node_modules_dir]
needs: npm i three@0.128.0 @fontsource/barlow@5.0.13 @fontsource/barlow-condensed@5.0.13   (in node_modules_dir's parent)
writes out_dir/AAA-Area-Model-<Asset>.html and .zip; prints sizes. Raises if any external URL is left."""
import sys, os, re, base64, zipfile
src, asset, out = sys.argv[1], sys.argv[2], sys.argv[3]
nm = sys.argv[4] if len(sys.argv) > 4 else 'node_modules'
page = open(src, encoding='utf-8').read()
# fonts named in the Google Fonts link -> @fontsource woff2 (latin), base64
faces = []
for fam in re.findall(r'family=([^&"]+)', page):
    name, _, spec = fam.partition(':')
    name = name.replace('+', ' ')
    weights = re.findall(r'\d{3}', spec) or ['400']
    pkg = name.lower().replace(' ', '-')
    for w in weights:
        f = f'{nm}/@fontsource/{pkg}/files/{pkg}-latin-{w}-normal.woff2'
        if not os.path.exists(f): raise SystemExit(f'missing font file {f} - npm i @fontsource/{pkg}@5.0.13')
        b64 = base64.b64encode(open(f, 'rb').read()).decode()
        faces.append(f"@font-face{{font-family:'{name}';font-style:normal;font-weight:{w};font-display:swap;src:url(data:font/woff2;base64,{b64}) format('woff2')}}")
# strip Google Fonts, preconnects, Claude-only meta, old title
page = re.sub(r'<link[^>]+(fonts\.googleapis|fonts\.gstatic)[^>]*>\s*', '', page)
page = re.sub(r'<meta name="claude:[^"]*"[^>]*>\s*', '', page)
page = re.sub(r'<title>.*?</title>\s*', '', page, flags=re.S)
# inline scripts (same r128 build)
LIBS = {r'<script src="https://cdnjs\.cloudflare\.com/ajax/libs/three\.js/r128/three\.min\.js"></script>': f'{nm}/three/build/three.min.js',
        r'<script src="https://cdn\.jsdelivr\.net/npm/three@0\.128\.0/examples/js/controls/OrbitControls\.js"></script>': f'{nm}/three/examples/js/controls/OrbitControls.js'}
for pat, f in LIBS.items():
    js = open(f, encoding='utf-8').read().replace('</script', '<\\/script')
    page, n = re.subn(pat, lambda m: '<script>\n' + js + '\n</script>', page)
    if n != 1: raise SystemExit(f'script tag not found once: {pat}')
title = f'Area Model – {asset} – AAA Consultancy'
reset = '[hidden]{display:none!important}img{max-width:100%}body{margin:0}'
doc = f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{title}</title>
<style>{''.join(faces)}{reset}</style>
</head>
<body>
{page}
</body>
</html>
'''
left = sorted(set(re.findall(r'https?://[^\s"\')<>]+', doc)))
bad = [u for u in left if any(h in u for h in ('cdnjs', 'googleapis', 'gstatic', 'jsdelivr', 'unpkg', 'claude.ai'))]
# three.js carries a few URLs inside comments/strings (e.g. docs links); they are not requests. Only fail on CDN/font/claude hosts.
if bad: raise SystemExit(f'external URLs remain: {bad}')
os.makedirs(out, exist_ok=True)
slug = re.sub(r'[^A-Za-z0-9]+', '-', asset).strip('-')
html_path = f'{out}/AAA-Area-Model-{slug}.html'; zip_path = f'{out}/AAA-Area-Model-{slug}.zip'
open(html_path, 'w', encoding='utf-8').write(doc)
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z: z.write(html_path, os.path.basename(html_path))
print(html_path, os.path.getsize(html_path), zip_path, os.path.getsize(zip_path), 'fonts', len(faces), 'other urls in text:', len(left))
