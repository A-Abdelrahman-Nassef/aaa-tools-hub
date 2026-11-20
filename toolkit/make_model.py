"""make_model.py - fill model_template.html with one asset's data.
usage: python make_model.py model_template.html asset.json out.html"""
import sys, json, html
tpl, data, out = sys.argv[1:4]
d = json.load(open(data, encoding='utf-8')); m = d.get('meta', {})
page = open(tpl, encoding='utf-8').read()
for k, v in {'__PAGETITLE__': f"{d['asset']} Area Model", '__EYEBROW__': m.get('eyebrow', 'AAA Consultancy'),
             '__ASSET__': d['asset'], '__NOTE__': m.get('note', '')}.items():
    page = page.replace(k, html.escape(v))
page = page.replace('__DATA__', json.dumps(d, separators=(',', ':')))
open(out, 'w', encoding='utf-8').write(page); print('wrote', out, len(page))
