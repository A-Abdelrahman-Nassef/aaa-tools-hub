"""render_stills.py - report stills of one asset model (for the deck).
usage: python render_stills.py model.html asset.json out_dir [name,name,...]
model.html must load three.js locally (the offline copy does; or swap the two CDN script srcs for node_modules paths).
Writes comp, skin, bua, ipms, compare (whole asset) and plan_<LEVEL> (components) + planc_<LEVEL> (compare) PNGs.
The optional 4th argument renders only the named stills (e.g. "compare")."""
import asyncio, os, sys, json, pathlib
from playwright.async_api import async_playwright
import browser
html, data, out = sys.argv[1:4]
lv = [l['id'] for l in json.load(open(data, encoding='utf-8'))['levels']]
SHOTS = [(m, 'all', m, False) for m in ('comp', 'skin', 'bua', 'ipms', 'compare')] + \
        [(f'plan_{k}', k, 'comp', True) for k in lv] + [(f'planc_{k}', k, 'compare', True) for k in lv]
if len(sys.argv) > 4:
    want = set(sys.argv[4].split(','))
    SHOTS = [s for s in SHOTS if s[0] in want]
async def main():
    os.makedirs(out, exist_ok=True)
    async with async_playwright() as p:
        b = await browser.launch(p)   # finds the browser even when Playwright cannot (browser.py)
        pg = await b.new_page(viewport={'width': 1500, 'height': 1000}, device_scale_factor=1.6)
        await pg.goto(pathlib.Path(os.path.abspath(html)).as_uri()); await pg.wait_for_timeout(2500)
        await pg.add_style_tag(content='.levels,.tools,.hint,.legend{display:none!important} :root{--stage:#F2F4F7!important}')
        await pg.evaluate("document.documentElement.setAttribute('data-theme','light')"); await pg.wait_for_timeout(300)
        for name, l, mode, plan in SHOTS:
            await pg.evaluate(f"__model.setLevel('{l}')"); await pg.wait_for_timeout(200)
            await pg.evaluate(f"__model.setPlan({str(plan).lower()})"); await pg.evaluate(f"__model.setMode('{mode}')")
            await pg.wait_for_timeout(1200)
            await pg.locator('#stage').screenshot(path=f'{out}/{name}.png')
        await b.close()
    # crop to content (light background) so images sit tight in the slide panels
    from PIL import Image, ImageChops
    for f in os.listdir(out):
        im = Image.open(f'{out}/{f}').convert('RGB'); bg = Image.new('RGB', im.size, im.getpixel((5, 5)))
        bb = ImageChops.difference(im, bg).point(lambda v: 255 if v > 12 else 0).getbbox()
        if bb:
            pad = 30; bb = (max(0, bb[0]-pad), max(0, bb[1]-pad), min(im.width, bb[2]+pad), min(im.height, bb[3]+pad))
            im.crop(bb).save(f'{out}/{f}')
    print('stills in', out, sorted(os.listdir(out)))
asyncio.run(main())
