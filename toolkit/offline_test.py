"""offline_test.py - open the offline model in headless Chromium with networking OFF.
usage: python offline_test.py AAA-Area-Model-<Asset>.html
Pass = errors [] , external requests [] , meshes > 0 , barlow True. Writes <file>-offline-test.png next to it."""
import asyncio, os, sys, pathlib
from playwright.async_api import async_playwright
import browser   # finds the browser even when Playwright cannot (browser.py)
f = os.path.abspath(sys.argv[1])
async def main():
    async with async_playwright() as p:
        b = await browser.launch(p)
        ctx = await b.new_context(offline=True, viewport={'width':1400,'height':860})
        pg = await ctx.new_page()
        errs, reqs = [], []
        pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.on('console', lambda m: errs.append('console: '+m.text) if m.type == 'error' else None)
        pg.on('request', lambda r: reqs.append(r.url) if not r.url.startswith(('file:', 'data:', 'blob:')) else None)
        await pg.goto(pathlib.Path(f).as_uri()); await pg.wait_for_timeout(3000)
        r = await pg.evaluate('''async () => { await document.fonts.ready;
          const faces=[...document.fonts].map(x=>x.family+' '+x.weight+' '+x.status);
          const m = window.__model || window.__villa; let meshes=0; m.levelGroups.forEach(L=>L.group.traverse(o=>{if(o.isMesh) meshes++;}));
          const c=document.getElementById('cv'); const g=c.getContext('webgl')||c.getContext('webgl2');
          return {faces, meshes, three: THREE.REVISION, canvas:[c.width,c.height],
            barlow: document.fonts.check('600 13px "Barlow Condensed"') && document.fonts.check('400 14px Barlow'),
            title: document.title, h1: getComputedStyle(document.querySelector('h1')).fontFamily}; }''')
        await pg.screenshot(path=f.rsplit('.',1)[0] + '-offline-test.png')
        print('errors', errs); print('external requests', reqs); print(r)
        await b.close()
asyncio.run(main())
