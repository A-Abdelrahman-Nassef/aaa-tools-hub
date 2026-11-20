"""shot_ui.py - full-screen model screenshot (side panel + labelled 3D view) for the "Explore the model yourself" page.
usage: python shot_ui.py AAA-Area-Model-<Asset>.html out.jpg [LEVEL_ID]
Uses the offline file with networking switched off; picks the given level (default: first above-ground level) in 3D, Components mode."""
import asyncio, os, sys, json, pathlib
from playwright.async_api import async_playwright
import browser
f, out = os.path.abspath(sys.argv[1]), sys.argv[2]
lvl = sys.argv[3] if len(sys.argv) > 3 else None
async def main():
    async with async_playwright() as p:
        b = await browser.launch(p)   # finds the browser even when Playwright cannot (browser.py)
        ctx = await b.new_context(offline=True, viewport={'width': 1500, 'height': 920}, device_scale_factor=1.5, color_scheme='light')
        pg = await ctx.new_page(); await pg.goto(pathlib.Path(f).as_uri()); await pg.wait_for_timeout(2500)
        await pg.evaluate("""(l) => { const m = window.__model || window.__villa;
          const id = l || (m.levelGroups.find(g => g.z0 >= -0.01) || m.levelGroups[0]).id;
          m.setLevel(id); m.setPlan(false); m.setMode('comp'); }""", lvl)
        await pg.wait_for_timeout(1800)
        png = out.rsplit('.', 1)[0] + '.png'; await pg.screenshot(path=png); await b.close()
    from PIL import Image
    Image.open(png).convert('RGB').save(out, quality=88); os.remove(png); print('wrote', out)
asyncio.run(main())
