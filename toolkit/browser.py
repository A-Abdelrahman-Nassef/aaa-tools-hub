"""browser.py - start the headless Chromium that the toolkit scripts use (offline test, stills, UI shots).

The hub started from the office-network window could not see %LOCALAPPDATA%\\ms-playwright (20 Nov 2026: the offline
test and the Annex D stills both failed with "Executable doesn't exist" while the browser was installed). So:
  1. a copy of the browser in the repo (aaa-tools-hub\\.browsers, git-ignored) is used first when present;
  2. if Playwright still cannot start its browser, every installed Chromium build Python can find is tried directly,
     and what Python saw is printed to the job log.
"""
import os, pathlib

REPO_BROWSERS = pathlib.Path(__file__).resolve().parent.parent / '.browsers'
if REPO_BROWSERS.is_dir() and not os.environ.get('PLAYWRIGHT_BROWSERS_PATH'):
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(REPO_BROWSERS)

ARGS = ['--use-gl=swiftshader', '--enable-webgl', '--ignore-gpu-blocklist', '--enable-unsafe-swiftshader']


def installed():
    """Chromium builds Playwright installed, newest first: headless shell, then full Chromium."""
    roots = [os.environ.get('PLAYWRIGHT_BROWSERS_PATH'), str(REPO_BROWSERS),
             os.path.join(os.environ.get('LOCALAPPDATA', ''), 'ms-playwright'),
             os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ms-playwright')]
    out = []
    for r in dict.fromkeys(x for x in roots if x):
        for pat in ('chromium_headless_shell-*/chrome-headless-shell-win64/chrome-headless-shell.exe',
                    'chromium-*/chrome-win64/chrome.exe'):
            out += sorted(pathlib.Path(r).glob(pat), reverse=True)
    return list(dict.fromkeys(out))


async def launch(p, args=ARGS):
    """p = the object from `async with async_playwright() as p`."""
    try:
        return await p.chromium.launch(args=args)
    except Exception as e:
        found = installed()
        print(f'Playwright could not start its browser: {str(e).splitlines()[0]}', flush=True)
        print(f'  LOCALAPPDATA={os.environ.get("LOCALAPPDATA")!r} PLAYWRIGHT_BROWSERS_PATH='
              f'{os.environ.get("PLAYWRIGHT_BROWSERS_PATH")!r}; Python sees: {[str(x) for x in found] or "no browser"}',
              flush=True)
        for exe in found:
            try:
                b = await p.chromium.launch(executable_path=str(exe), args=args)
                print(f'  started {exe} directly', flush=True)
                return b
            except Exception as e2:
                print(f'  {exe}: {str(e2).splitlines()[0]}', flush=True)
        raise
