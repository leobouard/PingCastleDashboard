"""Browser smoke test of a generated dashboard (used by the CI pipeline).

Usage: python tests/browser-smoke.py output/dashboard_<domain>.html
Needs: pip install playwright && python -m playwright install --with-deps chromium
"""
import asyncio, pathlib, sys
from playwright.async_api import async_playwright

failures = []
def check(name, ok, detail=''):
    print(('PASS ' if ok else 'FAIL ') + name + ('' if ok else f'  [{detail}]'))
    if not ok: failures.append(name)

async def main(path):
    url = pathlib.Path(path).resolve().as_uri()
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        for lang in ('en', 'fr'):
            page = await browser.new_page(viewport={'width': 1440, 'height': 900})
            errors = []
            page.on('pageerror', lambda e: errors.append(str(e)))
            page.on('dialog', lambda d: (errors.append('dialog: ' + d.message), asyncio.ensure_future(d.dismiss())))
            await page.goto(url); await page.wait_for_timeout(300)
            await page.click(f'[data-lang="{lang}"]'); await page.wait_for_timeout(200)
            check(f'{lang}: language applied', await page.evaluate('() => document.documentElement.lang') == lang)
            tabs = await page.eval_on_selector_all('[data-tab]', 'e => e.map(x => x.dataset.tab)')
            empty = []
            for theme in ('midnight', 'slate', 'report'):
                await page.click(f'[data-theme-btn="{theme}"]')
                for tab in tabs:
                    await page.click(f'[data-tab="{tab}"]'); await page.wait_for_timeout(80)
                    if not await page.evaluate("(t) => document.getElementById('pane-' + t).children.length", tab): empty.append(f'{theme}/{tab}')
            check(f'{lang}: {len(tabs)} tabs x 3 themes render', not empty, str(empty))
            nan = await page.evaluate("() => /NaN|Infinity/.test([...document.querySelectorAll('svg')].map(s => s.outerHTML).join(''))")
            check(f'{lang}: no NaN in charts', not nan)
            await page.click('[data-tab="home"]'); await page.wait_for_timeout(200)
            count = await page.locator('#pane-home [data-expand]').count(); opened = 0
            for i in range(count):
                btn = page.locator('#pane-home [data-expand]').nth(i)
                await btn.scroll_into_view_if_needed(); await btn.click(force=True); await page.wait_for_timeout(150)
                opened += await page.locator('.card.max').count()
                await page.keyboard.press('Escape'); await page.wait_for_timeout(100)
            check(f'{lang}: {count} expand buttons open and close', opened == count and await page.locator('.card.max').count() == 0, f'{opened}/{count}')
            check(f'{lang}: no JavaScript error', not errors, str(errors[:2]))
            await page.close()
        # phone width: no page-wide horizontal scroll
        for lang in ('en', 'fr'):
            page = await browser.new_page(viewport={'width': 375, 'height': 812})
            await page.goto(url); await page.click(f'[data-lang="{lang}"]'); await page.wait_for_timeout(200)
            wide = []
            for tab in await page.eval_on_selector_all('[data-tab]', 'e => e.map(x => x.dataset.tab)'):
                await page.click(f'[data-tab="{tab}"]'); await page.wait_for_timeout(80)
                if await page.evaluate('() => document.documentElement.scrollWidth > innerWidth + 1'): wide.append(tab)
            check(f'{lang} 375px: no horizontal scroll', not wide, str(wide))
            await page.close()
        await browser.close()
    print(f'\n{len(failures)} failure(s)')
    sys.exit(1 if failures else 0)

asyncio.run(main(sys.argv[1]))
