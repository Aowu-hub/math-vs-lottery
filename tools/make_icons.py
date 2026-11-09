# python3 tools/make_icons.py — renders web/icon.svg to the PNG app icons
# (needs Python Playwright; run once whenever the icon design changes).
from pathlib import Path
from playwright.sync_api import sync_playwright
root = Path(__file__).resolve().parent.parent
svg = (root / 'web' / 'icon.svg').read_text(encoding='utf-8')
(root / 'web' / 'icons').mkdir(parents=True, exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch()
    for name, size, pad in [('icon-180', 180, 0), ('icon-192', 192, 0), ('icon-512', 512, 0), ('icon-maskable-512', 512, 0.12)]:
        pg = b.new_page(viewport={'width': size, 'height': size})
        art = svg
        if pad:  # maskable: background stays full-bleed, the picture shrinks into the safe zone
            k = 1 - pad * 2
            bg = '<rect width="512" height="512" fill="url(#bg)"/>'
            art = art.replace(bg, bg + f'<g transform="translate({256 * (1 - k):.1f} {256 * (1 - k):.1f}) scale({k:.3f})">').replace('</svg>', '</g></svg>')
        art = art.replace('width="512" height="512"', f'width="{size}" height="{size}"', 1)
        pg.set_content(f'<html><body style="margin:0">{art}</body></html>')
        pg.screenshot(path=str(root / 'web' / 'icons' / f'{name}.png'))
        pg.close()
    b.close()
print('icons made')
