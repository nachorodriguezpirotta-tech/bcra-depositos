#!/usr/bin/env python3
"""Renderiza el dashboard HTML a PNG (foto) con Playwright/Chromium headless."""
import os, sys
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))


def render(html_path=None, out_png=None, width=760):
    html_path = html_path or os.path.join(HERE, "dashboard.html")
    out_png = out_png or os.path.join(HERE, "dashboard.png")
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": width, "height": 1200}, device_scale_factor=2)
        pg.goto("file://" + html_path)
        pg.wait_for_timeout(400)
        pg.screenshot(path=out_png, full_page=True)
        b.close()
    print("PNG:", out_png)
    return out_png


def render_svg(svg, out_png, width=360, height=100):
    """Rasteriza un SVG a PNG con Chromium (para usar como <img> en mail/standalone)."""
    html = f'<!doctype html><body style="margin:0">{svg}</body>'
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": width, "height": height}, device_scale_factor=2)
        pg.set_content(html)
        pg.wait_for_timeout(150)
        pg.locator("svg").screenshot(path=out_png)
        b.close()
    return out_png


if __name__ == "__main__":
    render(*(sys.argv[1:3] if len(sys.argv) > 1 else []))
