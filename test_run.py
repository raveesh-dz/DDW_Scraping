"""
Quick connectivity test — entry point.

Verifies that Playwright is installed, the browser launches, and the
DDW planner page loads with the Schedule iframe appearing.  Run this
before a full scrape to confirm the environment is working correctly.

Usage::

    python test_run.py
"""

import sys
import traceback
from pathlib import Path

# Ensure the project root is on sys.path when this script is run from any
# working directory (e.g. python /full/path/test_run.py).
sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    from playwright.sync_api import sync_playwright
except Exception as e:
    print(f"Import error: {e}", flush=True)
    sys.exit(1)

from ddw_scraper import config


def log(msg: str) -> None:
    print(msg, flush=True)


def main() -> None:
    log("Starting...")
    with sync_playwright() as p:
        log("Launching browser (headless)...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        page = context.new_page()
        try:
            log("Loading page (timeout 20s)...")
            page.goto(config.PLANNER_URL, wait_until="domcontentloaded", timeout=20000)
            log("Page loaded. Waiting 3s...")
            page.wait_for_timeout(3000)
            log("Clicking Schedule tab...")
            page.get_by_role("link", name="Schedule").click()
            log("Waiting for iframe (15s)...")
            page.wait_for_selector("iframe#link2_iframe", state="attached", timeout=15000)
            log("Iframe found. Done.")
        except Exception as e:
            log(f"Error: {type(e).__name__}: {e}")
            traceback.print_exc()
        finally:
            browser.close()
    log("Exit.")


if __name__ == "__main__":
    main()
