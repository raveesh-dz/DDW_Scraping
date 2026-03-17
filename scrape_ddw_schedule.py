"""
DDWLITE26 Schedule Scraper — entry point.

Scrapes the Schedule tab from the DDW planner, including expandable
sub-sessions.  Outputs one Excel sheet per day (Friday May 1 through
Tuesday May 5, 2026).

Run once with ``config.DEBUG_SAVE_HTML = True`` and ``config.HEADLESS =
False`` to capture the schedule iframe HTML and refine selectors if needed.

Usage::

    python scrape_ddw_schedule.py
"""

import re
import sys
from pathlib import Path

# Ensure the project root is on sys.path when this script is run from any
# working directory (e.g. python /full/path/scrape_ddw_schedule.py).
sys.path.insert(0, str(Path(__file__).resolve().parent))

from playwright.sync_api import sync_playwright

from ddw_scraper import config
from ddw_scraper.browser import get_schedule_frame
from ddw_scraper.exporter import write_excel
from ddw_scraper.logger import log
from ddw_scraper.scraper import scrape_schedule


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=config.HEADLESS)
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        try:
            log(f"Starting scrape. Headless={config.HEADLESS}. LogFile={config.LOG_PATH}")
            log("Loading planner page...")
            page.goto(config.PLANNER_URL, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(4000)

            # Accept cookies / consent dialog if one appears.
            try:
                page.get_by_role("button", name=re.compile(r"OK|Accept", re.I)).click(
                    timeout=3000
                )
                page.wait_for_timeout(500)
            except Exception:
                pass

            frame = get_schedule_frame(page)
            log("Scraping schedule content...")
            rows: list[dict] = []
            try:
                rows = scrape_schedule(frame)
                log(f"Scraped {len(rows)} session rows.")
                write_excel(rows, config.OUTPUT_EXCEL)
                log(f"Saved final: {config.OUTPUT_EXCEL}")
                log("Done.")
            except Exception as e:
                log(f"ERROR during scrape: {type(e).__name__}: {e}")
                try:
                    if rows:
                        write_excel(rows, config.OUTPUT_EXCEL_CHECKPOINT)
                        log(f"Saved partial (checkpoint): {config.OUTPUT_EXCEL_CHECKPOINT}")
                finally:
                    raise

        finally:
            browser.close()

    return 0


if __name__ == "__main__":
    exit(main())
