"""
Count sessions per day in the DDWLITE26 planner — entry point.

Counts how many sessions are shown for each day (Fri–Tue) on the website.
Paid sessions (titles containing '$$$') are included in these counts so that
you can see the true website totals, as opposed to the subset written to Excel
by the main scraper.

Usage::

    python count_sessions.py
"""

import sys
from pathlib import Path

# Ensure the project root is on sys.path when this script is run from any
# working directory (e.g. python /full/path/count_sessions.py).
sys.path.insert(0, str(Path(__file__).resolve().parent))

from playwright.sync_api import sync_playwright

from ddw_scraper import config
from ddw_scraper.browser import get_schedule_frame


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        page = context.new_page()
        page.goto(config.PLANNER_URL, wait_until="domcontentloaded", timeout=25000)
        page.wait_for_timeout(4000)

        frame = get_schedule_frame(page)

        # Day tabs with data-date attributes
        day_tabs = frame.locator(".agenda_multi_day_selector[data-date]")
        n_days = day_tabs.count()

        print("Session counts per day (as shown on website; includes paid $$$):")
        for idx in range(n_days):
            tab = day_tabs.nth(idx)
            data_date = tab.get_attribute("data-date")
            day_label = data_date or f"Day {idx + 1}"
            tab.click()
            page.wait_for_timeout(1200)

            list_container = frame.locator(
                ".agenda_list_container[data-table='agenda']"
            ).first
            container = frame.locator(
                f'.date_container.agenda_sublist[data-listdate="{data_date}"]'
            ).first

            # Scroll to trigger lazy-loading until the session count is stable.
            prev_count = -1
            stable_rounds = 0
            max_rounds = 20
            for _ in range(max_rounds):
                cells = container.locator(".agenda_cell.list_cell")
                count = cells.count()
                if count == prev_count:
                    stable_rounds += 1
                    if stable_rounds >= 3:
                        break
                else:
                    stable_rounds = 0
                    prev_count = count
                try:
                    list_container.evaluate("el => { el.scrollTop = el.scrollHeight; }")
                except Exception:
                    break
                page.wait_for_timeout(800)

            container = frame.locator(
                f'.date_container.agenda_sublist[data-listdate="{data_date}"]'
            ).first
            count = container.locator(".agenda_cell.list_cell").count()
            print(f"- {day_label}: {count} sessions")

        browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
