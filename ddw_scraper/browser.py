"""
Browser helpers for the DDW scraper.

Provides utilities for navigating the conference planner page and safely
reading text from Playwright locators.
"""

import re

from ddw_scraper.logger import log


def get_schedule_frame(page):
    """Click the Schedule tab and return the schedule iframe frame locator.

    Waits for the iframe to attach and for its body to become visible before
    returning so that callers can immediately start querying the frame.
    """
    log("Clicking Schedule tab...")
    schedule_tab = page.get_by_role("link", name=re.compile(r"Schedule", re.I))
    schedule_tab.click()

    log("Waiting for schedule iframe...")
    page.wait_for_selector("iframe#link2_iframe", state="attached", timeout=20000)
    page.wait_for_timeout(2000)

    frame = page.frame_locator("iframe#link2_iframe").first
    try:
        frame.locator("body").wait_for(state="visible", timeout=15000)
    except Exception:
        pass

    log("Waiting for agenda to render...")
    page.wait_for_timeout(4000)  # extra time for JS to render agenda
    return frame


def collect_text(el, timeout: int = 1500) -> str:
    """Return the inner text of a locator's first match, or '' on any error."""
    try:
        return el.inner_text(timeout=timeout).strip()
    except Exception:
        return ""
