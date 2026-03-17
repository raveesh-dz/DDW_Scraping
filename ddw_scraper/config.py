"""
Configuration constants for the DDW scraper.

All paths, URLs, flags, and column definitions live here so that other
modules never hard-code them.  Import this module as a whole (``from
ddw_scraper import config``) and access values as ``config.HEADLESS`` so
that any runtime patches to these flags are picked up immediately.
"""

from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Directory anchor (project root, one level above this package directory)
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------------------
# Output directory and run timestamp
# ---------------------------------------------------------------------------
RESULTS_DIR = BASE_DIR / "results"

# Captured once at import time so the filename reflects when the run started.
_RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d-%H%M%S")

# ---------------------------------------------------------------------------
# Network / browser settings
# ---------------------------------------------------------------------------
PLANNER_URL = "https://eppro01.ativ.me/web/planner.php?id=DDWLITE26"
SCHEDULE_IFRAME_TAB = "Schedule"

# Set to False to watch the browser during a run (recommended for first run /
# selector debugging).
HEADLESS = True

# ---------------------------------------------------------------------------
# Output paths
# ---------------------------------------------------------------------------
OUTPUT_EXCEL = RESULTS_DIR / f"ddw-conference-schedule-2026-{_RUN_TIMESTAMP}.xlsx"
OUTPUT_EXCEL_CHECKPOINT = RESULTS_DIR / f"ddw-conference-schedule-2026-{_RUN_TIMESTAMP}_checkpoint.xlsx"
LOG_PATH = BASE_DIR / "scrape_ddw_schedule.log"
DEBUG_HTML_PATH = BASE_DIR / "schedule_iframe_debug.html"
KEYWORDS_FILE = BASE_DIR / "DDW Keywords.xlsx"

# ---------------------------------------------------------------------------
# Runtime flags
# ---------------------------------------------------------------------------
# Set to True to save schedule iframe HTML for selector debugging.
DEBUG_SAVE_HTML = False

# Set to True to save the raw detail-view body text for every session that
# returns 0 presentations.  Files are written to DEBUG_ZERO_PRES_DIR so you
# can inspect the exact text the parser receives and diagnose missed sessions.
# Disable again before a full production run.
DEBUG_SAVE_ZERO_PRES_BODY = False
DEBUG_ZERO_PRES_DIR = BASE_DIR / "debug_body_text"

LOG_TO_FILE = True

# Print each session title as it is scraped.
VERBOSE_SESSION_LOGS = True

# Save a checkpoint Excel after each day completes.
CHECKPOINT_AFTER_EACH_DAY = True

# Restrict scraping to a subset of dates, e.g. {"2026-05-01"}.  None = all.
ONLY_DATES = None

# ---------------------------------------------------------------------------
# Schedule metadata
# ---------------------------------------------------------------------------
# Ordered list of (sheet_name, day_label) pairs used when writing Excel.
DAY_SHEETS = [
    ("May 1", "May 1"),
    ("May 2", "May 2"),
    ("May 3", "May 3"),
    ("May 4", "May 4"),
    ("May 5", "May 5"),
]

# Maps data-listdate DOM attribute values to date labels.
DATE_TO_LABEL = {
    "2026-05-01": "May 1",
    "2026-05-02": "May 2",
    "2026-05-03": "May 3",
    "2026-05-04": "May 4",
    "2026-05-05": "May 5",
}

# ---------------------------------------------------------------------------
# Excel column definitions
# See ddw_scraper/exporter.py for the full ECCO-format column layout.
# ---------------------------------------------------------------------------
# Keyword-matched columns (populated from DDW Keywords.xlsx).
KEYWORD_COLUMNS = [
    "Indication",
    "Company",
    "Product Name",
    "MOA",
]
