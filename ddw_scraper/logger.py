"""
Logging utility for the DDW scraper.

Writes timestamped messages to stdout and, when ``config.LOG_TO_FILE`` is
True, appends them to ``config.LOG_PATH``.

Usage::

    from ddw_scraper.logger import log
    log("Starting scrape...")

Note: ``config`` is referenced as a live module object (not as frozen
``from … import`` names) so that any runtime changes to ``config.LOG_TO_FILE``
or ``config.LOG_PATH`` are always respected.
"""

import sys
from datetime import datetime

from ddw_scraper import config


def log(msg: str) -> None:
    """Print a timestamped message and optionally write it to the log file."""
    ts = datetime.now().isoformat(timespec="seconds")
    line = f"[{ts}] {msg}"

    # Windows consoles can be cp1252; never crash on a logging call.
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        safe = line.encode(enc, errors="backslashreplace").decode(enc, errors="ignore")
        print(safe, flush=True)

    if config.LOG_TO_FILE:
        try:
            config.LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with config.LOG_PATH.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            # Logging must never interrupt the scraping process.
            pass
