"""
Core scraping logic for the DDW schedule planner.

Public API
----------
scrape_schedule(frame) -> list[dict]
    Scrape all days and sessions from the schedule iframe frame.

scrape_session_detail(frame, cell) -> dict
    Click a single session cell, parse its detail view, and close it.
"""

import time

from ddw_scraper import config
from ddw_scraper import keywords as kw
from ddw_scraper.browser import collect_text
from ddw_scraper.exporter import write_excel
from ddw_scraper.logger import log
from ddw_scraper.parser import (
    parse_date_time,
    parse_label_value,
    parse_presentations_from_detail,
)


# ---------------------------------------------------------------------------
# Detail view
# ---------------------------------------------------------------------------

_DETAIL_MAX_RETRIES = 2     # try up to this many times if detail comes back blank


def _close_detail_view(frame) -> None:
    """Dismiss any open detail overlay so the next cell click lands correctly."""
    try:
        close_btn = frame.locator(".closeViewButton").first
        if close_btn.is_visible(timeout=500):
            close_btn.click(timeout=2000)
            time.sleep(0.5)
    except Exception:
        pass


def _try_scrape_detail(frame, cell) -> tuple[dict, str]:
    """One attempt at opening and reading a session's detail iframe.

    Returns ``(result_dict, body_text)``.
    """
    result = {
        "track": "",
        "format": "",
        "society": "",
        "science_type": "",
        "sub_sessions": [],
    }
    body_text = ""
    try:
        cell.scroll_into_view_if_needed(timeout=3000)
        time.sleep(0.3)
        cell.click()
        time.sleep(1.5)

        iframe_el = frame.locator("#planner_agenda_detail_iframe")
        iframe_el.wait_for(state="visible", timeout=15000)
        time.sleep(1.5)

        detail_frame = frame.frame_locator("#planner_agenda_detail_iframe").first
        body_text = collect_text(detail_frame.locator("body"), timeout=8000)

        result["track"] = parse_label_value(body_text, "Track")
        result["format"] = parse_label_value(body_text, "Format")
        result["society"] = parse_label_value(body_text, "Society")
        result["science_type"] = parse_label_value(body_text, "Science Type")
        result["sub_sessions"] = parse_presentations_from_detail(body_text)
    except Exception:
        pass
    finally:
        _close_detail_view(frame)

    return result, body_text


def scrape_session_detail(frame, cell, debug_name: str = "") -> dict:
    """Click *cell*, parse its detail iframe, close the detail, and return metadata.

    Returns a dict with keys: ``track``, ``format``, ``society``,
    ``science_type``, ``sub_sessions`` (list of ``{name, start, end}``).
    All values default to empty strings / empty list on any failure.

    Before clicking the cell the function dismisses any stale detail overlay
    left open by a previous scrape and scrolls the cell into view.  If the
    first attempt returns an empty result it retries once.
    """
    # Dismiss any stale detail overlay before we click.
    _close_detail_view(frame)

    result, body_text = _try_scrape_detail(frame, cell)

    # Retry once if the detail came back completely blank.
    if not result["track"] and not result["format"]:
        log(f"    Retrying detail for: {debug_name[:80]}")
        _close_detail_view(frame)
        time.sleep(1)
        result, body_text = _try_scrape_detail(frame, cell)

    # Debug: save body text when no presentations were parsed.
    if config.DEBUG_SAVE_ZERO_PRES_BODY and not result["sub_sessions"] and body_text:
        try:
            config.DEBUG_ZERO_PRES_DIR.mkdir(parents=True, exist_ok=True)
            safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in debug_name)
            out_path = config.DEBUG_ZERO_PRES_DIR / f"{safe_name[:80]}.txt"
            out_path.write_text(body_text, encoding="utf-8")
        except Exception:
            pass

    return result


# ---------------------------------------------------------------------------
# Per-cell helpers
# ---------------------------------------------------------------------------

def _build_row(
    day_label: str,
    session_name: str,
    start_time: str,
    end_time: str,
    date_str: str,
    region: str,
    detail: dict,
) -> dict:
    """Assemble a single session row dict from its components.

    Keyword tags are stored per sub-session (in each sub dict) rather than at
    the session level.  For sessions with no sub-sessions, ``detail["_session_kw"]``
    holds the tags for the single synthetic row the exporter will produce.
    """
    return {
        "day_label": day_label,
        "session_name": session_name,
        "start_time": start_time,
        "end_time": end_time,
        "date": date_str,
        "region": region,
        "track": detail["track"],
        "format": detail["format"],
        "society": detail["society"],
        "science_type": detail["science_type"],
        "notes": "",
        "sub_sessions": detail["sub_sessions"],
        "_session_kw": detail.get("_session_kw", {}),
    }


def _scrape_cell(
    frame,
    cell,
    j: int,
    cell_count: int,
    day_label: str,
    patterns: dict,
) -> dict | None:
    """Scrape one session cell and return a row dict, or None if skipped.

    A cell is skipped when its title contains '``$$$``' (paid session).
    """
    session_name = collect_text(cell.locator(".session_title.list_cell_title span"), timeout=800)
    if not session_name:
        session_name = collect_text(cell.locator(".session_title"), timeout=800)

    if "$$$" in session_name:
        log(f"  Skipping paid session {j + 1}/{cell_count}: {session_name[:120]}")
        return None

    if config.VERBOSE_SESSION_LOGS:
        log(f"  Session {j + 1}/{cell_count}: {session_name[:120]}")

    category = collect_text(cell.locator(".category_label"), timeout=500)
    date_time_str = collect_text(cell.locator(".date_time"), timeout=500)
    start_time, end_time, date_str, region = parse_date_time(date_time_str)

    if not region and category:
        region = category

    detail = scrape_session_detail(frame, cell, debug_name=session_name)

    if config.VERBOSE_SESSION_LOGS:
        log(
            f"    Detail: Track='{detail['track']}', Format='{detail['format']}', "
            f"Society='{detail['society']}', ScienceType='{detail['science_type']}', "
            f"Presentations={len(detail['sub_sessions'])}"
        )

    # Session-level fields shared by every sub-session row.
    session_base_fields = [
        session_name,
        detail["track"],
        detail["format"],
        detail["society"],
        detail["science_type"],
    ]

    # Tag each sub-session individually: session-level fields + that sub's
    # own name.  A keyword in a session-level field propagates to all rows;
    # a keyword that only appears in one sub-session name tags only that row.
    for sub in detail["sub_sessions"]:
        sub_kw = kw.match_keywords(session_base_fields + [sub["name"]], patterns)
        sub["indication"]   = sub_kw.get("Indication", "")
        sub["company"]      = sub_kw.get("Company", "")
        sub["product_name"] = sub_kw.get("Product Name", "")
        sub["moa"]          = sub_kw.get("MOA", "")

    # For sessions with no sub-sessions the exporter creates a synthetic row
    # using the session name.  Pre-compute tags for that case too.
    if not detail["sub_sessions"]:
        session_kw = kw.match_keywords(session_base_fields, patterns)
        detail["_session_kw"] = session_kw
    else:
        detail["_session_kw"] = {}

    return _build_row(day_label, session_name, start_time, end_time, date_str, region, detail)


def _save_checkpoint(rows: list[dict]) -> None:
    """Save an intermediate checkpoint Excel file if enabled in config."""
    if not config.CHECKPOINT_AFTER_EACH_DAY:
        return
    try:
        write_excel(rows, config.OUTPUT_EXCEL_CHECKPOINT)
        log(f"Checkpoint saved: {config.OUTPUT_EXCEL_CHECKPOINT}")
    except Exception as e:
        log(f"Checkpoint save failed: {type(e).__name__}: {e}")


# ---------------------------------------------------------------------------
# Day-tab scraping
# ---------------------------------------------------------------------------

_SCROLL_MAX_ROUNDS = 20     # safety cap so we never scroll forever
_SCROLL_PAUSE_SECS = 1.5   # wait after each scroll for lazy-loaded cells


def _scroll_to_load_all(frame, container, day_label: str) -> int:
    """Scroll the session container until no new cells appear.

    The ATIV planner lazy-loads session cells — only a portion is rendered
    initially and more appear as the user scrolls down.  This helper scrolls
    the container to its bottom in a loop until the cell count stabilises,
    ensuring the scraper sees every session.

    Returns the final cell count.
    """
    cells_selector = ".agenda_cell.list_cell"
    prev_count = 0
    try:
        prev_count = container.locator(cells_selector).count()
    except Exception:
        return 0

    for rnd in range(_SCROLL_MAX_ROUNDS):
        # Scroll the last cell into view to trigger lazy loading.
        try:
            last_cell = container.locator(cells_selector).nth(prev_count - 1)
            last_cell.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass
        time.sleep(_SCROLL_PAUSE_SECS)

        try:
            new_count = container.locator(cells_selector).count()
        except Exception:
            break

        if new_count == prev_count:
            # No new cells loaded — we have them all.
            break

        log(f"  Scroll round {rnd + 1}: {prev_count} -> {new_count} cells ({day_label})")
        prev_count = new_count

    return prev_count


def _scrape_with_day_tabs(frame, day_tabs, n_days: int, patterns: dict) -> list[dict]:
    """Click each day tab and scrape its session container."""
    rows: list[dict] = []

    for day_idx in range(n_days):
        tab = day_tabs.nth(day_idx)
        try:
            data_date = tab.get_attribute("data-date")
        except Exception:
            data_date = None

        if config.ONLY_DATES is not None and data_date not in config.ONLY_DATES:
            continue

        day_label = config.DATE_TO_LABEL.get(
            data_date, data_date or config.DAY_SHEETS[day_idx][1]
        )
        log(f"Starting day {day_idx + 1}/{n_days}: {day_label} ({data_date})")

        tab.click()
        time.sleep(1.2)  # allow the session list to update

        container = frame.locator(
            f'.date_container.agenda_sublist[data-listdate="{data_date}"]'
        ).first

        # Scroll until every lazy-loaded cell is in the DOM.
        cell_count = _scroll_to_load_all(frame, container, day_label)
        log(f"  {day_label}: {cell_count} total cells after scroll.")

        cells = container.locator(".agenda_cell.list_cell")
        for j in range(cell_count):
            row = _scrape_cell(frame, cells.nth(j), j, cell_count, day_label, patterns)
            if row is not None:
                rows.append(row)

        log(f"Finished {day_label}: {cell_count} sessions scraped.")
        _save_checkpoint(rows)

    return rows


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def scrape_schedule(frame) -> list[dict]:
    """Scrape all days and sessions from the schedule iframe *frame*.

    Returns a flat list of session row dicts ready to be passed to
    :func:`~ddw_scraper.exporter.write_excel`.
    """
    if config.DEBUG_SAVE_HTML:
        try:
            html = frame.locator("html").first.evaluate("el => el.outerHTML")
            config.DEBUG_HTML_PATH.write_text(html, encoding="utf-8")
            print(f"Debug: saved iframe HTML to {config.DEBUG_HTML_PATH}")
        except Exception as e:
            print(f"Debug save HTML failed: {e}")

    # Load and compile keyword patterns once for the entire run.
    try:
        patterns = kw.build_patterns(kw.load_keywords())
        log(f"Loaded keywords: {sum(len(v) for v in patterns.values())} total across {len(patterns)} columns.")
    except Exception as e:
        log(f"Warning: could not load keywords file ({e}). Keyword columns will be empty.")
        patterns = {col: [] for col in config.KEYWORD_COLUMNS}

    day_tabs = frame.locator(".agenda_multi_day_selector[data-date]")
    try:
        n_days = day_tabs.count()
    except Exception:
        n_days = 0

    if n_days == 0:
        raise RuntimeError("No day-selector tabs found in the schedule iframe.")

    return _scrape_with_day_tabs(frame, day_tabs, n_days, patterns)
