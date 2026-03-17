"""
Excel export for the DDW scraper — ECCO 2026 CRDI Planner format.

Replicates the visual style of the ECCO 2026 CRDI Planner:

- Row 1  : blank spacer row (light blue)
- Row 2  : dark navy header row, white bold text, centered
- Rows 3+: light blue data rows, wrap text, top-aligned
- Session-level columns (I–M) are merged vertically across all sub-session
  rows that belong to the same parent session.

Column layout (20 columns):
  A  Priority               — left blank
  B  Coverage               — left blank
  C  Indication             — from keyword matching
  D  Company                — from keyword matching (dedicated)
  E  Product Name           — from keyword matching (dedicated)
  F  Mechanism/ Targets     — MOA from keyword matching
  G  Track                  — scraped track field
  H  Science Type           — scraped science type field
  I  Session Name           — session title              (merged per session)
  J  Session Type           — Format field               (merged per session)
  K  Session Start Time     — session-level start time   (merged per session)
  L  Session End Time       — session-level end time     (merged per session)
  M  Location               — Region field               (merged per session)
  N  Abstract #             — left blank
  O  Sub-Session Start      — sub-session start time
  P  Sub-Session End        — sub-session end time
  Q  Sub-Session Title      — sub-session / presentation name
  R  Presenter              — left blank
  S  Sponsor/ Collaboration — Society field
  T  Affiliations           — left blank
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ddw_scraper import config

# ---------------------------------------------------------------------------
# Styling constants (matched from ECCO 2026 CRDI Planner.xlsx)
# ---------------------------------------------------------------------------
_HEADER_FILL  = PatternFill("solid", fgColor="002060")   # dark navy blue
_HEADER_FONT  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

_DATA_FILL  = PatternFill("solid", fgColor="E9F7FD")     # light blue
_DATA_FONT  = Font(name="Calibri", size=14)
_DATA_ALIGN = Alignment(vertical="top", wrap_text=True)

_ROW1_HEIGHT   = 45.75   # blank spacer row
_HEADER_HEIGHT = 80.1
_DATA_HEIGHT   = 80.1

# ---------------------------------------------------------------------------
# Column definitions — (header label, width, data-key or None)
# data-key is looked up in the per-row values dict; None = always empty
# ---------------------------------------------------------------------------
_COLUMNS = [
    ("Priority",               18.7,  None),              # A — blank
    ("Coverage",               16.9,  None),              # B — blank
    ("Indication",             18.0,  "indication"),      # C
    ("Company",                25.0,  "company"),         # D — dedicated
    ("Product Name",           25.0,  "product_name"),    # E — dedicated
    ("Mechanism/ Targets",     22.0,  "moa"),             # F
    ("Track",                  30.0,  "track"),           # G — dedicated
    ("Science Type",           18.0,  "science_type"),    # H — dedicated
    ("Session Name",           68.7,  "session_name"),    # I ← merged
    ("Session Type",           22.0,  "format"),          # J ← merged
    ("Session Start Time",     20.0,  "start_time"),      # K ← merged
    ("Session End Time",       20.0,  "end_time"),        # L ← merged
    ("Location",               20.0,  "region"),          # M ← merged
    ("Abstract #",             18.0,  None),              # N — blank
    ("Sub-Session Start",      20.0,  "sub_start"),       # O
    ("Sub-Session End",        20.0,  "sub_end"),         # P
    ("Sub-Session Title",      80.9,  "sub_name"),        # Q
    ("Presenter",              28.9,  None),              # R — blank
    ("Sponsor/ Collaboration", 22.0,  "society"),         # S
    ("Affiliations",           23.0,  None),              # T — blank
]

# 1-based column numbers for session-level columns that get merged vertically
_MERGE_COL_NUMBERS = [9, 10, 11, 12, 13]   # I, J, K, L, M


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _style(cell, fill, font, alignment) -> None:
    cell.fill      = fill
    cell.font      = font
    cell.alignment = alignment


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def write_excel(rows: list[dict], path) -> None:
    """Write *rows* to an ECCO-style Excel workbook at *path*.

    One sheet per conference day (defined in ``config.DAY_SHEETS``).
    Session-level columns (I–M) are merged vertically across all sub-session
    rows belonging to the same parent session.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, day_label in config.DAY_SHEETS:
        ws = wb.create_sheet(title=sheet_name[:31])
        day_rows = [r for r in rows if r.get("day_label") == day_label]

        # --- Column widths ---
        for col_idx, (_h, width, _k) in enumerate(_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # --- Row 1: blank light-blue spacer ---
        ws.row_dimensions[1].height = _ROW1_HEIGHT
        for col_idx in range(1, len(_COLUMNS) + 1):
            _style(ws.cell(1, col_idx), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)

        # --- Row 2: header ---
        ws.row_dimensions[2].height = _HEADER_HEIGHT
        for col_idx, (header, _w, _k) in enumerate(_COLUMNS, start=1):
            _style(ws.cell(2, col_idx, value=header), _HEADER_FILL, _HEADER_FONT, _HEADER_ALIGN)

        # Freeze panes so the header stays visible while scrolling
        ws.freeze_panes = "A3"

        # --- No data case ---
        if not day_rows:
            ws.row_dimensions[3].height = _DATA_HEIGHT
            _style(ws.cell(3, 1, value="No sessions found for this day."), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)
            continue

        # --- Data rows ---
        current_row = 3
        for r in day_rows:
            subs = r.get("sub_sessions") or []

            # Sessions with no sub-sessions produce a single row
            if not subs:
                subs = [{
                    "name":  r.get("session_name", ""),
                    "start": r.get("start_time", ""),
                    "end":   r.get("end_time", ""),
                }]

            n_subs = len(subs)
            session_start_row = current_row

            for sub_idx, sub in enumerate(subs):
                ws.row_dimensions[current_row].height = _DATA_HEIGHT
                is_first = sub_idx == 0

                # Per-column values for this Excel row
                values = {
                    "indication":    r.get("indication", ""),
                    "company":       r.get("company", ""),
                    "product_name":  r.get("product_name", ""),
                    "moa":           r.get("moa", ""),
                    "track":         r.get("track", ""),
                    "science_type":  r.get("science_type", ""),
                    # Session-level fields — only on the first sub-session row
                    "session_name":  r.get("session_name", "") if is_first else "",
                    "format":        r.get("format", "")       if is_first else "",
                    "start_time":    r.get("start_time", "")   if is_first else "",
                    "end_time":      r.get("end_time", "")     if is_first else "",
                    "region":        r.get("region", "")       if is_first else "",
                    # Sub-session fields
                    "sub_start":     sub.get("start", ""),
                    "sub_end":       sub.get("end", ""),
                    "sub_name":      sub.get("name", ""),
                    "society":       r.get("society", ""),
                }

                for col_idx, (_h, _w, key) in enumerate(_COLUMNS, start=1):
                    val = values.get(key, "") if key else ""
                    _style(ws.cell(current_row, col_idx, value=val), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)

                current_row += 1

            # Merge session-level columns (I–M) vertically across sub-session rows
            if n_subs > 1:
                for col_num in _MERGE_COL_NUMBERS:
                    col_letter = get_column_letter(col_num)
                    ws.merge_cells(
                        f"{col_letter}{session_start_row}:{col_letter}{current_row - 1}"
                    )
                    # Re-apply style to the merged cell (openpyxl resets it on merge)
                    _style(
                        ws.cell(session_start_row, col_num),
                        _DATA_FILL, _DATA_FONT, _DATA_ALIGN,
                    )

    wb.save(path)
    print(f"Saved: {path}")
