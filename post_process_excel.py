"""
post_process_excel.py — Post-process a scraped DDW Excel file.

Applies three transformations to an existing results file and writes a new
``_formatted.xlsx`` output:

  1. Abstract # extraction
       Regex ``^([A-Za-z]{0,3}\\d{2,4})[:\\s]`` on Sub-Session Title.
       Matched prefix (e.g. Mo1910, Sp858, 991) → Abstract # column.
       Sub-Session Title keeps the full original text.

  2. Time format → HH:MM:SS (24-hour)
       All AM/PM time strings in session-level and sub-session time cells
       are converted: "8:10AM" → "08:10:00", "5:15pm" → "17:15:00".

  3. Session Title + Time merged into one column
       New header: "Session Title/ Session Time"
       Cell value:  "<session name>\\n<HH:MM:SS> - <HH:MM:SS>"
       The separate Session Start Time and Session End Time columns are
       removed; Location shifts left from col M to col K.

Column layout of the output file (18 columns):
  A  Priority               — blank
  B  Coverage               — blank
  C  Indication             — from input
  D  Company                — from input
  E  Product Name           — from input
  F  Mechanism/ Targets     — from input
  G  Track                  — from input
  H  Science Type           — from input
  I  Session Title/ Session Time  — merged per session (NEW: name + time)
  J  Session Type           — merged per session
  K  Location               — merged per session  (was col M)
  L  Abstract #             — extracted from sub-session title
  M  Sub-Session Start      — time-formatted  (was col O)
  N  Sub-Session End        — time-formatted  (was col P)
  O  Sub-Session Title      — from input       (was col Q)
  P  Presenter              — blank
  Q  Sponsor/ Collaboration — from input       (was col S)
  R  Affiliations           — blank

Usage::

    python post_process_excel.py                      # latest results file
    python post_process_excel.py path/to/input.xlsx   # explicit input
"""

import re
import sys
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Input column positions (1-based) — matches the existing 20-col output
# ---------------------------------------------------------------------------
_IN_INDICATION   = 3   # C
_IN_COMPANY      = 4   # D
_IN_PRODUCT      = 5   # E
_IN_MOA          = 6   # F
_IN_TRACK        = 7   # G
_IN_SCIENCE_TYPE = 8   # H
_IN_SESSION_NAME = 9   # I  (merged)
_IN_FORMAT       = 10  # J  (merged)
_IN_START        = 11  # K  (merged)
_IN_END          = 12  # L  (merged)
_IN_LOCATION     = 13  # M  (merged)
# N (14) = Abstract # — blank in input
_IN_SUB_START    = 15  # O
_IN_SUB_END      = 16  # P
_IN_SUB_NAME     = 17  # Q
# R (18) = Presenter — blank
_IN_SOCIETY      = 19  # S
# T (20) = Affiliations — blank

# ---------------------------------------------------------------------------
# Output column definitions — (header, width, data-key | None)
# ---------------------------------------------------------------------------
_OUT_COLUMNS = [
    ("Priority",                18.7, None),               # A
    ("Coverage",                16.9, None),               # B
    ("Indication",              18.0, "indication"),       # C
    ("Company",                 25.0, "company"),          # D
    ("Product Name",            25.0, "product_name"),     # E
    ("Mechanism/ Targets",      22.0, "moa"),              # F
    ("Track",                   30.0, "track"),            # G
    ("Science Type",            18.0, "science_type"),     # H
    ("Session Title/ Session Time", 68.7, "session_title_time"),  # I ← merged
    ("Session Type",            22.0, "format"),           # J ← merged
    ("Location",                20.0, "region"),           # K ← merged
    ("Abstract #",              18.0, "abstract_num"),     # L
    ("Sub-Session Start",       20.0, "sub_start"),        # M
    ("Sub-Session End",         20.0, "sub_end"),          # N
    ("Sub-Session Title",       80.9, "sub_name"),         # O
    ("Presenter",               28.9, None),               # P
    ("Sponsor/ Collaboration",  22.0, "society"),          # Q
    ("Affiliations",            23.0, None),               # R
]

_MERGE_COLS = [9, 10, 11]   # I, J, K — merged per session

# ---------------------------------------------------------------------------
# Styling (same palette as the original exporter)
# ---------------------------------------------------------------------------
_HEADER_FILL  = PatternFill("solid", fgColor="002060")
_HEADER_FONT  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

_DATA_FILL  = PatternFill("solid", fgColor="E9F7FD")
_DATA_FONT  = Font(name="Calibri", size=14)
_DATA_ALIGN = Alignment(vertical="top", wrap_text=True)

_ROW1_HEIGHT   = 45.75
_HEADER_HEIGHT = 80.1
_DATA_HEIGHT   = 80.1

# Day sheets — matches config.DAY_SHEETS
_DAY_SHEETS = [
    ("May 1", "May 1"),
    ("May 2", "May 2"),
    ("May 3", "May 3"),
    ("May 4", "May 4"),
    ("May 5", "May 5"),
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _to_24h(time_str: str) -> str:
    """Convert an AM/PM time string to HH:MM:SS (24-hour).

    Handles formats like "8:10AM", "8:10 AM", "5:15pm", "12:00PM".
    Returns the original string unchanged if it cannot be parsed.
    """
    if not time_str:
        return ""
    m = re.match(r"(\d{1,2})(?::(\d{2}))?\s*([AP]M)", time_str.strip(), re.I)
    if not m:
        return time_str
    hour   = int(m.group(1))
    minute = int(m.group(2) or "0")
    ampm   = m.group(3).upper()
    if ampm == "AM":
        if hour == 12:
            hour = 0
    else:
        if hour != 12:
            hour += 12
    return f"{hour:02d}:{minute:02d}:00"


_ABSTRACT_PAT = re.compile(r"^([A-Za-z]{0,3}\d{2,4})\s*[:\-]?\s+")


def _extract_abstract(sub_name: str) -> str:
    """Return the abstract number prefix from a sub-session title, or ''."""
    m = _ABSTRACT_PAT.match(sub_name.strip())
    return m.group(1) if m else ""


def _style(cell, fill, font, alignment) -> None:
    cell.fill      = fill
    cell.font      = font
    cell.alignment = alignment


# ---------------------------------------------------------------------------
# Reader — reconstructs session/sub-session structure from an existing Excel
# ---------------------------------------------------------------------------

def _read_sessions(wb) -> list[dict]:
    """Walk each day sheet and rebuild the session row dicts."""
    rows = []

    for sheet_name, day_label in _DAY_SHEETS:
        if sheet_name not in wb.sheetnames:
            print(f"  Sheet '{sheet_name}' not found, skipping.")
            continue

        ws = wb[sheet_name]

        cur_session  = ""
        cur_format   = ""
        cur_start    = ""
        cur_end      = ""
        cur_location = ""
        cur_row: dict | None = None

        for row_idx in range(3, ws.max_row + 1):
            def v(col, ri=row_idx):
                val = ws.cell(ri, col).value
                return str(val).strip() if val is not None else ""

            session_name = v(_IN_SESSION_NAME)
            sub_name     = v(_IN_SUB_NAME)

            if not session_name and not sub_name:
                continue  # blank spacer row

            if session_name:
                cur_session  = session_name
                cur_format   = v(_IN_FORMAT)
                cur_start    = v(_IN_START)
                cur_end      = v(_IN_END)
                cur_location = v(_IN_LOCATION)
                cur_row = {
                    "day_label":    day_label,
                    "session_name": cur_session,
                    "start_time":   cur_start,
                    "end_time":     cur_end,
                    "region":       cur_location,
                    "track":        v(_IN_TRACK),
                    "format":       cur_format,
                    "society":      v(_IN_SOCIETY),
                    "science_type": v(_IN_SCIENCE_TYPE),
                    "indication":   v(_IN_INDICATION),
                    "company":      v(_IN_COMPANY),
                    "product_name": v(_IN_PRODUCT),
                    "moa":          v(_IN_MOA),
                    "sub_sessions": [],
                }
                rows.append(cur_row)

            if cur_row is not None:
                # Non-merged fields repeat on every row — take latest non-empty value
                for field, col in (
                    ("track",        _IN_TRACK),
                    ("science_type", _IN_SCIENCE_TYPE),
                    ("society",      _IN_SOCIETY),
                    ("indication",   _IN_INDICATION),
                    ("company",      _IN_COMPANY),
                    ("product_name", _IN_PRODUCT),
                    ("moa",          _IN_MOA),
                ):
                    val = v(col)
                    if val:
                        cur_row[field] = val

            if sub_name and cur_row is not None:
                cur_row["sub_sessions"].append({
                    "name":  sub_name,
                    "start": v(_IN_SUB_START),
                    "end":   v(_IN_SUB_END),
                })

    return rows


# ---------------------------------------------------------------------------
# Transformer — apply all three changes in-place on the row dicts
# ---------------------------------------------------------------------------

def _transform(rows: list[dict]) -> None:
    """Apply time formatting, abstract extraction, and title+time merge."""
    for r in rows:
        # 1. Format session-level times
        start_fmt = _to_24h(r["start_time"])
        end_fmt   = _to_24h(r["end_time"])
        r["start_time"] = start_fmt
        r["end_time"]   = end_fmt

        # 2. Build "Session Title/ Session Time" combined value
        if start_fmt and end_fmt:
            time_range = f"{start_fmt} - {end_fmt}"
        elif start_fmt:
            time_range = start_fmt
        else:
            time_range = ""

        r["session_title_time"] = (
            f"{r['session_name']}\n{time_range}" if time_range else r["session_name"]
        )

        # 3. Format sub-session times
        for sub in r["sub_sessions"]:
            sub["start"] = _to_24h(sub["start"])
            sub["end"]   = _to_24h(sub["end"])


# ---------------------------------------------------------------------------
# Writer — writes the new 18-column Excel
# ---------------------------------------------------------------------------

def _write_excel(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, day_label in _DAY_SHEETS:
        ws = wb.create_sheet(title=sheet_name[:31])
        day_rows = [r for r in rows if r.get("day_label") == day_label]

        # Column widths
        for col_idx, (_h, width, _k) in enumerate(_OUT_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Row 1: blank spacer
        ws.row_dimensions[1].height = _ROW1_HEIGHT
        for col_idx in range(1, len(_OUT_COLUMNS) + 1):
            _style(ws.cell(1, col_idx), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)

        # Row 2: headers
        ws.row_dimensions[2].height = _HEADER_HEIGHT
        for col_idx, (header, _w, _k) in enumerate(_OUT_COLUMNS, start=1):
            _style(ws.cell(2, col_idx, value=header), _HEADER_FILL, _HEADER_FONT, _HEADER_ALIGN)

        ws.freeze_panes = "A3"

        if not day_rows:
            ws.row_dimensions[3].height = _DATA_HEIGHT
            _style(ws.cell(3, 1, value="No sessions found for this day."), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)
            continue

        current_row = 3
        for r in day_rows:
            subs = r.get("sub_sessions") or []

            # Sessions with no sub-sessions → single synthetic row
            if not subs:
                subs = [{
                    "name":         r.get("session_name", ""),
                    "start":        r.get("start_time", ""),
                    "end":          r.get("end_time", ""),
                    "abstract_num": "",
                }]

            n_subs = len(subs)
            session_start_row = current_row

            for sub_idx, sub in enumerate(subs):
                ws.row_dimensions[current_row].height = _DATA_HEIGHT
                is_first = sub_idx == 0

                values = {
                    # Keyword columns — per sub-session row
                    "indication":          r.get("indication", ""),
                    "company":             r.get("company", ""),
                    "product_name":        r.get("product_name", ""),
                    "moa":                 r.get("moa", ""),
                    # Session meta
                    "track":               r.get("track", ""),
                    "science_type":        r.get("science_type", ""),
                    # Merged session columns — first sub-session row only
                    "session_title_time":  r.get("session_title_time", "") if is_first else "",
                    "format":              r.get("format", "")              if is_first else "",
                    "region":              r.get("region", "")              if is_first else "",
                    # Sub-session columns
                    "abstract_num":        sub.get("abstract_num", ""),
                    "sub_start":           sub.get("start", ""),
                    "sub_end":             sub.get("end", ""),
                    "sub_name":            sub.get("name", ""),
                    "society":             r.get("society", ""),
                }

                for col_idx, (_h, _w, key) in enumerate(_OUT_COLUMNS, start=1):
                    val = values.get(key, "") if key else ""
                    _style(ws.cell(current_row, col_idx, value=val), _DATA_FILL, _DATA_FONT, _DATA_ALIGN)

                current_row += 1

            # Merge session-level columns (I, J, K) vertically
            if n_subs > 1:
                for col_num in _MERGE_COLS:
                    col_letter = get_column_letter(col_num)
                    ws.merge_cells(
                        f"{col_letter}{session_start_row}:{col_letter}{current_row - 1}"
                    )
                    _style(
                        ws.cell(session_start_row, col_num),
                        _DATA_FILL, _DATA_FONT, _DATA_ALIGN,
                    )

    wb.save(path)
    print(f"Saved: {path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _latest_result() -> Path:
    results_dir = Path(__file__).resolve().parent / "results"
    files = sorted(
        p for p in results_dir.glob("ddw-conference-schedule-2026-*.xlsx")
        if "_checkpoint" not in p.name
        and "_retagged"  not in p.name
        and "_formatted" not in p.name
        and "_diff"      not in p.name
        and not p.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"No result files found in {results_dir}")
    return files[-1]


def main() -> int:
    if len(sys.argv) >= 2:
        input_path = Path(sys.argv[1])
    else:
        input_path = _latest_result()

    output_path = input_path.with_name(input_path.stem + "_formatted.xlsx")

    print(f"Input : {input_path.name}")
    print(f"Output: {output_path.name}")

    print("Reading Excel...")
    wb = openpyxl.load_workbook(input_path, data_only=True)
    rows = _read_sessions(wb)
    wb.close()
    print(f"  {len(rows)} sessions read.")

    print("Applying transformations...")
    _transform(rows)

    print(f"  Times converted to HH:MM:SS.")
    print(f"  Session title + time merged into single column.")

    print("Writing formatted Excel...")
    _write_excel(rows, output_path)
    print("Done.")
    return 0


if __name__ == "__main__":
    exit(main())
