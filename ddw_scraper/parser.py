"""
Text parsing utilities for the DDW scraper.

All functions are pure (no side effects, no external dependencies) so they
are easy to unit-test in isolation.
"""

import re


def parse_date_time(date_time_str: str) -> tuple[str, str, str, str]:
    """Parse a combined date/time/region string into its components.

    Example input:  ``'12:00PM-5:15PM May 1 (Central)'``
    Example output: ``('12:00PM', '5:15PM', 'May 1', 'Central')``

    Returns a four-tuple ``(start_time, end_time, date_str, region)``.
    Any component that cannot be parsed is returned as an empty string.
    """
    start_time, end_time, date_str, region = "", "", "", ""

    if not date_time_str or not date_time_str.strip():
        return start_time, end_time, date_str, region

    # Time range: "12:00PM-5:15PM" or "12:00 PM - 5:15 PM"
    time_match = re.search(
        r"(\d{1,2}(?::\d{2})?\s*[AP]M)\s*[-–]\s*(\d{1,2}(?::\d{2})?\s*[AP]M)",
        date_time_str.replace(" ", ""),
        re.I,
    )
    if time_match:
        start_time = time_match.group(1).strip()
        end_time = time_match.group(2).strip()

    # Date: "May 1" or "May 01"
    date_match = re.search(r"(May\s+\d{1,2})", date_time_str, re.I)
    if date_match:
        date_str = date_match.group(1).strip()

    # Region: "(Central)", "(Eastern)", etc.
    region_match = re.search(r"\(([^)]+)\)", date_time_str)
    if region_match:
        region = region_match.group(1).strip()

    return start_time, end_time, date_str, region


def parse_label_value(text: str, label: str) -> str:
    """Extract the value that follows a labelled field in detail-page body text.

    The detail page uses a repeating pattern such as::

        Track
        Liver and Transplantation
        Format
        Lecture - Invited

    This function locates ``label`` in ``text``, trims everything up to the
    next known label, and returns the first non-empty, non-fragment line.
    """
    if not text:
        return ""
    idx = text.find(label)
    if idx < 0:
        return ""
    rest = text[idx + len(label):].lstrip(" :\n")
    for next_label in ("Format", "Society", "Science Type", "Audience", "Presentations"):
        if next_label != label and next_label in rest:
            rest = rest.split(next_label)[0]
    first_line = rest.split("\n")[0].strip()
    return first_line if first_line and not first_line.startswith("#") else ""


def parse_presentations_from_detail(text: str) -> list[dict]:
    """Extract presentation titles and times from the session detail body text.

    Scans for time-range patterns (e.g. ``"8:00am - 8:15am May 2 (Central)"``),
    then back-tracks through the preceding text to find the presentation title.

    Returns a list of dicts, each with keys ``name``, ``start``, and ``end``.
    """
    subs: list[dict] = []
    if not text:
        return subs

    time_pat = re.compile(
        r"(\d{1,2}:\d{2}\s*[ap]m)\s*[-–]\s*(\d{1,2}:\d{2}\s*[ap]m)\s+May\s+\d{1,2}\s*\([^)]+\)",
        re.I,
    )

    for m in time_pat.finditer(text):
        start_t = m.group(1).strip()
        end_t = m.group(2).strip()
        before = text[: m.start()].strip()

        title = ""
        if "#" in before:
            last_part = before.split("#")[-1].strip()
            title = last_part.split("\n")[0].strip() or ""

        if not title:
            skip_lines = {
                "Presentation", "ASGE", "Society", "Track",
                "Format", "Science Type", "Audience", "Presentations",
            }
            for line in reversed(before.replace("\r", "").split("\n")):
                line = line.strip()
                if line and not time_pat.search(line) and line not in skip_lines:
                    title = line.lstrip("#").strip()
                    break

        subs.append({"name": title or "Presentation", "start": start_t, "end": end_t})

    return subs
