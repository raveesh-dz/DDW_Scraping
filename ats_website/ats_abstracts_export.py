"""Flatten all_sessions.json into ATS Abstracts format matching reference Excel.

Reference cols: Company, Laterals, Date, Time (EST), Session Type,
Poster (if applicable), Indication, Abstract Title, Session Code, Location, Notes

Company / Laterals / Indication / Notes are user-curated — left blank.
"""
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
SRC = ROOT / "ats_output" / "all_sessions.json"
OUT = ROOT / "ats_output" / "ATS_Abstracts_2026_extracted.xlsx"

# Session-type whitelist that contains abstracts/posters.
ABSTRACT_PARENT_TYPES = {
    "Mini Symposium",
    "Poster Discussion Session",
    "Thematic Poster Session",
    "RAPiD Poster Discussion Session",
    "RAPiD Abstract Session",
    "Scientific Symposium",
    "Late Breaking Mini Symposium",
}

# Session types where the session itself is the row (no abstract children).
# Mostly company-sponsored content.
COMPANY_LED_TYPES = {
    "Industry Theater",
    "Innovation Hub",
    "Industry Supported Evening Symposia",
    "Mid-day Symposium",
    "Product Theater",
}

POSTER_RE = re.compile(r"\(\s*Poster\s*Board\s*#\s*([A-Za-z0-9\-]+)\s*\)", re.I)
SESSION_CODE_RE = re.compile(r"^([A-Z]{1,3}\d{1,4}[A-Z]?)\b")


def ordinal(n: int) -> str:
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th").join([str(n), ""])  # noqa


def fmt_date(iso_dt: str) -> str:
    if not iso_dt:
        return ""
    dt = datetime.fromisoformat(iso_dt.split("T")[0])
    suffix = ordinal(dt.day)
    return dt.strftime(f"%A, %B {suffix}")


def fmt_time(iso_start: str, iso_end: str) -> str:
    def _t(iso: str) -> str:
        if not iso or "T" not in iso:
            return ""
        t = datetime.fromisoformat(iso).strftime("%I:%M %p").lstrip("0")
        return t
    s = _t(iso_start)
    e = _t(iso_end)
    return f"{s} {e}".strip()


def session_code(parent_name: str) -> str:
    if not parent_name:
        return ""
    m = SESSION_CODE_RE.match(parent_name.strip())
    return m.group(1) if m else ""


def parent_location(geos) -> str:
    if not geos:
        return ""
    g = geos[0]
    nm = g.get("name") or ""
    hall = (g.get("exhibitHall") or {}).get("name") or ""
    if hall and hall not in nm:
        return f"{nm} ({hall})"
    return nm


def company_name(parent: dict) -> str:
    ex = parent.get("exhibitor") or {}
    if ex.get("name"):
        return ex["name"]
    for sr in parent.get("sponsorRoles") or []:
        sp = sr.get("sponsor") or {}
        if sp.get("name"):
            return sp["name"]
    return ""


def split_poster(name: str) -> tuple[str, str]:
    if not name:
        return "", ""
    m = POSTER_RE.search(name)
    if not m:
        return "", name
    poster = f"Poster Board #{m.group(1)}"
    title = POSTER_RE.sub("", name).strip(" :-")
    return poster, title


def main():
    with open(SRC, encoding="utf-8") as f:
        all_data = json.load(f)

    rows = []

    for date, sessions in all_data.items():
        for parent in sessions:
            ptypes = {c.get("name") for c in parent.get("categories") or []}
            session_type = next(iter(ptypes & ABSTRACT_PARENT_TYPES), None)
            company_led = next(iter(ptypes & COMPANY_LED_TYPES), None)
            if not session_type:
                session_type = company_led or next(
                    (c.get("name") for c in parent.get("categories") or []
                     if c.get("domain") == "SESSION_TYPE"),
                    "",
                )

            code = session_code(parent.get("name", ""))
            location = parent_location(parent.get("geozonesOfEdition"))
            parent_date_str = fmt_date(parent.get("startTime", ""))
            company = company_name(parent)

            # Path 1: company-led parents → emit parent itself as a row.
            if company_led and not (parent.get("subsessions") or []):
                rows.append({
                    "Company": company,
                    "Laterals": "",
                    "Date": parent_date_str,
                    "Time (EST)": fmt_time(parent.get("startTime", ""), parent.get("endTime", "")),
                    "Session Type": session_type,
                    "Poster (if applicable)": "",
                    "Indication": "",
                    "Abstract Title": parent.get("name", ""),
                    "Session Code": code,
                    "Location": location,
                    "Notes": "",
                    "Parent Session": parent.get("name"),
                    "Parent UID": parent.get("uid"),
                    "Subsession UID": "",
                    "URL": f"https://ats2026.d365.events/education/sessions/{parent.get('uid')}",
                })
                continue

            # Path 2: abstract subsessions.
            for sub in parent.get("subsessions") or []:
                if sub.get("activityType") != "Abstract":
                    continue
                poster, title = split_poster(sub.get("name", ""))
                rows.append({
                    "Company": company,
                    "Laterals": "",
                    "Date": fmt_date(sub.get("startTime") or parent.get("startTime", "")) or parent_date_str,
                    "Time (EST)": fmt_time(sub.get("startTime", ""), sub.get("endTime", "")),
                    "Session Type": session_type,
                    "Poster (if applicable)": poster,
                    "Indication": "",
                    "Abstract Title": title,
                    "Session Code": code,
                    "Location": location,
                    "Notes": "",
                    "Parent Session": parent.get("name"),
                    "Parent UID": parent.get("uid"),
                    "Subsession UID": sub.get("uid"),
                    "URL": f"https://ats2026.d365.events/education/sessions/{sub.get('uid')}",
                })

    df = pd.DataFrame(rows)
    print(f"Abstracts extracted: {len(df)}")
    print("Session types:", df["Session Type"].value_counts().to_dict())

    # Reference column order, then extras at end.
    ref_cols = [
        "Company", "Laterals", "Date", "Time (EST)", "Session Type",
        "Poster (if applicable)", "Indication", "Abstract Title",
        "Session Code", "Location", "Notes",
    ]
    extra_cols = ["Parent Session", "Parent UID", "Subsession UID", "URL"]
    df = df[ref_cols + extra_cols]

    with pd.ExcelWriter(OUT, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Page_1_01", index=False)

    df.to_csv(OUT.with_suffix(".csv"), index=False, encoding="utf-8-sig")
    print(f"Wrote: {OUT}")


if __name__ == "__main__":
    main()
