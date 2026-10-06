"""Flatten COPD+Bronchiectasis-filtered sessions into ATS Abstracts format.

Cols: Company, Laterals, Date, Time (EST), Session Type, Poster (if applicable),
Indication, Abstract Title, Session Code, Location, Notes
"""
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
SRC = ROOT / "ats_output" / "all_sessions_copd_bronchiectasis.json"
OUT = ROOT / "ats_output" / "ATS_COPD_Bronchiectasis_2026.xlsx"

TOPIC_NAMES = {"COPD", "Bronchiectasis"}


def indication_for(parent: dict) -> str:
    names = [c.get("name") for c in parent.get("categories") or []]
    matched = [n for n in ("COPD", "Bronchiectasis") if n in names]
    return " / ".join(matched)

ABSTRACT_PARENT_TYPES = {
    "Mini Symposium",
    "Poster Discussion Session",
    "Thematic Poster Session",
    "RAPiD Poster Discussion Session",
    "RAPiD Abstract Session",
    "Scientific Symposium",
    "Late Breaking Mini Symposium",
}
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
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th").join([str(n), ""])


def fmt_date(iso_dt: str) -> str:
    if not iso_dt:
        return ""
    dt = datetime.fromisoformat(iso_dt.split("T")[0])
    return dt.strftime(f"%A, %B {ordinal(dt.day)}")


def fmt_time(iso_start: str, iso_end: str) -> str:
    def _t(iso: str) -> str:
        if not iso or "T" not in iso:
            return ""
        return datetime.fromisoformat(iso).strftime("%I:%M %p").lstrip("0")
    return f"{_t(iso_start)} {_t(iso_end)}".strip()


def session_code(parent_name: str) -> str:
    m = SESSION_CODE_RE.match((parent_name or "").strip())
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
    sponsors = []
    for sr in parent.get("sponsorRoles") or []:
        sp = sr.get("sponsor") or {}
        if sp.get("name"):
            sponsors.append(sp["name"])
    if sponsors:
        return ", ".join(dict.fromkeys(sponsors))
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


def session_type_of(parent: dict) -> str:
    cats = parent.get("categories") or []
    types = [c.get("name") for c in cats if c.get("domain") == "SESSION_TYPE"]
    if not types:
        types = [c.get("name") for c in cats]
    abst = [t for t in types if t in ABSTRACT_PARENT_TYPES]
    if abst:
        return abst[0]
    cmp = [t for t in types if t in COMPANY_LED_TYPES]
    if cmp:
        return cmp[0]
    return types[0] if types else ""


def main():
    with open(SRC, encoding="utf-8") as f:
        all_data = json.load(f)

    rows = []
    for date, sessions in all_data.items():
        for parent in sessions:
            stype = session_type_of(parent)
            code = session_code(parent.get("name", ""))
            location = parent_location(parent.get("geozonesOfEdition"))
            parent_date_str = fmt_date(parent.get("startTime", ""))
            company = company_name(parent)
            indication = indication_for(parent)

            abstract_subs = [
                s for s in (parent.get("subsessions") or [])
                if s.get("activityType") == "Abstract"
            ]

            if abstract_subs:
                for sub in abstract_subs:
                    poster, title = split_poster(sub.get("name", ""))
                    rows.append({
                        "Company": company,
                        "Laterals": "",
                        "Date": fmt_date(sub.get("startTime") or parent.get("startTime", "")) or parent_date_str,
                        "Time (EST)": fmt_time(sub.get("startTime", ""), sub.get("endTime", "")),
                        "Start Time": sub.get("startTime") or "",
                        "End Time": sub.get("endTime") or "",
                        "Session Type": stype,
                        "Poster (if applicable)": poster,
                        "Indication": indication,
                        "Abstract Title": title,
                        "Session Code": code,
                        "Location": location,
                        "Notes": "",
                        "Parent Session": parent.get("name"),
                        "Parent UID": parent.get("uid"),
                        "Subsession UID": sub.get("uid"),
                        "URL": f"https://ats2026.d365.events/education/sessions/{sub.get('uid')}",
                    })
            else:
                rows.append({
                    "Company": company,
                    "Laterals": "",
                    "Date": parent_date_str,
                    "Time (EST)": fmt_time(parent.get("startTime", ""), parent.get("endTime", "")),
                    "Start Time": parent.get("startTime") or "",
                    "End Time": parent.get("endTime") or "",
                    "Session Type": stype,
                    "Poster (if applicable)": "",
                    "Indication": indication,
                    "Abstract Title": parent.get("name", ""),
                    "Session Code": code,
                    "Location": location,
                    "Notes": "",
                    "Parent Session": parent.get("name"),
                    "Parent UID": parent.get("uid"),
                    "Subsession UID": "",
                    "URL": f"https://ats2026.d365.events/education/sessions/{parent.get('uid')}",
                })

    df = pd.DataFrame(rows)
    print(f"Rows: {len(df)}")
    print("Session types:", df["Session Type"].value_counts().to_dict())

    ref_cols = [
        "Company", "Laterals", "Date", "Time (EST)", "Start Time", "End Time",
        "Session Type", "Poster (if applicable)", "Indication", "Abstract Title",
        "Session Code", "Location", "Notes",
    ]
    extra_cols = ["Parent Session", "Parent UID", "Subsession UID", "URL"]
    df = df[ref_cols + extra_cols]

    with pd.ExcelWriter(OUT, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="COPD_Bronchiectasis", index=False)
    df.to_csv(OUT.with_suffix(".csv"), index=False, encoding="utf-8-sig")
    print(f"Wrote: {OUT}")


if __name__ == "__main__":
    main()
