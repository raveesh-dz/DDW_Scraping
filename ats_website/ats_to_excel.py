"""Flatten ats_output/all_sessions.json into Excel: Sessions + Subsessions sheets."""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
SRC = ROOT / "ats_output" / "all_sessions.json"
OUT = ROOT / "ats_output" / "ATS_2026_Program.xlsx"


def speaker_str(roles):
    if not roles:
        return ""
    parts = []
    for r in roles:
        sp = r.get("speaker") or {}
        nm = sp.get("name") or f"{sp.get('firstName','')} {sp.get('lastName','')}".strip()
        role = r.get("role") or ""
        parts.append(f"{nm} ({role})" if role else nm)
    return "; ".join(parts)


def cats_str(cats, domain=None):
    if not cats:
        return ""
    if domain:
        cats = [c for c in cats if c.get("domain") == domain]
    return "; ".join(c.get("name", "") for c in cats if c.get("name"))


def location_str(geos):
    if not geos:
        return ""
    parts = []
    for g in geos:
        nm = g.get("name", "")
        hall = (g.get("exhibitHall") or {}).get("name", "")
        parts.append(f"{nm} ({hall})" if hall else nm)
    return "; ".join(parts)


def main():
    with open(SRC, encoding="utf-8") as f:
        all_data = json.load(f)

    session_rows = []
    sub_rows = []

    for date, sessions in all_data.items():
        for s in sessions:
            session_type = cats_str(s.get("categories"), "SESSION_TYPE")
            session_rows.append({
                "Date": date,
                "Session UID": s.get("uid"),
                "Session Name": s.get("name"),
                "Session Type": session_type,
                "Start Time": s.get("startTime"),
                "End Time": s.get("endTime"),
                "Description": s.get("description"),
                "Featured": s.get("featured"),
                "Exhibitor": (s.get("exhibitor") or {}).get("name"),
                "Location": location_str(s.get("geozonesOfEdition")),
                "Speakers/Roles": speaker_str(s.get("speakerRoles")),
                "Categories (all)": cats_str(s.get("categoriesInContext")),
                "Subsession Count": len(s.get("subsessions") or []),
                "URL": f"https://ats2026.d365.events/education/sessions/{s.get('uid')}",
            })
            for sub in s.get("subsessions") or []:
                sub_rows.append({
                    "Date": date,
                    "Parent Session UID": s.get("uid"),
                    "Parent Session Name": s.get("name"),
                    "Parent Session Type": session_type,
                    "Subsession UID": sub.get("uid"),
                    "Subsession Name": sub.get("name"),
                    "Activity Type": sub.get("activityType"),
                    "Start Time": sub.get("startTime"),
                    "End Time": sub.get("endTime"),
                    "Description": sub.get("description"),
                    "Location": (sub.get("geozone") or {}).get("name"),
                    "Speakers/Roles": speaker_str(sub.get("speakerRoles")),
                    "Categories": cats_str(sub.get("categories")),
                    "URL": f"https://ats2026.d365.events/education/sessions/{sub.get('uid')}",
                })

    sess_df = pd.DataFrame(session_rows)
    sub_df = pd.DataFrame(sub_rows)

    print(f"Sessions: {len(sess_df)}  Subsessions: {len(sub_df)}")

    with pd.ExcelWriter(OUT, engine="openpyxl") as w:
        sess_df.to_excel(w, sheet_name="Sessions", index=False)
        sub_df.to_excel(w, sheet_name="Subsessions", index=False)

    sess_df.to_csv(OUT.with_suffix(".sessions.csv"), index=False, encoding="utf-8-sig")
    sub_df.to_csv(OUT.with_suffix(".subsessions.csv"), index=False, encoding="utf-8-sig")

    print(f"Wrote: {OUT}")


if __name__ == "__main__":
    main()
