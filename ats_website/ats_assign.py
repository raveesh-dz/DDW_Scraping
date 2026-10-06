"""Assign attendees to ATS sessions, parent-level dedupe, priority-aware.

Reads:
  - all_sessions_copd_bronchiectasis.json (raw scraped data)
  - Molecule_List_COPD_NCFBE(Sheet1).csv (priority molecules)

Writes:
  - ATS_COPD_Bronchiectasis_2026_Assigned.xlsx
      sheet "All Sessions" — every abstract row + Priority + Attendee cols
      sheet "Parent Sessions" — deduped parent-level assignments
      sheets "Attendee 1".."Attendee N" — per-attendee schedule

Match rules:
  - Match molecule name, brand name (parens), originator/company in:
      parent.name, parent.description, child.name, child.description, Company
  - Case-insensitive substring; word-boundary for short tokens (≤4 chars)
  - NCFBE molecules → also tag if parent has Bronchiectasis topic

Priority tiers per parent:
  - "High"   → matched a molecule whose Priority=='High'
  - "Medium" → matched any molecule (no High)
  - ""       → no match

Attendee assignment:
  - Sort by tier (High, Medium, Low/empty), then by start time
  - Greedy: lowest-id attendee whose last_end <= start gets it; else allocate new
"""
from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from ats_db import PgConfig, upsert_abstract_rows_by_url

ROOT = Path(__file__).parent
SRC = ROOT / "ats_output" / "all_sessions_copd_bronchiectasis.json"
PRIORITY_CSV = ROOT / "Molecule_List_COPD_NCFBE(Sheet1).csv"
EXPORT_XLSX = ROOT / "ats_output" / "ATS_COPD_Bronchiectasis_2026.xlsx"
OUT = ROOT / "ats_output" / "ATS_COPD_Bronchiectasis_2026_Assigned.xlsx"

POSTER_RE = re.compile(r"\(\s*Poster\s*Board\s*#\s*([A-Za-z0-9\-]+)\s*\)", re.I)
SESSION_CODE_RE = re.compile(r"^([A-Z]{1,3}\d{1,4}[A-Z]?)\b")
PARENS_RE = re.compile(r"\(([^)]+)\)")

ABSTRACT_PARENT_TYPES = {
    "Mini Symposium", "Poster Discussion Session", "Thematic Poster Session",
    "RAPiD Poster Discussion Session", "RAPiD Abstract Session",
    "Scientific Symposium", "Late Breaking Mini Symposium",
}
COMPANY_LED_TYPES = {
    "Industry Theater", "Innovation Hub",
    "Industry Supported Evening Symposia", "Mid-day Symposium", "Product Theater",
}


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
    cmp_ = [t for t in types if t in COMPANY_LED_TYPES]
    if cmp_:
        return cmp_[0]
    return types[0] if types else ""


def indication_for(parent: dict) -> str:
    names = [c.get("name") for c in parent.get("categories") or []]
    matched = [n for n in ("COPD", "Bronchiectasis") if n in names]
    return " / ".join(matched)


# ---------- priority matching ----------

def load_priority_terms(csv_path: Path) -> list[dict]:
    """Return list of {tokens: [str], originator: str, molecule: str, indication: str, priority: str}."""
    df = pd.read_csv(csv_path)
    df = df.dropna(subset=["Molecule"]).reset_index(drop=True)
    out = []
    for _, r in df.iterrows():
        mol = str(r["Molecule"]).strip()
        orig = str(r.get("Originator") or "").strip()
        ind = str(r.get("Indication") or "").strip().upper()
        pri = str(r.get("Priority") or "").strip()
        # extract brand + generic from parens, e.g. "Nucala (Mepolizumab)"
        tokens = set()
        if mol:
            base = PARENS_RE.sub("", mol).strip(" /,")
            for t in re.split(r"[/,]", base):
                t = t.strip()
                if t:
                    tokens.add(t)
            for inner in PARENS_RE.findall(mol):
                for t in re.split(r"[/,]", inner):
                    t = t.strip()
                    if t:
                        tokens.add(t)
        # originator can be multiple ("Sanofi/Regeneron")
        for t in re.split(r"[/,]", orig):
            t = t.strip()
            if t:
                tokens.add(t)
        # drop very short tokens (noise)
        tokens = {t for t in tokens if len(t) >= 3}
        out.append({
            "tokens": sorted(tokens, key=len, reverse=True),
            "molecule": mol,
            "originator": orig,
            "indication": ind,
            "priority": pri.lower() == "high",
        })
    return out


def build_matchers(terms: list[dict]) -> list[tuple[re.Pattern, dict]]:
    """Compile per-term regex (word-boundary)."""
    out = []
    for t in terms:
        pats = []
        for tok in t["tokens"]:
            esc = re.escape(tok)
            pats.append(re.compile(rf"\b{esc}\b", re.I))
        out.append((pats, t))
    return out


def match_priority(text_blob: str, indication_tag: str, matchers: list) -> tuple[str, list[str], list[str]]:
    """Return (priority_label, matched_molecules, matched_originators)."""
    if not text_blob:
        return "", [], []
    is_brc = "Bronchiectasis" in indication_tag
    is_copd = "COPD" in indication_tag
    high_mols, all_mols, origs = [], [], []
    for pats, t in matchers:
        # NCFBE entries match Bronchiectasis-tagged sessions; COPD matches COPD-tagged
        if t["indication"] == "NCFBE" and not is_brc:
            continue
        if t["indication"] == "COPD" and not is_copd:
            continue
        if any(p.search(text_blob) for p in pats):
            all_mols.append(t["molecule"])
            if t["originator"]:
                origs.append(t["originator"])
            if t["priority"]:
                high_mols.append(t["molecule"])
    label = "High" if high_mols else ("Medium" if all_mols else "")
    return label, sorted(set(all_mols)), sorted(set(origs))


# ---------- main flow ----------

def build_parent_records(all_data: dict, matchers: list) -> list[dict]:
    """One record per parent session w/ priority + interval + child abstract list."""
    records = []
    for date, sessions in all_data.items():
        for parent in sessions:
            stype = session_type_of(parent)
            code = session_code(parent.get("name", ""))
            location = parent_location(parent.get("geozonesOfEdition"))
            ind = indication_for(parent)
            company = company_name(parent)

            abstract_subs = [
                s for s in (parent.get("subsessions") or [])
                if s.get("activityType") == "Abstract"
            ]

            text_parts = [
                parent.get("name") or "",
                parent.get("description") or "",
                company,
            ]
            for sub in abstract_subs:
                text_parts.append(sub.get("name") or "")
                text_parts.append(sub.get("description") or "")
            blob = " \n ".join(text_parts)

            priority, mols, origs = match_priority(blob, ind, matchers)

            start_iso = parent.get("startTime") or ""
            end_iso = parent.get("endTime") or ""
            try:
                start_dt = datetime.fromisoformat(start_iso) if start_iso else None
                end_dt = datetime.fromisoformat(end_iso) if end_iso else None
            except ValueError:
                start_dt = end_dt = None

            records.append({
                "parent_uid": parent.get("uid"),
                "parent_name": parent.get("name"),
                "company": company,
                "indication": ind,
                "session_type": stype,
                "session_code": code,
                "location": location,
                "date_str": fmt_date(start_iso),
                "time_str": fmt_time(start_iso, end_iso),
                "start_iso": start_iso,
                "end_iso": end_iso,
                "start_dt": start_dt,
                "end_dt": end_dt,
                "priority": priority,
                "matched_molecules": ", ".join(mols),
                "matched_originators": ", ".join(origs),
                "abstract_subs": abstract_subs,
            })
    return records


def assign_attendees(records: list[dict]) -> int:
    """Optimal interval coloring (= max concurrency) then renumber so attendees
    with most High-priority sessions get lowest IDs.

    Sort by start time (within ties, High first → ties resolved deterministically).
    Greedy lowest-free-id pick yields N = peak concurrency.
    Then permute IDs so High count desc → 1..N.
    """
    tier_rank = {"High": 0, "Medium": 1, "": 2}
    sortable = [r for r in records if r["start_dt"] and r["end_dt"]]
    sortable.sort(key=lambda r: (r["start_dt"], tier_rank.get(r["priority"], 3)))

    last_end: list[datetime] = []  # idx i = raw-id i's last end_dt

    for r in sortable:
        chosen = -1
        for i, le in enumerate(last_end):
            if le <= r["start_dt"]:
                chosen = i
                break
        if chosen == -1:
            last_end.append(r["end_dt"])
            chosen = len(last_end) - 1
        else:
            last_end[chosen] = r["end_dt"]
        r["_raw_attendee"] = chosen

    n = len(last_end)

    # tally High count per raw-id
    high_count = [0] * n
    medium_count = [0] * n
    for r in sortable:
        i = r["_raw_attendee"]
        if r["priority"] == "High":
            high_count[i] += 1
        elif r["priority"] == "Medium":
            medium_count[i] += 1

    # rank raw-ids by (-high, -medium) so most-High → final ID 1
    order = sorted(range(n), key=lambda i: (-high_count[i], -medium_count[i], i))
    remap = {raw: final + 1 for final, raw in enumerate(order)}

    for r in sortable:
        r["attendee"] = remap[r.pop("_raw_attendee")]

    for r in records:
        r.setdefault("attendee", None)

    return n


def main():
    print(f"Reading: {SRC.name}")
    with open(SRC, encoding="utf-8") as f:
        all_data = json.load(f)

    print(f"Reading priority list: {PRIORITY_CSV.name}")
    terms = load_priority_terms(PRIORITY_CSV)
    print(f"  {len(terms)} priority entries ({sum(1 for t in terms if t['priority'])} High)")

    matchers = build_matchers(terms)

    records = build_parent_records(all_data, matchers)
    print(f"Parent sessions: {len(records)}")

    pri_counts = pd.Series([r["priority"] or "Low" for r in records]).value_counts().to_dict()
    print(f"Priority distribution: {pri_counts}")

    n_attendees = assign_attendees(records)
    print(f"Attendees needed: {n_attendees}")

    # --- build All Sessions sheet (abstract rows expanded) ---
    all_rows = []
    for r in records:
        attendee = f"Attendee {r['attendee']}" if r["attendee"] else ""
        if r["abstract_subs"]:
            for sub in r["abstract_subs"]:
                poster, title = split_poster(sub.get("name", ""))
                all_rows.append({
                    "Company": r["company"],
                    "Laterals": "",
                    "Date": fmt_date(sub.get("startTime") or r["start_iso"]) or r["date_str"],
                    "Time (EST)": fmt_time(sub.get("startTime", ""), sub.get("endTime", "")),
                    "Start Time": sub.get("startTime") or "",
                    "End Time": sub.get("endTime") or "",
                    "Session Type": r["session_type"],
                    "Poster (if applicable)": poster,
                    "Indication": r["indication"],
                    "Abstract Title": title,
                    "Session Code": r["session_code"],
                    "Location": r["location"],
                    "Notes": "",
                    "Priority": r["priority"],
                    "Attendee": attendee,
                    "Matched Molecules": r["matched_molecules"],
                    "Matched Originators": r["matched_originators"],
                    "Parent Session": r["parent_name"],
                    "Parent UID": r["parent_uid"],
                    "Subsession UID": sub.get("uid"),
                    "URL": f"https://ats2026.d365.events/education/sessions/{sub.get('uid')}",
                })
        else:
            all_rows.append({
                "Company": r["company"],
                "Laterals": "",
                "Date": r["date_str"],
                "Time (EST)": r["time_str"],
                "Start Time": r["start_iso"],
                "End Time": r["end_iso"],
                "Session Type": r["session_type"],
                "Poster (if applicable)": "",
                "Indication": r["indication"],
                "Abstract Title": r["parent_name"],
                "Session Code": r["session_code"],
                "Location": r["location"],
                "Notes": "",
                "Priority": r["priority"],
                "Attendee": attendee,
                "Matched Molecules": r["matched_molecules"],
                "Matched Originators": r["matched_originators"],
                "Parent Session": r["parent_name"],
                "Parent UID": r["parent_uid"],
                "Subsession UID": "",
                "URL": f"https://ats2026.d365.events/education/sessions/{r['parent_uid']}",
            })

    df_all = pd.DataFrame(all_rows)

    # --- Parent Sessions sheet ---
    parent_rows = []
    for r in records:
        parent_rows.append({
            "Attendee": f"Attendee {r['attendee']}" if r["attendee"] else "",
            "Priority": r["priority"],
            "Date": r["date_str"],
            "Time (EST)": r["time_str"],
            "Start Time": r["start_iso"],
            "End Time": r["end_iso"],
            "Session Code": r["session_code"],
            "Session Type": r["session_type"],
            "Indication": r["indication"],
            "Parent Session": r["parent_name"],
            "Company": r["company"],
            "Location": r["location"],
            "Matched Molecules": r["matched_molecules"],
            "Matched Originators": r["matched_originators"],
            "Parent UID": r["parent_uid"],
            "URL": f"https://ats2026.d365.events/education/sessions/{r['parent_uid']}",
        })
    df_parent = pd.DataFrame(parent_rows).sort_values(
        ["Attendee", "Start Time"], na_position="last"
    )

    # --- write to DB instead of Excel (set ATS_WRITE_DB=1) ---
    if str(os.environ.get("ATS_WRITE_DB", "")).strip() in {"1", "true", "TRUE", "yes", "YES"}:
        pg = PgConfig.from_env(default_dbname="intellia_demo_v2")
        res = upsert_abstract_rows_by_url(
            df_all.to_dict(orient="records"),
            pg=pg,
            conference=os.environ.get("ATS_CONFERENCE", "ATS 2026"),
        )
        print(f'Wrote to DB table public."Abstract" (by URL). {res}')
        print(f"  All Sessions rows: {len(df_all)}")
        print(f"  Parent Sessions rows: {len(df_parent)}")
        return

    # --- per-attendee tabs (default behavior) ---
    with pd.ExcelWriter(OUT, engine="openpyxl") as w:
        df_all.to_excel(w, sheet_name="All Sessions", index=False)
        df_parent.to_excel(w, sheet_name="Parent Sessions", index=False)
        for i in range(1, n_attendees + 1):
            sub = df_parent[df_parent["Attendee"] == f"Attendee {i}"].sort_values("Start Time")
            sub.to_excel(w, sheet_name=f"Attendee {i}", index=False)

    print(f"Wrote: {OUT}")
    print(f"  All Sessions rows: {len(df_all)}")
    print(f"  Parent Sessions rows: {len(df_parent)}")


if __name__ == "__main__":
    main()
