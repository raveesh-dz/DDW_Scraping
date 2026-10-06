from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Optional

import psycopg
from psycopg.rows import dict_row


@dataclass(frozen=True)
class PgConfig:
    host: str
    port: int
    dbname: str
    user: str
    password: str

    @staticmethod
    def from_env(
        *,
        default_dbname: str = "intellia_demo_v2",
        default_host: str = "datazymes-db.c7llxfploany.us-west-2.rds.amazonaws.com",
        default_port: int = 5432,
        default_user: str = "intellia_dev_user",
    ) -> "PgConfig":
        return PgConfig(
            host=os.environ.get("PGHOST", default_host),
            port=int(os.environ.get("PGPORT", str(default_port))),
            dbname=os.environ.get("PGDATABASE", default_dbname),
            user=os.environ.get("PGUSER", default_user),
            password=os.environ.get("PGPASSWORD", ""),
        )

    def conninfo(self) -> str:
        return (
            f"host={self.host} port={self.port} dbname={self.dbname} "
            f"user={self.user} password={self.password}"
        )


def _parse_iso_to_naive_dt(value: str | None) -> Optional[datetime]:
    if not value:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Handle Zulu suffix and offsets.
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(tz=None).replace(tzinfo=None)
    return dt


def _uuid_or_none(value: object) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    return s or None


def upsert_abstract_rows_by_url(
    rows: Iterable[dict],
    *,
    pg: PgConfig,
    conference: str = "ATS 2026",
) -> dict:
    """
    Upsert into public."Abstract" using URL as the natural key.

    Note: The DB schema currently only has a PK on Id, so we do a two-phase approach:
      - fetch existing URL -> Id mapping
      - update rows with known URL, insert rows with new URL
    """
    rows = [r for r in rows if r.get("URL")]
    if not rows:
        return {"inserted": 0, "updated": 0, "skipped": 0}

    urls = [str(r["URL"]).strip() for r in rows if str(r.get("URL") or "").strip()]
    urls = list(dict.fromkeys(urls))
    if not urls:
        return {"inserted": 0, "updated": 0, "skipped": len(list(rows))}

    inserted = 0
    updated = 0

    with psycopg.connect(pg.conninfo(), row_factory=dict_row) as conn:
        with conn.cursor() as cur:
            cur.execute(
                'select "Id", "URL" from public."Abstract" where "URL" = any(%s)',
                (urls,),
            )
            existing = {r["URL"]: r["Id"] for r in cur.fetchall()}

            def as_record(r: dict) -> dict:
                return {
                    "Conference": conference,
                    "Company": (r.get("Company") or "") or None,
                    "Laterals": (r.get("Laterals") or "") or None,
                    "StartTime": _parse_iso_to_naive_dt(r.get("Start Time") or r.get("Start Time")),
                    "EndTime": _parse_iso_to_naive_dt(r.get("End Time") or r.get("End Time")),
                    "SessionType": (r.get("Session Type") or "") or None,
                    "Poster": (r.get("Poster (if applicable)") or "") or None,
                    "Indication": (r.get("Indication") or "") or None,
                    "AbstractTitle": (r.get("Abstract Title") or "") or None,
                    "SessionCode": (r.get("Session Code") or "") or None,
                    "Location": (r.get("Location") or "") or None,
                    "Notes": (r.get("Notes") or "") or None,
                    "Priority": (r.get("Priority") or "") or None,
                    "Attendee": (r.get("Attendee") or "") or None,
                    "MatchedMolecules": (r.get("Matched Molecules") or "") or None,
                    "MatchedOriginators": (r.get("Matched Originators") or "") or None,
                    "ParentSession": (r.get("Parent Session") or "") or None,
                    "ParentUID": _uuid_or_none(r.get("Parent UID")),
                    "SubsessionUID": _uuid_or_none(r.get("Subsession UID")),
                    "URL": (r.get("URL") or "") or None,
                    "AttachedFile": (r.get("AttachedFile") or r.get("Attached File") or "") or None,
                    "AISummary": (r.get("AISummary") or r.get("AI Summary") or "") or None,
                    "AbstractBody": (r.get("AbstractBody") or r.get("Abstract Body") or "") or None,
                }

            to_update = []
            to_insert = []
            for r in rows:
                url = str(r.get("URL") or "").strip()
                if not url:
                    continue
                rec = as_record(r)
                if url in existing:
                    rec["Id"] = existing[url]
                    to_update.append(rec)
                else:
                    to_insert.append(rec)

            if to_update:
                cur.executemany(
                    """
                    update public."Abstract"
                    set
                      "Conference"=%(Conference)s,
                      "Company"=%(Company)s,
                      "Laterals"=%(Laterals)s,
                      "StartTime"=%(StartTime)s,
                      "EndTime"=%(EndTime)s,
                      "SessionType"=%(SessionType)s,
                      "Poster"=%(Poster)s,
                      "Indication"=%(Indication)s,
                      "AbstractTitle"=%(AbstractTitle)s,
                      "SessionCode"=%(SessionCode)s,
                      "Location"=%(Location)s,
                      "Notes"=%(Notes)s,
                      "Priority"=%(Priority)s,
                      "Attendee"=%(Attendee)s,
                      "MatchedMolecules"=%(MatchedMolecules)s,
                      "MatchedOriginators"=%(MatchedOriginators)s,
                      "ParentSession"=%(ParentSession)s,
                      "ParentUID"=%(ParentUID)s::uuid,
                      "SubsessionUID"=%(SubsessionUID)s::uuid,
                      "AttachedFile"=%(AttachedFile)s,
                      "AISummary"=%(AISummary)s,
                      "AbstractBody"=%(AbstractBody)s
                    where "Id"=%(Id)s
                    """,
                    to_update,
                )
                updated = len(to_update)

            if to_insert:
                cur.executemany(
                    """
                    insert into public."Abstract" (
                      "Conference",
                      "Company",
                      "Laterals",
                      "StartTime",
                      "EndTime",
                      "SessionType",
                      "Poster",
                      "Indication",
                      "AbstractTitle",
                      "SessionCode",
                      "Location",
                      "Notes",
                      "Priority",
                      "Attendee",
                      "MatchedMolecules",
                      "MatchedOriginators",
                      "ParentSession",
                      "ParentUID",
                      "SubsessionUID",
                      "URL",
                      "AttachedFile",
                      "AISummary",
                      "AbstractBody"
                    ) values (
                      %(Conference)s,
                      %(Company)s,
                      %(Laterals)s,
                      %(StartTime)s,
                      %(EndTime)s,
                      %(SessionType)s,
                      %(Poster)s,
                      %(Indication)s,
                      %(AbstractTitle)s,
                      %(SessionCode)s,
                      %(Location)s,
                      %(Notes)s,
                      %(Priority)s,
                      %(Attendee)s,
                      %(MatchedMolecules)s,
                      %(MatchedOriginators)s,
                      %(ParentSession)s,
                      %(ParentUID)s::uuid,
                      %(SubsessionUID)s::uuid,
                      %(URL)s,
                      %(AttachedFile)s,
                      %(AISummary)s,
                      %(AbstractBody)s
                    )
                    """,
                    to_insert,
                )
                inserted = len(to_insert)

        conn.commit()

    return {"inserted": inserted, "updated": updated, "skipped": 0}

