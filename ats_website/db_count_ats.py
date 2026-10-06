from __future__ import annotations

import os

import psycopg


def main() -> None:
    host = os.environ.get("PGHOST", "datazymes-db.c7llxfploany.us-west-2.rds.amazonaws.com")
    port = int(os.environ.get("PGPORT", "5432"))
    db = os.environ.get("PGDATABASE", "intellia_demo_v2")
    user = os.environ.get("PGUSER", "intellia_dev_user")
    pwd = os.environ.get("PGPASSWORD", "")
    conference = os.environ.get("ATS_CONFERENCE", "ATS 2026")

    conninfo = f"host={host} port={port} dbname={db} user={user} password={pwd}"
    with psycopg.connect(conninfo) as conn:
        with conn.cursor() as cur:
            cur.execute('select count(*) from public."Abstract" where "Conference"=%s', (conference,))
            print(cur.fetchone()[0])


if __name__ == "__main__":
    main()

