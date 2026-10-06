from __future__ import annotations

import os

import psycopg


def main() -> None:
    host = os.environ.get("PGHOST", "datazymes-db.c7llxfploany.us-west-2.rds.amazonaws.com")
    port = int(os.environ.get("PGPORT", "5432"))
    db = os.environ.get("PGDATABASE", "intellia_demo_v2")
    user = os.environ.get("PGUSER", "intellia_dev_user")
    pwd = os.environ.get("PGPASSWORD", "")

    conninfo = f"host={host} port={port} dbname={db} user={user} password={pwd}"

    with psycopg.connect(conninfo) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                select conname, contype, pg_get_constraintdef(pg_constraint.oid)
                from pg_constraint
                where conrelid = 'public."Abstract"'::regclass
                order by contype, conname
                """
            )
            cons = cur.fetchall()

            cur.execute(
                """
                select indexname, indexdef
                from pg_indexes
                where schemaname = %s and tablename = %s
                order by indexname
                """,
                ("public", "Abstract"),
            )
            idx = cur.fetchall()

    print("Constraints:")
    for name, typ, definition in cons:
        print(f"- {name} ({typ}): {definition}")

    print("\nIndexes:")
    for name, definition in idx:
        print(f"- {name}: {definition}")


if __name__ == "__main__":
    main()

