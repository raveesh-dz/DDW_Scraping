"""ATS 2026 scraper filtered by Topics: COPD + Bronchiectasis."""
import json
import time
from pathlib import Path

import requests

API = "https://aggregator.d365.events/graphqlb"
TOKEN_URL = "https://aggregator.d365.events/guest/token?domain=ats2026.d365.events"
DATES = [f"2026-05-{d:02d}" for d in range(15, 21)]

TOPIC_UIDS = [
    "04a08343-68f1-477a-b688-a8700b1fca73",
    "14526bed-bd2f-4692-9cd5-a6c3e7e9e1a4",
]

OUT_DIR = Path(__file__).parent / "ats_output"
OUT_DIR.mkdir(exist_ok=True)

SESSION_PROGRAM_FRAGMENT = """
fragment sessionProgramFragment on Session {
  uid name startTime startTimestamp endTime endTimestamp description featured
  exhibitor { uid name }
  sponsorRoles { uid role sponsor { uid name } }
  categoriesInContext(context: "session_program") {
    uid schemaCode domain name displaySequence
  }
  categories { uid schemaCode domain name }
  speakerRoles(orderBy: displaySequence_asc) {
    uid role displaySequence
    speaker { uid firstName lastName name prefix suffix jobTitle employerName }
  }
  geozonesOfEdition: _geozonesOfEdition { name exhibitHall { name } }
}
"""

SUBSESSION_FRAGMENT = """
fragment subsessionFragment on Session {
  uid name startTime startTimestamp endTime endTimestamp description
  activityType
  geozone { uid name }
  categories { uid domain name }
  exhibitor { uid name }
  sponsorRoles { uid role sponsor { uid name } }
  speakerRoles(orderBy: displaySequence_asc) {
    uid role displaySequence
    speaker { uid firstName lastName name prefix suffix jobTitle employerName }
  }
}
"""

PROGRAM_QUERY = """
query LoadSessionsByDate($filter: _SessionFilter, $orderBy: [_SessionOrdering!]) {
  sessions: session(filter: $filter, orderBy: $orderBy) { ...sessionProgramFragment }
}
""" + SESSION_PROGRAM_FRAGMENT

SUB_QUERY = """
query LoadSubsessions($parentUid: ID!) {
  subsessions: session(
    filter: {parentSession: {uid: $parentUid}}
    orderBy: [displaySequence_asc, startTimestamp_asc]
  ) { ...subsessionFragment }
}
""" + SUBSESSION_FRAGMENT


def get_token() -> str:
    r = requests.post(TOKEN_URL, headers={"accept": "application/json"})
    r.raise_for_status()
    return r.json()["token"]


def gql(token: str, query: str, variables: dict, op_name: str) -> dict:
    payload = {"operationName": op_name, "variables": variables, "query": query}
    r = requests.post(
        API,
        headers={"jwt-token": token, "content-type": "application/json", "accept": "*/*"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()
    data = r.json()
    if "errors" in data:
        raise RuntimeError(f"GraphQL errors for {op_name}: {data['errors']}")
    return data["data"]


def fetch_sessions_for_date(token: str, date: str) -> list[dict]:
    filt = {
        "schemaCode": "ats2026365",
        "parentSession": None,
        "categories_none": {"name_in": ["Conference Event", "Meeting", "Reception"]},
        "AND": [{"categories": {"uid_in": TOPIC_UIDS}}],
        "_inStartDate": {"dates": [date], "timezone": "US/Eastern"},
    }
    data = gql(
        token,
        PROGRAM_QUERY,
        {"filter": filt, "orderBy": ["startTime_asc"]},
        "LoadSessionsByDate",
    )
    return data.get("sessions") or []


def fetch_subsessions(token: str, parent_uid: str) -> list[dict]:
    data = gql(token, SUB_QUERY, {"parentUid": parent_uid}, "LoadSubsessions")
    return data.get("subsessions") or []


def main():
    token = get_token()
    print(f"Got token: {token[:30]}...")

    all_data = {}
    total_sessions = 0
    total_subs = 0

    for date in DATES:
        print(f"\n=== {date} ===")
        sessions = fetch_sessions_for_date(token, date)
        print(f"  {len(sessions)} top-level sessions matching topics")
        total_sessions += len(sessions)

        enriched = []
        for i, s in enumerate(sessions, 1):
            uid = s["uid"]
            try:
                subs = fetch_subsessions(token, uid)
            except Exception as e:
                print(f"  [{i}/{len(sessions)}] ERR {uid}: {e}")
                subs = []
            s["subsessions"] = subs
            total_subs += len(subs)
            enriched.append(s)
            time.sleep(0.05)

        all_data[date] = enriched

    with open(OUT_DIR / "all_sessions_copd_bronchiectasis.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)

    print(f"\n=== DONE === parents={total_sessions} subsessions={total_subs}")


if __name__ == "__main__":
    main()
