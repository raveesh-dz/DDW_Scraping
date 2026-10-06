import json
from pathlib import Path

f = Path(__file__).parent / "ats_output" / "sessions_2026-05-17.json"
data = json.load(open(f, encoding="utf-8"))

# Find Mini Symposium and Poster Discussion sample
for s in data:
    for sub in s.get("subsessions") or []:
        if sub.get("activityType") == "Abstract":
            print("=== PARENT NAME ===", s.get("name"))
            print("=== SUB ABSTRACT ===")
            print(json.dumps(sub, indent=2)[:3000])
            break
    else:
        continue
    break

# Distinct activityTypes
types = set()
for s in data:
    for sub in s.get("subsessions") or []:
        types.add(sub.get("activityType"))
print("\nactivityTypes:", types)
