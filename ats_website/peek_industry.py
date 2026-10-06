import json
from pathlib import Path

f = Path(__file__).parent / "ats_output" / "all_sessions.json"
data = json.load(open(f, encoding="utf-8"))

types_seen = {}
industry_samples = []
exhibitor_count = 0
total = 0
for date, sessions in data.items():
    for s in sessions:
        total += 1
        for c in s.get("categories") or []:
            n = c.get("name") or ""
            types_seen[n] = types_seen.get(n, 0) + 1
        if s.get("exhibitor"):
            exhibitor_count += 1
        cats = [c.get("name") for c in s.get("categories") or []]
        if any("Industry" in x or "Innovation Hub" in x for x in cats) and len(industry_samples) < 5:
            industry_samples.append({
                "date": date,
                "name": s.get("name"),
                "exhibitor": s.get("exhibitor"),
                "categories": cats,
                "sub_count": len(s.get("subsessions") or []),
            })

print(f"Total parents: {total}, with exhibitor: {exhibitor_count}")
print("\nSession types:")
for k, v in sorted(types_seen.items(), key=lambda x: -x[1]):
    print(f"  {v:>4} {k}")
print("\nIndustry samples:")
for s in industry_samples:
    print(json.dumps(s, indent=2))
