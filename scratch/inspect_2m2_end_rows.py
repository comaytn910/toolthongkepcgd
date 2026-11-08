import json

with open("scratch/inspect_2m2_template.json", "r", encoding="utf-8") as f:
    data = json.load(f)

rows = data["rows"]
print(f"Total non-empty rows: {len(rows)}")

for r_str in sorted(rows.keys(), key=lambda x: int(x)):
    r = int(r_str)
    r_data = rows[r_str]
    a_val = r_data.get("A", {}).get("val")
    b_val = r_data.get("B", {}).get("val")
    c_val = r_data.get("C", {}).get("val")
    if r > 30:
        print(f"Row {r:3d}: A={a_val}, B={b_val}, C={c_val}")
