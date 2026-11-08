import json

with open("scratch/inspect_2m3_template.json", "r", encoding="utf-8") as f:
    data = json.load(f)

rows = data["rows"]
report = {}
for r_str, r_data in rows.items():
    r = int(r_str)
    if r >= 35:
        row_brief = {}
        for c, c_info in r_data.items():
            row_brief[c] = c_info.get("val") or c_info.get("formula")
        report[f"Row_{r}"] = row_brief

with open("scratch/inspect_2m3_end.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported 2.M3 end rows!")
