import json

with open("scratch/scan_2m2_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]

# 1. Kiểm tra I != G + H
i_diffs = []
for comm in communes:
    name = comm["name"]
    vals = comm["vals"]
    g = float(vals.get("G") or 0)
    h = float(vals.get("H") or 0)
    i_src = float(vals.get("I") or 0)
    if abs(i_src - (g + h)) > 0.01:
        i_diffs.append({"name": name, "G": g, "H": h, "I_src": i_src, "I_calc": g + h, "diff": (g + h) - i_src})

print(f"Total communes with I != G + H: {len(i_diffs)}")
for d in i_diffs:
    print(d)

with open("scratch/i_diffs.json", "w", encoding="utf-8") as f:
    json.dump(i_diffs, f, ensure_ascii=False, indent=2)
