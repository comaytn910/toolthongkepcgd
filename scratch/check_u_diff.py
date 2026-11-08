import json

with open("scratch/scan_2m2_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]

u_diffs = []
for comm in communes:
    name = comm["name"]
    vals = comm["vals"]
    s = float(vals.get("S") or 0)
    t = float(vals.get("T") or 0)
    u_src = float(vals.get("U") or 0)
    u_calc = s + t
    
    if abs(u_src - u_calc) > 0.01:
        u_diffs.append({
            "name": name,
            "S": s,
            "T": t,
            "U_src": u_src,
            "U_calc": u_calc,
            "diff": u_calc - u_src
        })

print(f"Total communes with U != S + T: {len(u_diffs)}")
for d in u_diffs:
    print(d)

with open("scratch/u_diffs.json", "w", encoding="utf-8") as f:
    json.dump(u_diffs, f, ensure_ascii=False, indent=2)
