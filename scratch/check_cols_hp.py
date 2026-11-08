import json

with open("scratch/scan_2m2_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]

has_h = []
has_p = []

for comm in communes:
    name = comm["name"]
    vals = comm["vals"]
    detail = comm["detail_sums"]
    
    h_src = float(vals.get("H") or 0)
    h_dt = float(detail.get("H") or 0)
    p_src = float(vals.get("P") or 0)
    p_dt = float(detail.get("P") or 0)
    
    if h_src > 0 or h_dt > 0:
        has_h.append({"name": name, "h_src": h_src, "h_detail": h_dt})
    if p_src > 0 or p_dt > 0:
        has_p.append({"name": name, "p_src": p_src, "p_detail": p_dt})

report = {
    "has_h": has_h,
    "has_p": has_p
}

with open("scratch/check_cols_hp.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported check_cols_hp.json successfully!")
