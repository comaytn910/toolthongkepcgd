import json

with open("scratch/scan_1m7_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes_data"]

diff_communes = []
for c in communes:
    name = c["commune"]
    vals = c["vals"]
    p1_diff = abs(float(vals["p1_ar"] or 0) - float(vals["sum_p1_caq"] or 0))
    p2_diff = abs(float(vals["p2_ar"] or 0) - float(vals["sum_p2_caq"] or 0))
    tot_diff = abs(float(vals["tot_ar"] or 0) - float(vals["sum_tot_caq"] or 0))
    
    p1_val = float(vals["sum_p1_caq"] or 0)
    p2_val = float(vals["sum_p2_caq"] or 0)
    tot_val = float(vals["sum_tot_caq"] or 0)
    comb_diff = abs(tot_val - (p1_val + p2_val))
    
    if p1_diff > 0.01 or p2_diff > 0.01 or tot_diff > 0.01 or comb_diff > 0.01:
        diff_communes.append({
            "name": name,
            "rows": c["rows"],
            "vals": vals,
            "comb_diff": comb_diff
        })

with open("scratch/1m7_discrepancies.json", "w", encoding="utf-8") as f:
    json.dump(diff_communes, f, ensure_ascii=False, indent=2)

print(f"Exported discrepancies to scratch/1m7_discrepancies.json")
