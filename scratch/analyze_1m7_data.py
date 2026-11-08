import json

with open("scratch/scan_1m7_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes_data"]

total_p1_ar = sum(float(c["vals"]["p1_ar"] or 0) for c in communes)
total_p1_sum = sum(float(c["vals"]["sum_p1_caq"] or 0) for c in communes)

total_p2_ar = sum(float(c["vals"]["p2_ar"] or 0) for c in communes)
total_p2_sum = sum(float(c["vals"]["sum_p2_caq"] or 0) for c in communes)

total_tot_ar = sum(float(c["vals"]["tot_ar"] or 0) for c in communes)
total_tot_sum = sum(float(c["vals"]["sum_tot_caq"] or 0) for c in communes)

print(f"Total communes: {len(communes)}")
print(f"P1 (Trong tinh): AR sum = {total_p1_ar}, C..AQ sum = {total_p1_sum}")
print(f"P2 (Ngoai tinh): AR sum = {total_p2_ar}, C..AQ sum = {total_p2_sum}")
print(f"Total (I + II) : AR sum = {total_tot_ar}, C..AQ sum = {total_tot_sum}")

# Kiểm tra các xã có chênh lệch
diff_communes = []
for c in communes:
    name = c["commune"]
    vals = c["vals"]
    p1_diff = abs(float(vals["p1_ar"] or 0) - float(vals["sum_p1_caq"] or 0))
    p2_diff = abs(float(vals["p2_ar"] or 0) - float(vals["sum_p2_caq"] or 0))
    tot_diff = abs(float(vals["tot_ar"] or 0) - float(vals["sum_tot_caq"] or 0))
    
    # Kiểm tra xem tot có bằng p1 + p2 không
    p1_val = float(vals["sum_p1_caq"] or 0)
    p2_val = float(vals["sum_p2_caq"] or 0)
    tot_val = float(vals["sum_tot_caq"] or 0)
    comb_diff = abs(tot_val - (p1_val + p2_val))
    
    if p1_diff > 0.01 or p2_diff > 0.01 or tot_diff > 0.01 or comb_diff > 0.01:
        diff_communes.append({
            "name": name,
            "rows": c["rows"],
            "p1_ar": vals["p1_ar"], "p1_sum": vals["sum_p1_caq"],
            "p2_ar": vals["p2_ar"], "p2_sum": vals["sum_p2_caq"],
            "tot_ar": vals["tot_ar"], "tot_sum": vals["sum_tot_caq"],
            "comb_diff": comb_diff
        })

print(f"\nCommunes with discrepancies: {len(diff_communes)}")
for d in diff_communes:
    print(d)

# Top 10 xã có số lượng học sinh ngoại địa bàn nhiều nhất
top_communes = sorted(communes, key=lambda x: x["vals"]["sum_tot_caq"], reverse=True)
print("\nTop 10 xã có học sinh ngoại địa bàn học tại xã nhiều nhất:")
for c in top_communes[:10]:
    print(f"  {c['commune']}: Tổng = {c['vals']['sum_tot_caq']} (P1 = {c['vals']['sum_p1_caq']}, P2 = {c['vals']['sum_p2_caq']})")
