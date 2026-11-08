import json

with open("scratch/scan_2m2_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]
COUNT_COLS = ['C', 'D', 'F', 'G', 'H', 'I', 'K', 'L', 'N', 'O', 'P', 'R', 'S', 'T', 'U']

def clean_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(',', ''))
        except ValueError:
            return 0.0
    return 0.0

# 1. So sánh giữa Hàng Tổng nguồn và Tổng chi tiết từng xã
discrepancies = []
prov_tot_src = {c: 0.0 for c in COUNT_COLS}
prov_tot_detail = {c: 0.0 for c in COUNT_COLS}

for comm in communes:
    cname = comm["name"]
    tot_row = comm["tot_row"]
    vals = comm["vals"]
    detail = comm["detail_sums"]
    
    comm_diff = {}
    for c in COUNT_COLS:
        src_v = clean_num(vals.get(c))
        dt_v = clean_num(detail.get(c))
        prov_tot_src[c] += src_v
        prov_tot_detail[c] += dt_v
        
        if abs(src_v - dt_v) > 0.01:
            comm_diff[c] = {"src": src_v, "detail": dt_v, "diff": src_v - dt_v}
            
    if comm_diff:
        discrepancies.append({
            "commune": cname,
            "tot_row": tot_row,
            "diffs": comm_diff
        })

# 2. Tính tỷ lệ toàn tỉnh
def calc_rate(num, den):
    return (num / den * 100) if den > 0 else 0.0

rates_src = {
    "E_rate_6t_L1": calc_rate(prov_tot_src['D'], prov_tot_src['C']),
    "J_rate_vao_L6": calc_rate(prov_tot_src['I'], prov_tot_src['F']),
    "M_rate_11_14t_HTCTTH": calc_rate(prov_tot_src['L'], prov_tot_src['K']),
    "Q_rate_TN_THCS_lop9": calc_rate(prov_tot_src['O'] + prov_tot_src['P'], prov_tot_src['N']),
    "V_rate_TN_THCS_15_18t": calc_rate(prov_tot_src['U'], prov_tot_src['R'])
}

rates_detail = {
    "E_rate_6t_L1": calc_rate(prov_tot_detail['D'], prov_tot_detail['C']),
    "J_rate_vao_L6": calc_rate(prov_tot_detail['I'], prov_tot_detail['F']),
    "M_rate_11_14t_HTCTTH": calc_rate(prov_tot_detail['L'], prov_tot_detail['K']),
    "Q_rate_TN_THCS_lop9": calc_rate(prov_tot_detail['O'] + prov_tot_detail['P'], prov_tot_detail['N']),
    "V_rate_TN_THCS_15_18t": calc_rate(prov_tot_detail['U'], prov_tot_detail['R'])
}

# 3. Thống kê cột W: Đạt chuẩn PCGD THCS của các xã
w_counts = {}
for comm in communes:
    w_val = str(comm["vals"].get("W") or "").strip()
    if not w_val:
        w_val = "[Trống]"
    w_counts[w_val] = w_counts.get(w_val, 0) + 1

report = {
    "total_communes": len(communes),
    "prov_tot_src": prov_tot_src,
    "prov_tot_detail": prov_tot_detail,
    "rates_src": rates_src,
    "rates_detail": rates_detail,
    "w_standard_counts": w_counts,
    "discrepancies_count": len(discrepancies),
    "discrepancies": discrepancies
}

with open("scratch/analyze_2m2_report.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported analyze_2m2_report.json successfully!")
