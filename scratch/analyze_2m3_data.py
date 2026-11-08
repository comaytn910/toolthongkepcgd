import json

with open("scratch/scan_2m3_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

communes = data["communes"]
COUNT_COLS = ['C', 'D', 'E', 'H', 'I', 'K', 'M', 'N', 'O']

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

prov_tot_src = {c: 0.0 for c in COUNT_COLS}
prov_tot_detail = {c: 0.0 for c in COUNT_COLS}
discrepancies = []

for comm in communes:
    name = comm["name"]
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
            "commune": name,
            "tot_row": tot_row,
            "diffs": comm_diff
        })

def calc_rate(num, den):
    return (num / den * 100) if den > 0 else 0.0

rates_src = {
    "J_rate_TN_THCS": calc_rate(prov_tot_src['I'], prov_tot_src['H']),
    "L_rate_hoc_tiep": calc_rate(prov_tot_src['K'], prov_tot_src['H']),
    "P_rate_KT_tiep_can": calc_rate(prov_tot_src['O'], prov_tot_src['N'])
}

rates_detail = {
    "J_rate_TN_THCS": calc_rate(prov_tot_detail['I'], prov_tot_detail['H']),
    "L_rate_hoc_tiep": calc_rate(prov_tot_detail['K'], prov_tot_detail['H']),
    "P_rate_KT_tiep_can": calc_rate(prov_tot_detail['O'], prov_tot_detail['N'])
}

# Thống kê cột S (Đạt chuẩn)
s_counts = {}
for comm in communes:
    s_val = str(comm["vals"].get("S") or "").strip()
    if not s_val:
        s_val = "[Trống]"
    s_counts[s_val] = s_counts.get(s_val, 0) + 1

report = {
    "total_communes": len(communes),
    "prov_tot_src": prov_tot_src,
    "prov_tot_detail": prov_tot_detail,
    "rates_src": rates_src,
    "rates_detail": rates_detail,
    "s_standard_counts": s_counts,
    "discrepancies_count": len(discrepancies),
    "discrepancies": discrepancies
}

with open("scratch/analyze_2m3_report.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported analyze_2m3_report.json successfully!")
