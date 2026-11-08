import json
import openpyxl
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

AGE_COL_GROUPS = {
    '6 Tuổi': ['C', 'D', 'E'],
    '7 Tuổi': ['F', 'G', 'H', 'I'],
    '8 Tuổi': ['J', 'K', 'L', 'M', 'N'],
    '9 Tuổi': ['O', 'P', 'Q', 'R', 'S', 'T'],
    '10 Tuổi': ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA'],
    '11 Tuổi': ['AB', 'AC', 'AD', 'AE'],
    '12 Tuổi': ['AF', 'AG', 'AH', 'AI'],
    '13 Tuổi': ['AJ', 'AK', 'AL', 'AM'],
    '14 Tuổi': ['AN', 'AO', 'AP', 'AQ']
}

with open("scratch/scan_1m7_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Phân tích theo từng cột và từng nhóm tuổi
from core.fast_excel import fast_read_sheet
from pathlib import Path

files = sorted([f for f in Path("data").glob("*.xlsx") if not f.name.startswith("~$")])

ALL_COLS = [openpyxl.utils.get_column_letter(c) for c in range(3, 44)] # C đến AQ

prov_p1 = {c: 0.0 for c in ALL_COLS}
prov_p2 = {c: 0.0 for c in ALL_COLS}
prov_tot = {c: 0.0 for c in ALL_COLS}

for comm in data["communes_data"]:
    cname = comm["commune"]
    r_p1 = comm["rows"]["p1"]
    r_p2 = comm["rows"]["p2"]
    
    fpath = Path("data") / f"{cname}.xlsx"
    cells = fast_read_sheet(str(fpath), "1.M7")
    
    if r_p1:
        for c in ALL_COLS:
            val = float(cells.get(f"{c}{r_p1}", {}).get('val') or 0)
            prov_p1[c] += val
            
    if r_p2:
        for c in ALL_COLS:
            val = float(cells.get(f"{c}{r_p2}", {}).get('val') or 0)
            prov_p2[c] += val

for c in ALL_COLS:
    prov_tot[c] = prov_p1[c] + prov_p2[c]

age_summary = {}
for age_name, cols in AGE_COL_GROUPS.items():
    p1_sum = sum(prov_p1[c] for c in cols)
    p2_sum = sum(prov_p2[c] for c in cols)
    tot_sum = sum(prov_tot[c] for c in cols)
    age_summary[age_name] = {
        "p1": p1_sum,
        "p2": p2_sum,
        "total": tot_sum
    }

report = {
    "grand_total": {
        "p1": sum(prov_p1.values()),
        "p2": sum(prov_p2.values()),
        "total": sum(prov_tot.values())
    },
    "age_summary": age_summary,
    "col_p1": prov_p1,
    "col_p2": prov_p2,
    "col_tot": prov_tot
}

with open("scratch/1m7_province_summary.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("Exported 1m7_province_summary.json successfully!")
print("Grand Total:", report["grand_total"])
for age, s in age_summary.items():
    print(f"  {age:10s}: P1={s['p1']:6.0f}, P2={s['p2']:5.0f}, Total={s['total']:6.0f}")
