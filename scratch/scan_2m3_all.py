import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.fast_excel import fast_read_sheet

data_dir = Path("data")
files = sorted([f for f in data_dir.glob("*.xlsx") if not f.name.startswith("~$")])

ALL_COLS_2M3 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
COUNT_COLS_2M3 = ['C', 'D', 'E', 'H', 'I', 'K', 'M', 'N', 'O']
RATE_COLS_2M3 = ['J', 'L', 'P']
TEXT_COLS_2M3 = ['F', 'G', 'Q', 'R', 'S']

stats = {
    "total_files": len(files),
    "has_2m3": 0,
    "missing_2m3": [],
    "total_row_dist": {},
    "communes": []
}

for idx, fpath in enumerate(files, 1):
    cname = fpath.stem
    cells = fast_read_sheet(str(fpath), "2.M3")
    if not cells:
        stats["missing_2m3"].append(cname)
        continue
    
    stats["has_2m3"] += 1
    
    # Tìm dòng Tổng của xã (tìm từ dòng 7 đến 150)
    tot_row = None
    for r in range(7, 150):
        b_val = str(cells.get(f"B{r}", {}).get("val") or "").strip().lower()
        a_val = str(cells.get(f"A{r}", {}).get("val") or "").strip().lower()
        if b_val in ["cộng", "tổng", "tổng cộng"] or a_val in ["cộng", "tổng", "tổng cộng"]:
            tot_row = r
            break
            
    if tot_row is None:
        tot_row = 46
        
    stats["total_row_dist"][tot_row] = stats["total_row_dist"].get(tot_row, 0) + 1
    
    row_vals = {}
    for c in ALL_COLS_2M3:
        v = cells.get(f"{c}{tot_row}", {}).get("val")
        row_vals[c] = v
        
    # Tính sum từ các dòng chi tiết (từ hàng 7 đến tot_row - 1)
    detail_sums = {c: 0.0 for c in COUNT_COLS_2M3}
    for r in range(7, tot_row):
        for c in COUNT_COLS_2M3:
            val = cells.get(f"{c}{r}", {}).get("val")
            if isinstance(val, (int, float)):
                detail_sums[c] += float(val)
            elif isinstance(val, str):
                try:
                    detail_sums[c] += float(val.strip().replace(',', ''))
                except ValueError:
                    pass
                    
    stats["communes"].append({
        "stt": idx,
        "name": cname,
        "tot_row": tot_row,
        "vals": row_vals,
        "detail_sums": detail_sums
    })

with open("scratch/scan_2m3_all.json", "w", encoding="utf-8") as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)

print(f"Done scanning 2.M3! Total: {len(files)}, Has 2.M3: {stats['has_2m3']}")
print(f"Total row distribution: {stats['total_row_dist']}")
