import os
import sys
from pathlib import Path
import json

sys.path.insert(0, str(Path(__file__).parent.parent))
from core.fast_excel import fast_read_sheet
import openpyxl

data_dir = Path("data")
files = sorted([f for f in data_dir.glob("*.xlsx") if not f.name.startswith("~$")])

ALL_COLS = [openpyxl.utils.get_column_letter(c) for c in range(3, 44)] # C đến AQ

print(f"Total files: {len(files)}")

stats = {
    "total_files": len(files),
    "has_1m7": 0,
    "missing_1m7": [],
    "row_structures": {}, # lưu các vị trí (row_tot, row_p1, row_p2)
    "communes_data": []
}

for idx, fpath in enumerate(files, 1):
    cname = fpath.stem
    cells = fast_read_sheet(str(fpath), "1.M7")
    if not cells:
        stats["missing_1m7"].append(cname)
        continue
    
    stats["has_1m7"] += 1
    
    # Tìm các dòng:
    # row_tot: dòng có text Cộng (I + II)
    # p1_header: dòng có text HỌC SINH XÃ/PHƯỜNG NGOÀI (TRONG TỈNH)
    # row_p1: dòng có text Cộng I
    # p2_header: dòng có text HỌC SINH TỈNH/THÀNH PHỐ NGOÀI
    # row_p2: dòng có text Cộng II
    row_tot = None
    p1_header = None
    row_p1 = None
    p2_header = None
    row_p2 = None
    
    for r in range(5, 150):
        a_val = str(cells.get(f"A{r}", {}).get('val') or '').strip().lower()
        b_val = str(cells.get(f"B{r}", {}).get('val') or '').strip().lower()
        
        if 'cộng (i' in a_val or 'cộng ( i' in a_val or 'cộng (i' in b_val:
            row_tot = r
        elif 'ngoài (trong tỉnh)' in b_val or ('ngoài' in b_val and 'trong tỉnh' in b_val):
            p1_header = r
        elif b_val == 'cộng i' or a_val == 'cộng i' or (b_val == 'cộng' and p2_header is None and p1_header is not None and r > p1_header):
            row_p1 = r
        elif 'ngoài đang học tại xã' in b_val and ('tỉnh' in b_val or 'thành phố' in b_val):
            p2_header = r
        elif b_val == 'cộng ii' or a_val == 'cộng ii' or (b_val == 'cộng' and p2_header is not None and r > p2_header):
            row_p2 = r
            
    # Lấy giá trị tổng từ row_tot, row_p1, row_p2 nếu có
    tot_ar = cells.get(f"AR{row_tot}", {}).get('val') if row_tot else None
    p1_ar = cells.get(f"AR{row_p1}", {}).get('val') if row_p1 else None
    p2_ar = cells.get(f"AR{row_p2}", {}).get('val') if row_p2 else None
    
    # Tính sum từ các cột C..AQ
    sum_tot_caq = sum(float(cells.get(f"{c}{row_tot}", {}).get('val') or 0) for c in ALL_COLS) if row_tot else 0.0
    sum_p1_caq = sum(float(cells.get(f"{c}{row_p1}", {}).get('val') or 0) for c in ALL_COLS) if row_p1 else 0.0
    sum_p2_caq = sum(float(cells.get(f"{c}{row_p2}", {}).get('val') or 0) for c in ALL_COLS) if row_p2 else 0.0
    
    # Đếm số trường có khai báo dữ liệu trong Phần I và Phần II
    schools_p1 = []
    if p1_header and row_p1:
        for r in range(p1_header + 1, row_p1):
            s_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
            s_sum = sum(float(cells.get(f"{c}{r}", {}).get('val') or 0) for c in ALL_COLS)
            if s_name or s_sum > 0:
                schools_p1.append({'row': r, 'name': s_name, 'sum': s_sum})
                
    schools_p2 = []
    if p2_header and row_p2:
        for r in range(p2_header + 1, row_p2):
            s_name = str(cells.get(f"B{r}", {}).get('val') or '').strip()
            s_sum = sum(float(cells.get(f"{c}{r}", {}).get('val') or 0) for c in ALL_COLS)
            if s_name or s_sum > 0:
                schools_p2.append({'row': r, 'name': s_name, 'sum': s_sum})
                
    struct_key = f"tot:{row_tot}_p1:{row_p1}_p2:{row_p2}"
    stats["row_structures"][struct_key] = stats["row_structures"].get(struct_key, 0) + 1
    
    stats["communes_data"].append({
        "commune": cname,
        "rows": {"tot": row_tot, "p1_head": p1_header, "p1": row_p1, "p2_head": p2_header, "p2": row_p2},
        "vals": {
            "p1_ar": p1_ar, "sum_p1_caq": sum_p1_caq, "schools_p1_count": len(schools_p1),
            "p2_ar": p2_ar, "sum_p2_caq": sum_p2_caq, "schools_p2_count": len(schools_p2),
            "tot_ar": tot_ar, "sum_tot_caq": sum_tot_caq
        },
        "schools_p1": schools_p1,
        "schools_p2": schools_p2
    })

with open("scratch/scan_1m7_all.json", "w", encoding="utf-8") as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)

print(f"Done scanning! Has 1.M7: {stats['has_1m7']}/{len(files)}")
print(f"Row structures distribution: {stats['row_structures']}")
