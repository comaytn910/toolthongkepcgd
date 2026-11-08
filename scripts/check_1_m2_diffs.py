import os
import sys

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.fast_excel import fast_read_sheet
from core.excel_reader import scan_data_directory

data_dir = os.path.join(BASE_DIR, 'data')
scan = scan_data_directory(data_dir)
valid_files = [f for f in scan['files'] if f['is_valid']]

print("=== SO SÁNH GIỮA HÀNG 8 VÀ CỘNG CHI TIẾT TRƯỜNG DÒNG 9..29 Ở 92 XÃ ===")
diffs = []
for f in valid_files:
    fpath = f['filepath']
    cname = f['commune_name']
    cells = fast_read_sheet(fpath, '1.M2')
    if not cells: continue
    
    i8_val = cells.get('I8', {}).get('val', 0) or 0
    calc_val = 0
    school_rows = []
    for r in range(9, 30):
        b = cells.get(f'B{r}', {}).get('val')
        i = cells.get(f'I{r}', {}).get('val', 0) or 0
        if b and str(b).strip():
            school_rows.append((r, str(b).strip(), i))
            if isinstance(i, (int, float)):
                calc_val += i
                
    if abs(i8_val - calc_val) > 0.001:
        diffs.append({
            'commune': cname,
            'file': f['filename'],
            'row8_formula': cells.get('I8', {}).get('formula'),
            'row8_val': i8_val,
            'calc_val': calc_val,
            'schools': school_rows
        })

print(f"Số xã có chênh lệch giữa Hàng 8 và tổng chi tiết: {len(diffs)} / 92 xã")
for d in diffs[:10]:
    print(f"\n[{d['commune']}] Row 8 = {d['row8_val']} (F: {d['row8_formula']}) | Calc = {d['calc_val']}")
    for s in d['schools']:
        print(f"  Row {s[0]}: {s[1]} -> GV: {s[2]}")
