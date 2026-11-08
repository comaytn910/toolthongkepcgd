import os
import sys
import glob

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

print(f"=== KHẢO SÁT SHEET 1.M4 TRÊN TOÀN BỘ {len(valid_files)} XÃ ===")
levels_count = {}
total_schools = 0
total_students = 0
total_kt = 0

for f in valid_files:
    fname = f['filename']
    cname = f['commune_name']
    cells = fast_read_sheet(f['filepath'], '1.M4')
    if not cells:
        print(f"Thiếu 1.M4: {cname}")
        continue
        
    c = cells.get('C8', {}).get('val') or 0
    d = cells.get('D8', {}).get('val') or 0
    e = cells.get('E8', {}).get('val') or 0
    muc_do = cells.get('T8', {}).get('val')
    
    try: total_schools += float(c)
    except: pass
    try: total_students += float(d)
    except: pass
    try: total_kt += float(e)
    except: pass
    
    m_str = str(muc_do).strip() if muc_do is not None else "None"
    levels_count[m_str] = levels_count.get(m_str, 0) + 1

print(f"\nKết quả khảo sát 1.M4 trên {len(valid_files)} xã:")
print(f"- Tổng số trường tiểu học (C8): {total_schools}")
print(f"- Tổng số học sinh PPC (D8): {total_students:,.0f}")
print(f"- Tổng số trẻ khuyết tật (E8): {total_kt:,.0f}")
print(f"- Thống kê mức độ đạt chuẩn PCGD Tiểu học (Cột T8):")
for k, v in sorted(levels_count.items()):
    print(f"    Mức độ '{k}': {v} xã")
