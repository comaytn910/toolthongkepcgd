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
from processors.m1_m1 import process_1_m1

data_dir = os.path.join(BASE_DIR, 'data')
scan = scan_data_directory(data_dir)
valid_files = [f for f in scan['files'] if f['is_valid']]

print("=== TÍNH TOÁN TOÀN DIỆN BIỂU 1.M4 TRÊN 92 XÃ ===")

# Đọc kết quả chuẩn từ 1.M1 để đối soát liên biểu
r_1_m1 = process_1_m1(valid_files)
m1_crit = r_1_m1['criteria_results']

cols_sum = ['C', 'D', 'E', 'F', 'H', 'J', 'L', 'N', 'O', 'P']
sums = {c: 0.0 for c in cols_sum}

guarantee_staff = 0
guarantee_fac = 0
level_3_count = 0

for f in valid_files:
    cells = fast_read_sheet(f['filepath'], '1.M4')
    for c in cols_sum:
        v = cells.get(f"{c}8", {}).get('val') or 0
        try: sums[c] += float(v)
        except: pass
        
    r_val = str(cells.get('R8', {}).get('val') or '').lower()
    s_val = str(cells.get('S8', {}).get('val') or '').lower()
    t_val = str(cells.get('T8', {}).get('val') or '').lower()
    
    if 'đảm bảo' in r_val or 'dam bao' in r_val or 'đạt' in r_val:
        guarantee_staff += 1
    if 'đảm bảo' in s_val or 'dam bao' in s_val or 'đạt' in s_val:
        guarantee_fac += 1
    if '3' in t_val:
        level_3_count += 1

print("\n--- KẾT QUẢ SỐ LIỆU TOÀN TỈNH 1.M4 ---")
print(f"1. Tổng số trường tiểu học (C): {sums['C']:,.0f}")
print(f"2. Tổng số học sinh PPC (D): {sums['D']:,.0f}")
print(f"3. Trẻ khuyết tật (E, N): {sums['E']:,.0f}")
print(f"4. Trẻ 6t vào lớp 1 (F): {sums['F']:,.0f} | Tỉ lệ G = {m1_crit[58]['rate']:.2f}% (Từ 1.M1)")
print(f"5. Trẻ 11t HTCTTH (H): {sums['H']:,.0f} | Tỉ lệ I = {m1_crit[61]['rate']:.2f}% (Từ 1.M1)")
print(f"6. Trẻ đến 14t HTCTTH (J): {sums['J']:,.0f} | Tỉ lệ K = {m1_crit[62]['rate']:.2f}% (Từ 1.M1)")
print(f"7. Trẻ 14t HTCTTH (L): {sums['L']:,.0f} | Tỉ lệ M = {m1_crit[62]['rate']:.2f}% (Từ 1.M1)")
print(f"8. Trẻ KT có khả năng HT (O): {sums['O']:,.0f} | Tiếp cận GD (P): {sums['P']:,.0f}")
rate_q = round(sums['P'] / sums['O'] * 100, 2) if sums['O'] > 0 else 0.0
print(f"   Tỉ lệ trẻ KT tiếp cận GD (Q): {rate_q:.2f}%")
print(f"9. Điều kiện Đội ngũ đảm bảo (R): {guarantee_staff} / {len(valid_files)} xã")
print(f"10. Điều kiện CSVC đảm bảo (S): {guarantee_fac} / {len(valid_files)} xã")
print(f"11. Đạt chuẩn Mức độ 3 (T): {level_3_count} / {len(valid_files)} xã")
