import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.excel_reader import scan_data_directory
scan_res = scan_data_directory(os.path.join(BASE_DIR, 'data'))
valid_files = [f for f in scan_res['files'] if f['is_valid']]

print("=== XÂY DỰNG BẢNG ĐÓNG GÓP BIỂU 0.M5 CỦA 92 XÃ ===")

results = []

for f in valid_files:
    c_name = f['commune_name']
    fp = f['filepath']
    try:
        wb = openpyxl.load_workbook(fp, data_only=True)
        wb_raw = openpyxl.load_workbook(fp, data_only=False)
        ws = wb['0.M5']
        ws_r = wb_raw['0.M5']
        
        # Đơn vị tính ghi trên biểu (C8)
        c8_unit = str(ws.cell(8, 3).value or '').strip()
        
        # Giá trị E8 (2026), F8 (2027), G8 (2028), H8 (2029), I8 (2030), D8 (Tổng số nguồn)
        d8_val = ws.cell(8, 4).value
        e8_val = ws.cell(8, 5).value
        f8_val = ws.cell(8, 6).value
        g8_val = ws.cell(8, 7).value
        h8_val = ws.cell(8, 8).value
        i8_val = ws.cell(8, 9).value
        
        e8_formula = ws_r.cell(8, 5).value
        d8_formula = ws_r.cell(8, 4).value
        
        def to_num(x):
            if isinstance(x, (int, float)): return float(x)
            try: return float(str(x).replace(',', ''))
            except: return 0.0
            
        e8_num = to_num(e8_val)
        d8_num = to_num(d8_val)
        
        # Tính sum E..I của xã
        sum_ei = sum(to_num(ws.cell(8, c).value) for c in range(5, 10))
        
        # Hàng 9: Ngân sách thường xuyên (E9, D9)
        e9_val = to_num(ws.cell(9, 5).value)
        d9_val = to_num(ws.cell(9, 4).value)
        
        results.append({
            'commune': c_name,
            'file': os.path.basename(fp),
            'unit': c8_unit,
            'e8': e8_num,
            'd8': d8_num,
            'sum_ei': sum_ei,
            'e8_raw': str(e8_formula),
            'd8_raw': str(d8_formula),
            'e9': e9_val,
            'd9': d9_val,
            'f8': to_num(f8_val),
            'g8': to_num(g8_val),
            'h8': to_num(h8_val),
            'i8': to_num(i8_val)
        })
        wb.close()
        wb_raw.close()
    except Exception as e:
        print(f"Lỗi đọc {c_name}: {e}")

# Sắp xếp theo E8 giảm dần
sorted_res = sorted(results, key=lambda x: x['e8'], reverse=True)

tot_e8 = sum(x['e8'] for x in sorted_res)
tot_d8 = sum(x['d8'] for x in sorted_res)
tot_sum_ei = sum(x['sum_ei'] for x in sorted_res)

print(f"\nTổng số xã: {len(sorted_res)}")
print(f"Tổng E8 (Năm 2026) = {tot_e8:,.2f}")
print(f"Tổng D8 nguồn = {tot_d8:,.2f}")
print(f"Tổng D8 tính lại SUM(E:I) = {tot_sum_ei:,.2f}")

print("\n--- BẢNG CHI TIẾT ĐÓNG GÓP THEO XÃ CHO E8 VÀ D8 (TOP 25 XÃ LỚN NHẤT) ---")
print(f"{'STT':<4} | {'Tên Xã':<28} | {'Sheet/Ô':<10} | {'ĐVT':<6} | {'E8 (2026)':<20} | {'Tỷ trọng E8':<12} | {'D8 (Tổng nguồn)':<20} | {'D tính lại SUM':<20} | {'Ghi chú/Công thức E8'}")
print("-" * 155)

for idx, r in enumerate(sorted_res[:25], 1):
    pct_e8 = (r['e8'] / tot_e8 * 100) if tot_e8 > 0 else 0
    note = f"Công thức={r['e8_raw']}" if r['e8_raw'].startswith('=') else "Số tĩnh"
    if r['e8'] >= 1_000_000_000:
        note += " [BẤT THƯỜNG: Hàng nghìn tỷ (Đơn vị ĐỒNG?)]"
    elif r['e8'] >= 1_000_000:
        note += " [BẤT THƯỜNG: Hàng triệu-tỷ (Đơn vị ĐỒNG?)]"
    print(f"{idx:<4} | {r['commune']:<28} | 0.M5!E8   | {r['unit']:<6} | {r['e8']:<20,.2f} | {pct_e8:<11.4f}% | {r['d8']:<20,.2f} | {r['sum_ei']:<20,.2f} | {note}")

# Nhóm các xã chuẩn
print("\n--- TỔNG KẾT THEO NHÓM ĐỘ LỚN ---")
abnormal_group = [x for x in sorted_res if x['e8'] >= 1_000_000]
normal_group = [x for x in sorted_res if 0 < x['e8'] < 1_000_000]
zero_group = [x for x in sorted_res if x['e8'] == 0]

sum_abnormal_e8 = sum(x['e8'] for x in abnormal_group)
sum_normal_e8 = sum(x['e8'] for x in normal_group)

print(f"1. Nhóm 8 xã bất thường (>= 1,000,000):")
print(f"   - Số lượng: {len(abnormal_group)} xã")
print(f"   - Tổng E8 nhóm này: {sum_abnormal_e8:,.2f} ({sum_abnormal_e8/tot_e8*100:.4f}% toàn tỉnh)")
print(f"   - Trong đó riêng Xã Phúc Lộc: {sorted_res[0]['e8']:,.2f} ({sorted_res[0]['e8']/tot_e8*100:.4f}%)")

print(f"\n2. Nhóm 74 xã chuẩn thông thường (< 1,000,000 Tr.đg, trung bình 10-35 tỷ đồng/xã/năm):")
print(f"   - Số lượng: {len(normal_group)} xã")
print(f"   - Tổng E8 nhóm này: {sum_normal_e8:,.2f} Tr.đg")
print(f"   - Trung bình mỗi xã: {sum_normal_e8/len(normal_group):,.2f} Tr.đg (tương đương {sum_normal_e8/len(normal_group)/1000:.2f} tỷ đồng/xã/năm)")

print(f"\n3. Nhóm 10 xã không có số liệu E8 (0 hoặc trống):")
print(f"   - Số lượng: {len(zero_group)} xã")
