import os
import sys
import glob
import openpyxl

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_dir = os.path.join(BASE_DIR, 'data')
files = sorted(glob.glob(os.path.join(data_dir, '*.xlsx')))

print(f"=== KHẢO SÁT 1.M2 TRÊN TOÀN BỘ {len(files)} XÃ ===")
total_schools = 0
total_gv = 0
total_cbql_ht = 0
total_cbql_pht = 0
missing_1_m2 = []

for f in files:
    fname = os.path.basename(f)
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        if '1.M2' not in wb.sheetnames:
            missing_1_m2.append(fname)
            wb.close()
            continue
        ws = wb['1.M2']
        
        # Đếm các trường thực tế có tên trường ở B9:B29
        commune_schools = 0
        for r in range(9, 30):
            b_val = ws.cell(r, 2).value
            i_val = ws.cell(r, 9).value # Tổng GV
            if b_val and str(b_val).strip() != "":
                commune_schools += 1
                
        # Lấy giá trị tổng cộng của xã từ Hàng 8 (hoặc tính lại từ các trường)
        # Kiểm tra Hàng 8
        gv_row8 = ws.cell(8, 9).value or 0
        ht_row8 = ws.cell(8, 7).value or 0
        pht_row8 = ws.cell(8, 8).value or 0
        
        total_schools += commune_schools
        total_gv += float(gv_row8) if isinstance(gv_row8, (int, float)) else 0
        total_cbql_ht += float(ht_row8) if isinstance(ht_row8, (int, float)) else 0
        total_cbql_pht += float(pht_row8) if isinstance(pht_row8, (int, float)) else 0
        wb.close()
    except Exception as e:
        print(f"Lỗi đọc {fname}: {e}")

print(f"\nKết quả kiểm kê 1.M2:")
print(f"- Số xã thiếu biểu 1.M2: {len(missing_1_m2)}")
print(f"- Tổng số trường tiểu học ghi nhận: {total_schools}")
print(f"- Tổng số Hiệu trưởng (G8): {total_cbql_ht}")
print(f"- Tổng số Phó Hiệu trưởng (H8): {total_cbql_pht}")
print(f"- Tổng số Giáo viên Tiểu học toàn tỉnh (I8): {total_gv:,.0f}")
