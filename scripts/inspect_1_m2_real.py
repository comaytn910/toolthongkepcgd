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

sample_files = sorted(glob.glob(os.path.join(data_dir, '*.xlsx')))[:5]

print("=== KIỂM TRA DỮ LIỆU THỰC TẾ SHEET 1.M2 Ở 5 XÃ ===")
for f in sample_files:
    fname = os.path.basename(f)
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        if '1.M2' not in wb.sheetnames:
            print(f"{fname}: KHÔNG CÓ sheet 1.M2!")
            continue
        sheet = wb['1.M2']
        # Kiểm tra Hàng 8 (Cộng toàn xã) và Hàng 30
        row8 = [sheet.cell(8, c).value for c in range(1, 34)]
        row30 = [sheet.cell(30, c).value for c in range(1, 34)]
        
        # Xem có những hàng trường nào được nhập
        schools = []
        for r in range(9, 30):
            b_val = sheet.cell(r, 2).value
            c_val = sheet.cell(r, 3).value
            i_val = sheet.cell(r, 9).value # Tổng GV
            if b_val or i_val:
                schools.append((r, b_val, i_val))
                
        # Xem tiêu chí hàng 32-35
        crit = {}
        for r in range(32, 36):
            crit[r] = (sheet.cell(r, 1).value, sheet.cell(r, 2).value, sheet.cell(r, 3).value, sheet.cell(r, 4).value)
            
        print(f"\n[{fname}]")
        print(f"  Hàng 8: STT={row8[0]}, B={row8[1]}, E(Lớp)={row8[4]}, I(Tổng GV)={row8[8]}, N(TL GV/Lớp)={row8[13]}")
        print(f"  Hàng 30: STT={row30[0]}, B={row30[1]}, E(Lớp)={row30[4]}, I(Tổng GV)={row30[8]}, N(TL GV/Lớp)={row30[13]}")
        print(f"  Số trường học nhập liệu: {len(schools)} trường:")
        for s in schools:
            print(f"    Row {s[0]}: Trường='{s[1]}', Tổng GV={s[2]}")
        print(f"  Tiêu chí hàng 32-35: {crit}")
        wb.close()
    except Exception as e:
        print(f"Lỗi đọc {fname}: {e}")
