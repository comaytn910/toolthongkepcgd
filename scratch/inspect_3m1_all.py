import sys
import glob
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

files = sorted(glob.glob('data/*.xlsx'))

# Khảo sát chi tiết 3.M1
print("=== CHI TIẾT 3.M1 TRÊN 92 XÃ ===")
sum_c = 0 # Dân số 15-25
sum_d = 0 # Biết chữ M1
sum_f = 0 # Biết chữ M2
communes_m1 = 0
communes_m2 = 0

rows_detail = []

for f in files:
    commune_name = f.split('\\')[-1].split(' - ')[0]
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb['3.M1']
    
    total_row_idx = None
    total_vals = None
    
    # Tìm dòng tổng (@ Cộng)
    for r in range(1, ws.max_row + 1):
        v0 = str(ws.cell(r, 1).value or '').strip()
        v1 = str(ws.cell(r, 2).value or '').strip().lower()
        if v0 == '@' or 'cộng' in v1:
            total_row_idx = r
            total_vals = [ws.cell(r, c).value for c in range(1, 10)]
            break
            
    # Kiểm tra đạt chuẩn (các dòng dưới dòng tổng)
    dat_m1 = False
    dat_m2 = False
    if total_row_idx:
        for r in range(total_row_idx + 1, min(total_row_idx + 5, ws.max_row + 1)):
            label = str(ws.cell(r, 1).value or '') + str(ws.cell(r, 2).value or '')
            # Cột D hoặc Cột C/E thường có số 1 hoặc dấu đạt
            val_d = ws.cell(r, 4).value
            val_c = ws.cell(r, 3).value
            if 'mức 1' in label.lower() or 'mức độ 1' in label.lower():
                if val_d == 1 or val_c == 1 or str(val_d).strip() in ['1', 'x', 'X', '1.0']:
                    dat_m1 = True
            elif 'mức 2' in label.lower() or 'mức độ 2' in label.lower():
                if val_d == 1 or val_c == 1 or str(val_d).strip() in ['1', 'x', 'X', '1.0']:
                    dat_m2 = True
                    
    c = float(total_vals[2] or 0) if total_vals else 0
    d = float(total_vals[3] or 0) if total_vals else 0
    f_val = float(total_vals[5] or 0) if total_vals else 0
    
    sum_c += c
    sum_d += d
    sum_f += f_val
    if dat_m1: communes_m1 += 1
    if dat_m2: communes_m2 += 1
    
    rows_detail.append((commune_name, total_row_idx, c, d, f_val, dat_m1, dat_m2))

rate_d = (sum_d / sum_c * 100) if sum_c > 0 else 0
rate_f = (sum_f / sum_c * 100) if sum_c > 0 else 0

print(f"Tổng số dân số 15-25 tuổi toàn tỉnh (Cột C): {sum_c:,.0f}")
print(f"Số người biết chữ Mức 1 toàn tỉnh (Cột D): {sum_d:,.0f} (Tỷ lệ: {rate_d:.2f}%)")
print(f"Số người biết chữ Mức 2 toàn tỉnh (Cột F): {sum_f:,.0f} (Tỷ lệ: {rate_f:.2f}%)")
print(f"Số xã đạt chuẩn XMC Mức 1: {communes_m1}/92")
print(f"Số xã đạt chuẩn XMC Mức 2: {communes_m2}/92")

# In 10 xã đầu
print("\nMẫu 10 xã đầu:")
for r in rows_detail[:10]:
    print(r)
