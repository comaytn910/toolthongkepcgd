import sys
import glob
import time
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

t0 = time.time()
files = sorted(glob.glob('data/*.xlsx'))
print(f"Bắt đầu quét nhanh 3.M3 trên {len(files)} xã...")

sum_c = 0.0
sum_d = 0.0
sum_f = 0.0
count_m1 = 0
count_m2 = 0
samples = []

for idx, f in enumerate(files, 1):
    c_name = f.split('\\')[-1].split(' - ')[0]
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb['3.M3']
    
    total_vals = None
    total_r = None
    r_idx = 0
    for row in ws.iter_rows(values_only=True):
        r_idx += 1
        if not row or len(row) < 3:
            continue
        v0 = str(row[0] or '').strip().lower()
        v1 = str(row[1] or '').strip().lower()
        if v0 == '@' or 'cộng' in v1:
            total_vals = row
            total_r = r_idx
            break
            
    c = float(total_vals[2] or 0) if total_vals and len(total_vals) > 2 else 0.0
    d = float(total_vals[3] or 0) if total_vals and len(total_vals) > 3 else 0.0
    f_val = float(total_vals[5] or 0) if total_vals and len(total_vals) > 5 else 0.0
    
    sum_c += c
    sum_d += d
    sum_f += f_val
    
    # Kiểm tra các dòng đạt chuẩn dưới dòng tổng
    dat1 = False
    dat2 = False
    for row in ws.iter_rows(min_row=total_r + 1, max_row=total_r + 5, values_only=True):
        label = (str(row[0] or '') + " " + str(row[1] or '')).lower()
        val_d = row[3] if len(row) > 3 else None
        val_c = row[2] if len(row) > 2 else None
        val_b = row[1] if len(row) > 1 else None
        target = val_d if val_d is not None else (val_c if val_c is not None else val_b)
        t_str = str(target or '').strip().lower()
        is_ok = t_str in ['1', '1.0', 'đạt', 'x', 'true', 'v'] or target == 1 or target == 1.0
        if 'mức 1' in label or 'mức độ 1' in label:
            if is_ok: dat1 = True
        elif 'mức 2' in label or 'mức độ 2' in label:
            if is_ok: dat2 = True
            
    if dat1: count_m1 += 1
    if dat2: count_m2 += 1
    
    if idx <= 5:
        samples.append((c_name, total_r, c, d, f_val, dat1, dat2))
        
    wb.close()

elapsed = time.time() - t0
rate_m1 = (sum_d / sum_c * 100) if sum_c > 0 else 0
rate_m2 = (sum_f / sum_c * 100) if sum_c > 0 else 0

print(f"Quét xong 92 xã trong {elapsed:.2f} giây!")
print(f"Tổng số dân số 15-60 tuổi toàn tỉnh (Cột C): {sum_c:,.0f}")
print(f"Số người biết chữ Mức 1 toàn tỉnh (Cột D): {sum_d:,.0f} (Tỷ lệ: {rate_m1:.2f}%)")
print(f"Số người biết chữ Mức 2 toàn tỉnh (Cột F): {sum_f:,.0f} (Tỷ lệ: {rate_m2:.2f}%)")
print(f"Số xã đạt chuẩn XMC Mức 1: {count_m1}/92")
print(f"Số xã đạt chuẩn XMC Mức 2: {count_m2}/92")
print("Mẫu 5 xã đầu:", samples)
