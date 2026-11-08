import sys
import glob
import time
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

files = sorted(glob.glob('data/*.xlsx'))
print(f"Kiểm thử thuật toán đọc thông minh 3.M4..3.M9 trên {len(files)} xã...")

sheet_pairs = [
    ('3.M4', 48, 9),  # (sheet, default_row_sum, start_row_xom)
    ('3.M5', 47, 8),
    ('3.M6', 47, 8),
    ('3.M7', 48, 9),
    ('3.M8', 47, 8),
    ('3.M9', 47, 8),
]

def parse_num(v):
    if v is None: return 0.0
    try:
        if isinstance(v, str): v = v.replace(',', '').strip()
        return float(v)
    except: return 0.0

for sname, def_sum_r, start_r in sheet_pairs:
    t0 = time.time()
    found_count = 0
    sum_c = 0.0
    sum_g_mu_chu = 0.0
    
    for f in files:
        c_name = f.split('\\')[-1].split(' - ')[0]
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        if sname not in wb.sheetnames:
            wb.close()
            continue
        ws = wb[sname]
        
        # Đọc tất cả các dòng từ start_r đến 55
        rows = list(ws.iter_rows(min_row=1, max_row=60, values_only=True))
        wb.close()
        
        total_row = None
        # Cách 1: Tìm theo từ khóa @ hoặc tổng / cộng
        for r_idx in range(len(rows) - 1, start_r - 2, -1):
            row = rows[r_idx]
            v0 = str(row[0] or '').strip().lower()
            v1 = str(row[1] or '').strip().lower()
            if v0 == '@' or any(k in v1 for k in ['tổng', 'cộng']):
                if parse_num(row[2]) > 0:
                    total_row = row
                    break
                    
        # Cách 2: Nếu chưa tìm thấy, lấy dòng mặc định (def_sum_r - 1) nếu có số > 0
        if not total_row and len(rows) >= def_sum_r:
            candidate = rows[def_sum_r - 1]
            if parse_num(candidate[2]) > 0:
                total_row = candidate
                
        # Cách 3: Nếu vẫn chưa thấy, tính tổng từ các xóm
        if not total_row:
            # Tính tổng các dòng xóm (các dòng có tên xóm ở cột B và số ở cột C)
            sum_row_calc = [0.0] * 15
            has_xom = False
            for r_idx in range(start_r - 1, min(len(rows), def_sum_r - 1)):
                row = rows[r_idx]
                if row and row[1] and parse_num(row[2]) > 0:
                    has_xom = True
                    for c_idx in range(2, min(15, len(row))):
                        sum_row_calc[c_idx] += parse_num(row[c_idx])
            if has_xom and sum_row_calc[2] > 0:
                total_row = sum_row_calc
                
        if total_row:
            found_count += 1
            c_val = parse_num(total_row[2])
            g_val = parse_num(total_row[6]) if len(total_row) > 6 else 0.0
            sum_c += c_val
            sum_g_mu_chu += g_val
            
    rate_g = (sum_g_mu_chu / sum_c * 100) if sum_c > 0 else 0.0
    elapsed = time.time() - t0
    print(f"{sname:4s}: Đã đọc thành công {found_count}/92 xã | Dân số C={sum_c:12,.0f} | Mù chữ G={sum_g_mu_chu:8,.0f} ({rate_g:5.2f}%) [{elapsed:.1f}s]")
