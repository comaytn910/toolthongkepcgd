import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.excel_reader import scan_data_directory
from processors.m3_m1 import process_3_m1_single_file
from processors.m3_m2 import process_3_m2_single_file
from processors.m3_m3 import process_3_m3_single_file
from processors.m3_mu_chu import process_3_mu_chu_single_file

scan_res = scan_data_directory(os.path.join(BASE_DIR, 'data'))
valid_files = [f for f in scan_res['files'] if f['is_valid']]

print("=== TRUY VẾT CHI TIẾT TỪNG XÃ CHÊNH LỆCH BIỂU XÓA MÙ CHỮ ===")

diff_3m4 = []
diff_3m5 = []
diff_3m6 = []
diff_3m7 = []
diff_3m8 = []
diff_3m9 = []

for f in valid_files:
    c_name = f['commune_name']
    fp = f['filepath']
    
    # 3.M1, 3.M2, 3.M3
    r_3m1 = process_3_m1_single_file(fp, c_name)
    r_3m2 = process_3_m2_single_file(fp, c_name)
    r_3m3 = process_3_m3_single_file(fp, c_name)
    
    # 3.M4 .. 3.M9
    r_3m4 = process_3_mu_chu_single_file(fp, c_name, '3.M4')
    r_3m5 = process_3_mu_chu_single_file(fp, c_name, '3.M5')
    r_3m6 = process_3_mu_chu_single_file(fp, c_name, '3.M6')
    r_3m7 = process_3_mu_chu_single_file(fp, c_name, '3.M7')
    r_3m8 = process_3_mu_chu_single_file(fp, c_name, '3.M8')
    r_3m9 = process_3_mu_chu_single_file(fp, c_name, '3.M9')
    
    # 1. Đối soát 3.M4 (Mù chữ M1 15-25t) vs 3.M1
    # Logic: Dân số (C) = Biết chữ M1 (3.M1 Cột D) + Mù chữ M1 (3.M4 Cột G)
    dan_so_1 = r_3m1['c_dan_so']
    biet_chu_m1_1 = r_3m1['d_biet_chu_m1']
    mu_chu_m1_1 = r_3m4['g_mu_chu']
    tong_m1_1 = biet_chu_m1_1 + mu_chu_m1_1
    if abs(tong_m1_1 - dan_so_1) > 0.01:
        diff_3m4.append({
            'commune': c_name,
            'dan_so': dan_so_1,
            'biet_chu': biet_chu_m1_1,
            'mu_chu': mu_chu_m1_1,
            'tong': tong_m1_1,
            'diff': round(tong_m1_1 - dan_so_1, 2)
        })
        
    # 2. Đối soát 3.M5 (Mù chữ M1 15-35t) vs 3.M2
    dan_so_2 = r_3m2['c_dan_so']
    biet_chu_m1_2 = r_3m2['d_biet_chu_m1']
    mu_chu_m1_2 = r_3m5['g_mu_chu']
    tong_m1_2 = biet_chu_m1_2 + mu_chu_m1_2
    if abs(tong_m1_2 - dan_so_2) > 0.01:
        diff_3m5.append({
            'commune': c_name,
            'dan_so': dan_so_2,
            'biet_chu': biet_chu_m1_2,
            'mu_chu': mu_chu_m1_2,
            'tong': tong_m1_2,
            'diff': round(tong_m1_2 - dan_so_2, 2)
        })
        
    # 3. Đối soát 3.M6 (Mù chữ M1 15-60t) vs 3.M3
    dan_so_3 = r_3m3['c_dan_so']
    biet_chu_m1_3 = r_3m3['d_biet_chu_m1']
    mu_chu_m1_3 = r_3m6['g_mu_chu']
    tong_m1_3 = biet_chu_m1_3 + mu_chu_m1_3
    if abs(tong_m1_3 - dan_so_3) > 0.01:
        diff_3m6.append({
            'commune': c_name,
            'dan_so': dan_so_3,
            'biet_chu': biet_chu_m1_3,
            'mu_chu': mu_chu_m1_3,
            'tong': tong_m1_3,
            'diff': round(tong_m1_3 - dan_so_3, 2)
        })

print("\n--- 1. CHI TIẾT CÁC XÃ LỆCH GIỮA 3.M4 VÀ 3.M1 (15-25 TUỔI) ---")
print(f"Tổng số xã lệch: {len(diff_3m4)}")
for x in diff_3m4:
    print(f"  + {x['commune']}: Dân số (3.M1!C)={x['dan_so']}, Biết chữ M1 (3.M1!D)={x['biet_chu']}, Mù chữ (3.M4!G)={x['mu_chu']} -> Tổng={x['tong']}, Lệch={x['diff']:+g} người")

print("\n--- 2. CHI TIẾT CÁC XÃ LỆCH GIỮA 3.M5 VÀ 3.M2 (15-35 TUỔI) ---")
print(f"Tổng số xã lệch: {len(diff_3m5)}")
for x in diff_3m5:
    print(f"  + {x['commune']}: Dân số (3.M2!C)={x['dan_so']}, Biết chữ M1 (3.M2!D)={x['biet_chu']}, Mù chữ (3.M5!G)={x['mu_chu']} -> Tổng={x['tong']}, Lệch={x['diff']:+g} người")

print("\n--- 3. CHI TIẾT CÁC XÃ LỆCH GIỮA 3.M6 VÀ 3.M3 (15-60 TUỔI) ---")
print(f"Tổng số xã lệch: {len(diff_3m6)}")
for x in diff_3m6:
    print(f"  + {x['commune']}: Dân số (3.M3!C)={x['dan_so']}, Biết chữ M1 (3.M3!D)={x['biet_chu']}, Mù chữ (3.M6!G)={x['mu_chu']} -> Tổng={x['tong']}, Lệch={x['diff']:+g} người")
