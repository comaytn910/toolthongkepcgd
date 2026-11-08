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

print("==========================================================================")
print("TRUY TÌM CHÍNH XÁC CÁC XÃ GÂY RA CHÊNH LỆCH ĐỐI SOÁT LIÊN BIỂU XÓA MÙ CHỮ")
print("==========================================================================")

# 1. 15-25 tuổi: 3.M1 vs 3.M4
print("\n--- 1. CẶP BIỂU 3.M1 VÀ 3.M4 (ĐỘ TUỔI 15-25) ---")
print("Quy tắc kiểm tra: Dân số (3.M1!C hoặc 3.M4!C) = Biết chữ M1 (3.M1!D) + Mù chữ M1 (3.M4!G)")
tot_pop1 = 0
tot_bc1 = 0
tot_mc1 = 0
tot_pop4 = 0

diffs_1 = []

for f in valid_files:
    c = f['commune_name']
    fp = f['filepath']
    d1 = process_3_m1_single_file(fp, c)
    d4 = process_3_mu_chu_single_file(fp, c, '3.M4')
    
    tot_pop1 += d1['c_dan_so']
    tot_bc1 += d1['d_biet_chu_m1']
    tot_mc1 += d4['g_mu_chu']
    tot_pop4 += d4['c_dan_so']
    
    # Kiểm tra
    sum_1 = d1['d_biet_chu_m1'] + d4['g_mu_chu']
    diff_sum = sum_1 - d1['c_dan_so']
    diff_pop = d4['c_dan_so'] - d1['c_dan_so']
    
    if diff_sum != 0 or diff_pop != 0:
        diffs_1.append({
            'commune': c,
            '3m1_pop': d1['c_dan_so'],
            '3m4_pop': d4['c_dan_so'],
            'biet_chu': d1['d_biet_chu_m1'],
            'mu_chu': d4['g_mu_chu'],
            'sum': sum_1,
            'diff_sum': diff_sum,
            'diff_pop': diff_pop,
            'raw_3m1_row': d1.get('total_row_idx'),
            'raw_3m4_row': d4.get('total_row_idx')
        })

print(f"Tổng hợp toàn tỉnh: Dân số 3.M1={tot_pop1:,.0f} | Dân số 3.M4={tot_pop4:,.0f} | Biết chữ={tot_bc1:,.0f} | Mù chữ={tot_mc1:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_bc1 + tot_mc1:,.0f} vs Dân số 3.M1 = {tot_pop1:,.0f} (Chênh lệch: {tot_bc1 + tot_mc1 - tot_pop1:+.0f} người)")
print(f"Số xã phát hiện chênh lệch: {len(diffs_1)}")
for x in diffs_1:
    print(f"  * Xã: {x['commune']:<30} | 3.M1!C={x['3m1_pop']} | 3.M4!C={x['3m4_pop']} | 3.M1!D(BC)={x['biet_chu']} | 3.M4!G(MC)={x['mu_chu']} | Lệch Tổng={x['diff_sum']:+g} | Lệch Dân số={x['diff_pop']:+g}")


# 2. 15-35 tuổi: 3.M2 vs 3.M5
print("\n--- 2. CẶP BIỂU 3.M2 VÀ 3.M5 (ĐỘ TUỔI 15-35) ---")
tot_pop2 = 0
tot_bc2 = 0
tot_mc5 = 0
tot_pop5 = 0
diffs_2 = []

for f in valid_files:
    c = f['commune_name']
    fp = f['filepath']
    d2 = process_3_m2_single_file(fp, c)
    d5 = process_3_mu_chu_single_file(fp, c, '3.M5')
    
    tot_pop2 += d2['c_dan_so']
    tot_bc2 += d2['d_biet_chu_m1']
    tot_mc5 += d5['g_mu_chu']
    tot_pop5 += d5['c_dan_so']
    
    sum_2 = d2['d_biet_chu_m1'] + d5['g_mu_chu']
    diff_sum = sum_2 - d2['c_dan_so']
    diff_pop = d5['c_dan_so'] - d2['c_dan_so']
    
    if diff_sum != 0 or diff_pop != 0:
        diffs_2.append({
            'commune': c,
            '3m2_pop': d2['c_dan_so'],
            '3m5_pop': d5['c_dan_so'],
            'biet_chu': d2['d_biet_chu_m1'],
            'mu_chu': d5['g_mu_chu'],
            'sum': sum_2,
            'diff_sum': diff_sum,
            'diff_pop': diff_pop
        })

print(f"Tổng hợp toàn tỉnh: Dân số 3.M2={tot_pop2:,.0f} | Dân số 3.M5={tot_pop5:,.0f} | Biết chữ={tot_bc2:,.0f} | Mù chữ={tot_mc5:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_bc2 + tot_mc5:,.0f} vs Dân số 3.M2 = {tot_pop2:,.0f} (Chênh lệch: {tot_bc2 + tot_mc5 - tot_pop2:+.0f} người)")
print(f"Số xã phát hiện chênh lệch: {len(diffs_2)}")
for x in diffs_2:
    print(f"  * Xã: {x['commune']:<30} | 3.M2!C={x['3m2_pop']} | 3.M5!C={x['3m5_pop']} | 3.M2!D(BC)={x['biet_chu']} | 3.M5!G(MC)={x['mu_chu']} | Lệch Tổng={x['diff_sum']:+g} | Lệch Dân số={x['diff_pop']:+g}")


# 3. 15-60 tuổi: 3.M3 vs 3.M6
print("\n--- 3. CẶP BIỂU 3.M3 VÀ 3.M6 (ĐỘ TUỔI 15-60) ---")
tot_pop3 = 0
tot_bc3 = 0
tot_mc6 = 0
tot_pop6 = 0
diffs_3 = []

for f in valid_files:
    c = f['commune_name']
    fp = f['filepath']
    d3 = process_3_m3_single_file(fp, c)
    d6 = process_3_mu_chu_single_file(fp, c, '3.M6')
    
    tot_pop3 += d3['c_dan_so']
    tot_bc3 += d3['d_biet_chu_m1']
    tot_mc6 += d6['g_mu_chu']
    tot_pop6 += d6['c_dan_so']
    
    sum_3 = d3['d_biet_chu_m1'] + d6['g_mu_chu']
    diff_sum = sum_3 - d3['c_dan_so']
    diff_pop = d6['c_dan_so'] - d3['c_dan_so']
    
    if diff_sum != 0 or diff_pop != 0:
        diffs_3.append({
            'commune': c,
            '3m3_pop': d3['c_dan_so'],
            '3m6_pop': d6['c_dan_so'],
            'biet_chu': d3['d_biet_chu_m1'],
            'mu_chu': d6['g_mu_chu'],
            'sum': sum_3,
            'diff_sum': diff_sum,
            'diff_pop': diff_pop
        })

print(f"Tổng hợp toàn tỉnh: Dân số 3.M3={tot_pop3:,.0f} | Dân số 3.M6={tot_pop6:,.0f} | Biết chữ={tot_bc3:,.0f} | Mù chữ={tot_mc6:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_bc3 + tot_mc6:,.0f} vs Dân số 3.M3 = {tot_pop3:,.0f} (Chênh lệch: {tot_bc3 + tot_mc6 - tot_pop3:+.0f} người)")
print(f"Số xã phát hiện chênh lệch: {len(diffs_3)}")
for x in diffs_3:
    print(f"  * Xã: {x['commune']:<30} | 3.M3!C={x['3m3_pop']} | 3.M6!C={x['3m6_pop']} | 3.M3!D(BC)={x['biet_chu']} | 3.M6!G(MC)={x['mu_chu']} | Lệch Tổng={x['diff_sum']:+g} | Lệch Dân số={x['diff_pop']:+g}")
