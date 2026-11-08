import os
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from processors.m3_m1 import process_3_m1_single_file
from processors.m3_m2 import process_3_m2_single_file
from processors.m3_m3 import process_3_m3_single_file
from processors.m3_mu_chu import process_3_mu_chu_single_file

DATA_92_DIR = r"D:\DATA PCGD\1810"
files_92 = sorted(glob.glob(os.path.join(DATA_92_DIR, "*.xlsx")))
print(f"Tổng số files trong {DATA_92_DIR}: {len(files_92)}")

# 1. Khảo sát 3.M1 vs 3.M4 (15-25t)
print("\n==========================================================================")
print("1. TRUY VẾT CHÊNH LỆCH 3.M1 VS 3.M4 (ĐỘ TUỔI 15-25) TRÊN ĐỦ 92 XÃ")
print("==========================================================================")
diff_1 = []
tot_c1 = 0
tot_d1 = 0
tot_g4 = 0

for fp in files_92:
    c_name = os.path.basename(fp).split(' - ')[0]
    r1 = process_3_m1_single_file(fp, c_name)
    r4 = process_3_mu_chu_single_file(fp, c_name, '3.M4')
    if not r1 or not r4:
        continue
        
    pop = r1['c_dan_so']
    bc = r1['d_biet_chu_m1']
    mc = r4['g_mu_chu']
    tot_c1 += pop
    tot_d1 += bc
    tot_g4 += mc
    
    diff = (bc + mc) - pop
    if abs(diff) > 0.001:
        diff_1.append({
            'commune': c_name,
            'pop': pop,
            'bc': bc,
            'mc': mc,
            'sum': bc + mc,
            'diff': diff,
            'file': os.path.basename(fp)
        })

print(f"Tổng 92 xã: Dân số (3.M1!C) = {tot_c1:,.0f} | Biết chữ M1 (3.M1!D) = {tot_d1:,.0f} | Mù chữ M1 (3.M4!G) = {tot_g4:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_d1 + tot_g4:,.0f} vs Dân số = {tot_c1:,.0f} (Chênh lệch: {tot_d1 + tot_g4 - tot_c1:+.0f} người)")
print(f"Số xã gây chênh lệch: {len(diff_1)}")
for x in diff_1:
    print(f"  * Xã: {x['commune']:<30} | 3.M1!C={x['pop']} | 3.M1!D={x['bc']} | 3.M4!G={x['mc']} | Tổng BC+MC={x['sum']} | Lệch={x['diff']:+g} người")

# 2. Khảo sát 3.M2 vs 3.M5 (15-35t)
print("\n==========================================================================")
print("2. TRUY VẾT CHÊNH LỆCH 3.M2 VS 3.M5 (ĐỘ TUỔI 15-35) TRÊN ĐỦ 92 XÃ")
print("==========================================================================")
diff_2 = []
tot_c2 = 0
tot_d2 = 0
tot_g5 = 0

for fp in files_92:
    c_name = os.path.basename(fp).split(' - ')[0]
    r2 = process_3_m2_single_file(fp, c_name)
    r5 = process_3_mu_chu_single_file(fp, c_name, '3.M5')
    if not r2 or not r5:
        continue
        
    pop = r2['c_dan_so']
    bc = r2['d_biet_chu_m1']
    mc = r5['g_mu_chu']
    tot_c2 += pop
    tot_d2 += bc
    tot_g5 += mc
    
    diff = (bc + mc) - pop
    if abs(diff) > 0.001:
        diff_2.append({
            'commune': c_name,
            'pop': pop,
            'bc': bc,
            'mc': mc,
            'sum': bc + mc,
            'diff': diff,
            'file': os.path.basename(fp)
        })

print(f"Tổng 92 xã: Dân số (3.M2!C) = {tot_c2:,.0f} | Biết chữ M1 (3.M2!D) = {tot_d2:,.0f} | Mù chữ M1 (3.M5!G) = {tot_g5:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_d2 + tot_g5:,.0f} vs Dân số = {tot_c2:,.0f} (Chênh lệch: {tot_d2 + tot_g5 - tot_c2:+.0f} người)")
print(f"Số xã gây chênh lệch: {len(diff_2)}")
for x in diff_2:
    print(f"  * Xã: {x['commune']:<30} | 3.M2!C={x['pop']} | 3.M2!D={x['bc']} | 3.M5!G={x['mc']} | Tổng BC+MC={x['sum']} | Lệch={x['diff']:+g} người")

# 3. Khảo sát 3.M3 vs 3.M6 (15-60t)
print("\n==========================================================================")
print("3. TRUY VẾT CHÊNH LỆCH 3.M3 VS 3.M6 (ĐỘ TUỔI 15-60) TRÊN ĐỦ 92 XÃ")
print("==========================================================================")
diff_3 = []
tot_c3 = 0
tot_d3 = 0
tot_g6 = 0

for fp in files_92:
    c_name = os.path.basename(fp).split(' - ')[0]
    r3 = process_3_m3_single_file(fp, c_name)
    r6 = process_3_mu_chu_single_file(fp, c_name, '3.M6')
    if not r3 or not r6:
        continue
        
    pop = r3['c_dan_so']
    bc = r3['d_biet_chu_m1']
    mc = r6['g_mu_chu']
    tot_c3 += pop
    tot_d3 += bc
    tot_g6 += mc
    
    diff = (bc + mc) - pop
    if abs(diff) > 0.001:
        diff_3.append({
            'commune': c_name,
            'pop': pop,
            'bc': bc,
            'mc': mc,
            'sum': bc + mc,
            'diff': diff,
            'file': os.path.basename(fp)
        })

print(f"Tổng 92 xã: Dân số (3.M3!C) = {tot_c3:,.0f} | Biết chữ M1 (3.M3!D) = {tot_d3:,.0f} | Mù chữ M1 (3.M6!G) = {tot_g6:,.0f}")
print(f"Tổng Biết chữ + Mù chữ = {tot_d3 + tot_g6:,.0f} vs Dân số = {tot_c3:,.0f} (Chênh lệch: {tot_d3 + tot_g6 - tot_c3:+.0f} người)")
print(f"Số xã gây chênh lệch: {len(diff_3)}")
for x in diff_3:
    print(f"  * Xã: {x['commune']:<30} | 3.M3!C={x['pop']} | 3.M3!D={x['bc']} | 3.M6!G={x['mc']} | Tổng BC+MC={x['sum']} | Lệch={x['diff']:+g} người")
