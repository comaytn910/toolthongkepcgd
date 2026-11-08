import os
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from processors.m3_m1 import process_3_m1_single_file
from processors.m3_mu_chu import process_3_mu_chu_single_file

files = sorted(glob.glob('data/*.xlsx'))
print(f"Tổng số files trong data/: {len(files)}")

none_3m1 = []
none_3m4 = []
data_3m1 = []
data_3m4 = []

for f in files:
    c_name = os.path.basename(f).split(' - ')[0]
    r1 = process_3_m1_single_file(f, c_name)
    r4 = process_3_mu_chu_single_file(f, c_name, '3.M4')
    if r1 is None:
        none_3m1.append(c_name)
    else:
        data_3m1.append(r1)
        
    if r4 is None:
        none_3m4.append(c_name)
    else:
        data_3m4.append(r4)

print(f"Số xã đọc được 3.M1: {len(data_3m1)} | Bị None: {len(none_3m1)} -> {none_3m1}")
print(f"Số xã đọc được 3.M4: {len(data_3m4)} | Bị None: {len(none_3m4)} -> {none_3m4}")

tot_c_3m1 = sum(x['c_dan_so'] for x in data_3m1)
tot_d_3m1 = sum(x['d_biet_chu_m1'] for x in data_3m1)
tot_c_3m4 = sum(x['c_dan_so'] for x in data_3m4)
tot_g_3m4 = sum(x['g_mu_chu'] for x in data_3m4)

print(f"\n3.M1: Tổng C (Dân số) = {tot_c_3m1:,.0f} | Tổng D (Biết chữ M1) = {tot_d_3m1:,.0f}")
print(f"3.M4: Tổng C (Dân số) = {tot_c_3m4:,.0f} | Tổng G (Mù chữ M1) = {tot_g_3m4:,.0f}")
print(f"Biết chữ (3.M1) + Mù chữ (3.M4) = {tot_d_3m1 + tot_g_3m4:,.0f}")
print(f"Chênh lệch: {tot_d_3m1 + tot_g_3m4 - tot_c_3m1:+.0f} người")
