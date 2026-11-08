import os
import sys
import time

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from core.excel_reader import scan_data_directory
from processors.m0_m1 import process_0_m1
from processors.m0_m2 import process_0_m2
from processors.m0_m3 import process_0_m3
from processors.m0_m4 import process_0_m4
from processors.m0_m5 import process_0_m5
from processors.m0_m6 import process_0_m6
from processors.m0_m7 import process_0_m7
from processors.m1_m1 import process_1_m1
from processors.m1_m2 import process_1_m2
from processors.m1_m3 import process_1_m3
from processors.m1_m4 import process_1_m4
from processors.m1_m5 import process_1_m5
from processors.m1_m6 import process_1_m6
from processors.m1_m7 import process_1_m7
from processors.m2_m1 import process_2_m1
from processors.m2_m2 import process_2_m2
from processors.m2_m3 import process_2_m3
from processors.m2_m4 import process_2_m4
from processors.m2_m5 import process_2_m5
from processors.m2_m6 import process_2_m6
from processors.m2_m7 import process_2_m7
from processors.m2_m8 import process_2_m8
from processors.m2_m9 import process_2_m9
from processors.m2_m10 import process_2_m10
from processors.m3_m1 import process_3_m1_single_file, aggregate_3_m1
from processors.m3_m2 import process_3_m2_single_file, aggregate_3_m2
from processors.m3_m3 import process_3_m3_single_file, aggregate_3_m3
from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu

def run_all():
    data_dir = os.path.join(BASE_DIR, 'data')
    print("=== CHẠY KIỂM THỬ HỒI QUY TOÀN BỘ 33 BIỂU NGHIỆP VỤ PCGD - XMC TRÊN 92 XÃ ===", flush=True)
    t0 = time.time()
    
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, có: {len(valid_files)}"
    
    # 0.M1
    t = time.time()
    r_0_m1 = process_0_m1(valid_files)
    print(f"  [1/14] 0.M1: Trẻ MN 5t (I5) = {r_0_m1['aggregated_cells']['I5']:,.0f}, Tổng MN (K5) = {r_0_m1['aggregated_cells']['K5']:,.0f} ({time.time()-t:.2f}s)")
    
    # 0.M2
    t = time.time()
    r_0_m2 = process_0_m2(valid_files)
    print(f"  [2/14] 0.M2: Lớp MG 3,4t = {r_0_m2['aggregated_cells']['G8']:,.0f}, Lớp MG 5t = {r_0_m2['aggregated_cells']['G9']:,.0f} ({time.time()-t:.2f}s)")

    # 0.M3
    t = time.time()
    r_0_m3 = process_0_m3(valid_files)
    print(f"  [3/14] 0.M3: Tổng CBQL+GV+NV (C8) = {r_0_m3['aggregated_cells']['C8']:,.0f}, TL GV/lớp (G8) = {r_0_m3['aggregated_cells']['G8']:.2f} ({time.time()-t:.2f}s)")

    # 0.M4
    t = time.time()
    r_0_m4 = process_0_m4(valid_files)
    print(f"  [4/14] 0.M4: Số trường/cơ sở (C13) = {r_0_m4['aggregated_cells']['C13']:,.0f}, Lớp MG 3,4t (H13) = {r_0_m4['aggregated_cells']['H13']:,.0f}, Tỷ lệ phòng/lớp (L13) = {r_0_m4['aggregated_cells']['L13']:.2f} ({time.time()-t:.2f}s)")

    # 0.M5
    t = time.time()
    r_0_m5 = process_0_m5(valid_files)
    print(f"  [5/14] 0.M5: Tổng chi NS (E8) = {r_0_m5['aggregated_cells']['E8']:,.2f} Tr.đ, TL chi CM (H13) = {r_0_m5['aggregated_cells']['H13']:.2f}% ({time.time()-t:.2f}s)")

    # 0.M6
    t = time.time()
    r_0_m6 = process_0_m6(valid_files)
    print(f"  [6/14] 0.M6: Tổng trẻ KT MN (C13) = {r_0_m6['aggregated_cells']['C13']:,.0f}, Được can thiệp (L13) = {r_0_m6['aggregated_cells']['L13']:,.0f} ({time.time()-t:.2f}s)")

    # 0.M7
    t = time.time()
    r_0_m7 = process_0_m7(valid_files, collect_records=False)
    print(f"  [7/14] 0.M7: Tổng học sinh trong sổ = {r_0_m7['total_children']:,} trẻ, Số trẻ nữ = {r_0_m7['total_female']:,} ({r_0_m7['rate_female']:.2f}%) ({time.time()-t:.2f}s)")

    # 1.M1
    t = time.time()
    r_1_m1 = process_1_m1(valid_files)
    print(f"  [8/14] 1.M1: Trẻ 6t (F8) = {r_1_m1['aggregated_cells']['F8']:,.0f}, Tổng 6-10t (K8) = {r_1_m1['aggregated_cells']['K8']:,.0f}, Vào L1 (G58) = {r_1_m1['aggregated_cells']['G58']:.2f}% ({time.time()-t:.2f}s)")

    # 1.M2
    t = time.time()
    r_1_m2 = process_1_m2(valid_files)
    print(f"  [9/14] 1.M2: Tổng GV TH (I8) = {r_1_m2['aggregated_cells']['I8']:,.0f}, Biên chế (J8) = {r_1_m2['aggregated_cells']['J8']:,.0f}, TL GV/lớp (N8) = {r_1_m2['aggregated_cells']['N8']:.2f} ({time.time()-t:.2f}s)")

    # 1.M3
    t = time.time()
    r_1_m3 = process_1_m3(valid_files)
    print(f"  [10/14] 1.M3: Điểm trường (C8) = {r_1_m3['aggregated_cells']['C8']:,.0f}, Tổng lớp (D8) = {r_1_m3['aggregated_cells']['D8']:,.0f}, Phòng học = {r_1_m3['total_rooms']:,.0f}, TL Ph/lớp (J8) = {r_1_m3['aggregated_cells']['J8']:.2f}% ({time.time()-t:.2f}s)")

    # 1.M4
    t = time.time()
    r_1_m4 = process_1_m4(valid_files)
    print(f"  [11/14] 1.M4: Trường TH (C8) = {r_1_m4['aggregated_cells']['C8']:,.0f}, Tổng HS TH (D8) = {r_1_m4['aggregated_cells']['D8']:,.0f}, 6t vào L1 = {r_1_m4['aggregated_cells']['G8']:.2f}%, 14t HTCTTH = {r_1_m4['aggregated_cells']['M8']:.2f}%, Chuẩn = {r_1_m4['provincial_standard']} ({time.time()-t:.2f}s)")

    # 1.M5
    t = time.time()
    r_1_m5 = process_1_m5(valid_files)
    print(f"  [12/14] 1.M5: Tổng trẻ KT TH (C17) = {r_1_m5['aggregated_cells']['C17']:,.0f}, Tiếp cận GD (L17) = {r_1_m5['aggregated_cells']['L17']:,.0f}, Tỉ lệ = {r_1_m5['aggregated_cells']['M17']:.2f}% ({time.time()-t:.2f}s)")

    # 1.M6
    t = time.time()
    r_1_m6 = process_1_m6(valid_files)
    print(f"  [13/15] 1.M6: Tổng HS học ngoài (AR49) = {r_1_m6['total_grand']:,.0f}, Ngoài xã trong tỉnh = {r_1_m6['total_part1']:,.0f}, Ngoài tỉnh = {r_1_m6['total_part2']:,.0f} ({time.time()-t:.2f}s)")

    # 1.M7
    t = time.time()
    r_1_m7 = process_1_m7(valid_files)
    print(f"  [14/15] 1.M7: Tổng HS ngoại địa bàn đến học (AR7) = {r_1_m7['total_grand']:,.0f}, Trong tỉnh = {r_1_m7['total_part1']:,.0f}, Ngoài tỉnh = {r_1_m7['total_part2']:,.0f} ({time.time()-t:.2f}s)")

    # 2.M1
    t = time.time()
    r_2_m1 = process_2_m1(valid_files)
    print(f"  [15/16] 2.M1: TTN 15t (N7) = {r_2_m1['aggregated_cells']['N7']:,.0f}, TN THCS hàng năm (H55) = {r_2_m1['aggregated_cells']['H55']:.2f}%, Học tiếp (H59) = {r_2_m1['aggregated_cells']['H59']:.2f}% ({time.time()-t:.2f}s)", flush=True)

    # 2.M2
    t = time.time()
    r_2_m2 = process_2_m2(valid_files)
    p_cnt = r_2_m2['province_counts']
    p_rt = r_2_m2['province_rates']
    print(f"  [16/17] 2.M2: Trẻ 6t vào L1 (D48) = {p_cnt['D']:,.0f} ({p_rt['E']:.2f}%), Vào Lớp 6 (I48) = {p_cnt['I']:,.0f} ({p_rt['J']:.2f}%), 11-14t HTCTTH (L48) = {p_cnt['L']:,.0f} ({p_rt['M']:.2f}%), 15-18t có bằng THCS (U48) = {p_cnt['U']:,.0f} ({p_rt['V']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 2.M3
    t = time.time()
    r_2_m3 = process_2_m3(valid_files)
    p3_cnt = r_2_m3['province_counts']
    p3_rt = r_2_m3['province_rates']
    print(f"  [17/18] 2.M3: TTN 11-18t (D46) = {p3_cnt['D']:,.0f}, TN THCS 15-18t (I46) = {p3_cnt['I']:,.0f} ({p3_rt['J']:.2f}%), Học tiếp GDPT/TX/NN (K46) = {p3_cnt['K']:,.0f} ({p3_rt['L']:.2f}%), KT tiếp cận GD (O46) = {p3_cnt['O']:,.0f} ({p3_rt['P']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 2.M4
    t = time.time()
    r_2_m4 = process_2_m4(valid_files)
    p4_cnt = r_2_m4['province_counts']
    print(f"  [18/19] 2.M4: Số trường THCS = {r_2_m4['total_schools']}, Lớp học (D12) = {p4_cnt['D']:,.0f}, Phòng học = {r_2_m4['province_tot_rooms']:,.0f} (KC: {p4_cnt['E']:,.0f}), TL Ph/lớp (H12) = {r_2_m4['province_rate_h']:.2f}% ({time.time()-t:.2f}s)", flush=True)

    # 2.M5
    t = time.time()
    r_2_m5 = process_2_m5(valid_files)
    p5_cnt = r_2_m5['province_counts']
    p5_cr = r_2_m5['criteria']
    print(f"  [19/20] 2.M5: Tổng GV THCS (H12) = {p5_cnt['H']:,.0f} (BC: {p5_cnt['I']:,.0f}, HĐ: {p5_cnt['J']:,.0f}), TL GV/lớp (M12) = {r_2_m5['prov_ratio_gv_class']:.2f}, Đạt chuẩn ĐH = {p5_cr['dat_chuan']['tl']:.2f}%, Chuẩn NN = {p5_cr['chuan_nghe_nghiep']['tl']:.2f}% ({time.time()-t:.2f}s)", flush=True)

    # 2.M6
    t = time.time()
    r_2_m6 = process_2_m6(valid_files)
    print(f"  [20/21] 2.M6: Tổng đối tượng KT THCS (C14) = {r_2_m6['total_kt']:,.0f}, Tiếp cận GD (L14) = {r_2_m6['total_access']:,.0f} (Tỷ lệ = {r_2_m6['rate_access']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 2.M7
    t = time.time()
    r_2_m7 = process_2_m7(valid_files)
    print(f"  [21/22] 2.M7: HS học ngoài xã trong tỉnh (AM21) = {r_2_m7['total_part1']:,.0f}, Ngoài tỉnh (AM39) = {r_2_m7['total_part2']:,.0f}, Tổng cộng (AM40) = {r_2_m7['total_grand']:,.0f} ({r_2_m7['total_schools_count']} lượt trường) ({time.time()-t:.2f}s)", flush=True)

    # 2.M8
    t = time.time()
    r_2_m8 = process_2_m8(valid_files)
    print(f"  [22/23] 2.M8: HS ngoại địa bàn đến học trong tỉnh (AM24) = {r_2_m8['total_part1']:,.0f}, Ngoài tỉnh (AM38) = {r_2_m8['total_part2']:,.0f}, Tổng cộng (AM39) = {r_2_m8['total_grand']:,.0f} ({r_2_m8['total_schools_count']} lượt trường) ({time.time()-t:.2f}s)", flush=True)

    # 2.M9
    t = time.time()
    r_2_m9 = process_2_m9(valid_files)
    p9_cnt = r_2_m9['province_counts']
    p9_rt = r_2_m9['province_rates']
    print(f"  [23/24] 2.M9: Đối tượng PC 15-21 (C48) = {p9_cnt['C']:,.0f}, TN THCS 15-18 (E48) = {p9_cnt['E']:,.0f} (Tỷ lệ = {p9_rt['F']:.2f}%), Lớp 9 TN năm qua (J48) = {p9_cnt['J']:,.0f} ({p9_rt['K']:.2f}%), Có bằng TN THPT/Nghề (V48) = {p9_rt['V']:.2f}% ({time.time()-t:.2f}s)", flush=True)

    # 2.M10
    t = time.time()
    r_2_m10 = process_2_m10(valid_files)
    p10_cnt = r_2_m10['province_counts']
    p10_rt = r_2_m10['province_rates']
    print(f"  [24/33] 2.M10: Đối tượng PC 18-21 (C47) = {p10_cnt['C']:,.0f}, Có bằng TN THPT/BT/Nghề = {p10_cnt['E']+p10_cnt['F']+p10_cnt['G']:,.0f} (Tỷ lệ = {p10_rt['H']:.2f}%), TN THPT năm qua = {p10_cnt['J']:,.0f} ({p10_rt['K']:.2f}%), Có bằng nghề (S47) = {p10_cnt['S']:,.0f} ({p10_rt['T']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 3.M1
    t = time.time()
    c_data_3m1 = [process_3_m1_single_file(f['filepath'], f['commune_name']) for f in valid_files]
    r_3_m1 = aggregate_3_m1(c_data_3m1)
    print(f"  [25/33] 3.M1: Dân số 15-25t (C) = {r_3_m1['sum_c_dan_so']:,.0f}, Biết chữ M1 (D) = {r_3_m1['sum_d_biet_chu_m1']:,.0f} ({r_3_m1['rate_m1_prov']:.2f}%), Biết chữ M2 (F) = {r_3_m1['sum_f_biet_chu_m2']:,.0f} ({r_3_m1['rate_m2_prov']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 3.M2
    t = time.time()
    c_data_3m2 = [process_3_m2_single_file(f['filepath'], f['commune_name']) for f in valid_files]
    r_3_m2 = aggregate_3_m2(c_data_3m2)
    print(f"  [26/33] 3.M2: Dân số 15-35t (C) = {r_3_m2['sum_c_dan_so']:,.0f}, Biết chữ M1 (D) = {r_3_m2['sum_d_biet_chu_m1']:,.0f} ({r_3_m2['rate_m1_prov']:.2f}%), Biết chữ M2 (F) = {r_3_m2['sum_f_biet_chu_m2']:,.0f} ({r_3_m2['rate_m2_prov']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 3.M3
    t = time.time()
    c_data_3m3 = [process_3_m3_single_file(f['filepath'], f['commune_name']) for f in valid_files]
    r_3_m3 = aggregate_3_m3(c_data_3m3)
    print(f"  [27/33] 3.M3: Dân số 15-60t (C) = {r_3_m3['sum_c_dan_so']:,.0f}, Biết chữ M1 (D) = {r_3_m3['sum_d_biet_chu_m1']:,.0f} ({r_3_m3['rate_m1_prov']:.2f}%), Biết chữ M2 (F) = {r_3_m3['sum_f_biet_chu_m2']:,.0f} ({r_3_m3['rate_m2_prov']:.2f}%) ({time.time()-t:.2f}s)", flush=True)

    # 3.M4 .. 3.M9
    for idx_m, sc in enumerate(['3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9'], start=28):
        t = time.time()
        c_data_mu_chu = [process_3_mu_chu_single_file(f['filepath'], f['commune_name'], sc) for f in valid_files]
        r_mu_chu = aggregate_3_mu_chu(c_data_mu_chu, sc)
        print(f"  [{idx_m}/33] {sc}: Dân số = {r_mu_chu['sum_c_dan_so']:,.0f}, Mù chữ (G) = {r_mu_chu['sum_g_mu_chu']:,.0f} ({r_mu_chu['rate_h_prov']:.2f}%), Nữ = {r_mu_chu['sum_i_mu_chu_nu']:,.0f}, DT = {r_mu_chu['sum_k_mu_chu_dt']:,.0f} ({time.time()-t:.2f}s)", flush=True)

    dur_all = time.time() - t0
    print(f"\n=========================================================================================", flush=True)
    print(f">>> HOÀN THÀNH TỔNG HỢP TOÀN BỘ 33/33 BIỂU NGHIỆP VỤ TRÊN 92 XÃ TRONG {dur_all:.2f} GIÂY! <<<", flush=True)
    print(f">>> 100% KHỐI MẦM NON (7/7) | 100% TIỂU HỌC (7/7) | 100% THCS (10/10) | 100% XMC (9/9) <<<", flush=True)
    print(f"=========================================================================================\n", flush=True)

if __name__ == '__main__':
    run_all()

