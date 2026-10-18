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
from processors.m2_m1 import process_2_m1

def run_all():
    data_dir = os.path.join(BASE_DIR, 'data')
    print("=== CHẠY KIỂM THỬ HỒI QUY TẤT CẢ 9 BIỂU ĐÃ HOÀN THÀNH TRÊN 92 XÃ ===")
    t0 = time.time()
    
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, có: {len(valid_files)}"
    
    # 0.M1
    t = time.time()
    r_0_m1 = process_0_m1(valid_files)
    print(f"  [1/9] 0.M1: Trẻ MN 5t (I5) = {r_0_m1['aggregated_cells']['I5']:,.0f}, Tổng MN (K5) = {r_0_m1['aggregated_cells']['K5']:,.0f} ({time.time()-t:.2f}s)")
    
    # 0.M2
    t = time.time()
    r_0_m2 = process_0_m2(valid_files)
    print(f"  [2/9] 0.M2: Lớp MG 3,4t = {r_0_m2['aggregated_cells']['G8']:,.0f}, Lớp MG 5t = {r_0_m2['aggregated_cells']['G9']:,.0f} ({time.time()-t:.2f}s)")

    # 0.M3
    t = time.time()
    r_0_m3 = process_0_m3(valid_files)
    print(f"  [3/9] 0.M3: Tổng CBQL+GV+NV (C8) = {r_0_m3['aggregated_cells']['C8']:,.0f}, TL GV/lớp (G8) = {r_0_m3['aggregated_cells']['G8']:.2f} ({time.time()-t:.2f}s)")

    # 0.M4
    t = time.time()
    r_0_m4 = process_0_m4(valid_files)
    print(f"  [4/9] 0.M4: Số trường/cơ sở (C13) = {r_0_m4['aggregated_cells']['C13']:,.0f}, Lớp MG 3,4t (H13) = {r_0_m4['aggregated_cells']['H13']:,.0f}, Tỷ lệ phòng/lớp (L13) = {r_0_m4['aggregated_cells']['L13']:.2f} ({time.time()-t:.2f}s)")

    # 0.M5
    t = time.time()
    r_0_m5 = process_0_m5(valid_files)
    print(f"  [5/9] 0.M5: Tổng chi NS (E8) = {r_0_m5['aggregated_cells']['E8']:,.2f} Tr.đ, TL chi CM (H13) = {r_0_m5['aggregated_cells']['H13']:.2f}% ({time.time()-t:.2f}s)")

    # 0.M6
    t = time.time()
    r_0_m6 = process_0_m6(valid_files)
    print(f"  [6/9] 0.M6: Tổng trẻ KT MN (C13) = {r_0_m6['aggregated_cells']['C13']:,.0f}, Được can thiệp (L13) = {r_0_m6['aggregated_cells']['L13']:,.0f} ({time.time()-t:.2f}s)")

    # 0.M7
    t = time.time()
    r_0_m7 = process_0_m7(valid_files, collect_records=False)
    print(f"  [7/9] 0.M7: Tổng học sinh trong sổ = {r_0_m7['total_children']:,} trẻ, Số trẻ nữ = {r_0_m7['total_female']:,} ({r_0_m7['rate_female']:.2f}%) ({time.time()-t:.2f}s)")

    # 1.M1
    t = time.time()
    r_1_m1 = process_1_m1(valid_files)
    print(f"  [8/9] 1.M1: Trẻ 6t (F8) = {r_1_m1['aggregated_cells']['F8']:,.0f}, Tổng 6-10t (K8) = {r_1_m1['aggregated_cells']['K8']:,.0f}, Vào L1 (G58) = {r_1_m1['aggregated_cells']['G58']:.2f}% ({time.time()-t:.2f}s)")

    # 2.M1
    t = time.time()
    r_2_m1 = process_2_m1(valid_files)
    print(f"  [9/9] 2.M1: TTN 15t (N7) = {r_2_m1['aggregated_cells']['N7']:,.0f}, TN THCS hàng năm (H55) = {r_2_m1['aggregated_cells']['H55']:.2f}%, Học tiếp (H59) = {r_2_m1['aggregated_cells']['H59']:.2f}% ({time.time()-t:.2f}s)")

    dur_all = time.time() - t0
    print(f"\n>>> HOÀN THÀNH TỔNG HỢP CẢ 9 BIỂU TRÊN 92 XÃ TRONG {dur_all:.2f} GIÂY! <<<")

if __name__ == '__main__':
    run_all()
