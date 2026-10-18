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
from processors.m0_m5 import process_0_m5
from processors.m0_m6 import process_0_m6
from processors.m1_m1 import process_1_m1
from processors.m2_m1 import process_2_m1

from exporters.excel_exporter import (
    export_0_m5_test,
    export_0_m6_test,
    export_1_m1_test,
    export_2_m1_test
)

def run_remaining():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    scan_res = scan_data_directory(data_dir)
    valid_files = [f for f in scan_res['files'] if f['is_valid']]
    print(f"Tổng số xã hợp lệ: {len(valid_files)}/92", flush=True)
    
    # 1. BIỂU 0.M5
    print("\n>>> BẮT ĐẦU TỔNG HỢP BIỂU 0.M5 (Tài chính GDMN)...", flush=True)
    t0 = time.time()
    res_0_m5 = process_0_m5(valid_files)
    out_0_m5 = export_0_m5_test(res_0_m5, template_path, outputs_dir)
    dur = time.time() - t0
    agg5 = res_0_m5['aggregated_cells']
    print(f"  [XONG 0.M5] trong {dur:.2f}s -> {os.path.basename(out_0_m5)} ({os.path.getsize(out_0_m5):,} bytes)", flush=True)
    print(f"  Tổng chi NS tỉnh (E8): {agg5.get('E8', 0):,.2f} Tr.đg", flush=True)
    print(f"  Tổng kinh phí (D8): {agg5.get('D8', 0):,.2f} Tr.đg", flush=True)
    print(f"  Tỷ lệ chi chuyên môn (H13): {agg5.get('E13', 0):.2f}%", flush=True)
    
    # 2. BIỂU 0.M6
    print("\n>>> BẮT ĐẦU TỔNG HỢP BIỂU 0.M6 (Trẻ khuyết tật MN)...", flush=True)
    t0 = time.time()
    res_0_m6 = process_0_m6(valid_files)
    out_0_m6 = export_0_m6_test(res_0_m6, template_path, outputs_dir)
    dur = time.time() - t0
    agg6 = res_0_m6['aggregated_cells']
    print(f"  [XONG 0.M6] trong {dur:.2f}s -> {os.path.basename(out_0_m6)} ({os.path.getsize(out_0_m6):,} bytes)", flush=True)
    print(f"  Tổng số trẻ KT MN (C13): {agg6.get('C13', 0):,} trẻ", flush=True)
    print(f"  Số trẻ can thiệp (L13): {agg6.get('L13', 0):,} trẻ", flush=True)
    print(f"  Tỷ lệ can thiệp (M13): {agg6.get('M13', 0):.2f}%", flush=True)
    
    # 3. BIỂU 1.M1
    print("\n>>> BẮT ĐẦU TỔNG HỢP BIỂU 1.M1 (PCGD Tiểu học)...", flush=True)
    t0 = time.time()
    res_1_m1 = process_1_m1(valid_files)
    out_1_m1 = export_1_m1_test(res_1_m1, template_path, outputs_dir)
    dur = time.time() - t0
    agg1 = res_1_m1['aggregated_cells']
    print(f"  [XONG 1.M1] trong {dur:.2f}s -> {os.path.basename(out_1_m1)} ({os.path.getsize(out_1_m1):,} bytes)", flush=True)
    print(f"  Tổng số trẻ 6 tuổi (F8): {agg1.get('F8', 0):,} trẻ", flush=True)
    print(f"  Tiêu chí H58 (% Trẻ 6t vào Lớp 1): {agg1.get('G58', 0):.2f}%", flush=True)
    print(f"  Tiêu chí H61 (% Trẻ 11t HTCTTH): {agg1.get('G61', 0):.2f}%", flush=True)
    print(f"  Tiêu chí H62 (% Trẻ 14t HTCTTH): {agg1.get('G62', 0):.2f}%", flush=True)
    
    # 4. BIỂU 2.M1
    print("\n>>> BẮT ĐẦU TỔNG HỢP BIỂU 2.M1 (PCGD THCS)...", flush=True)
    t0 = time.time()
    res_2_m1 = process_2_m1(valid_files)
    out_2_m1 = export_2_m1_test(res_2_m1, template_path, outputs_dir)
    dur = time.time() - t0
    agg2 = res_2_m1['aggregated_cells']
    print(f"  [XONG 2.M1] trong {dur:.2f}s -> {os.path.basename(out_2_m1)} ({os.path.getsize(out_2_m1):,} bytes)", flush=True)
    print(f"  Tổng TTN 15t (N7): {agg2.get('N7', 0):,} người", flush=True)
    print(f"  Tiêu chí H55 (% TN THCS hàng năm): {agg2.get('H55', 0):.2f}%", flush=True)
    print(f"  Tiêu chí H58 (% 15-18t TN THCS): {agg2.get('H58', 0):.2f}%", flush=True)
    print(f"  Tiêu chí H59 (% Học tiếp sau THCS): {agg2.get('H59', 0):.2f}%", flush=True)

if __name__ == '__main__':
    run_remaining()
