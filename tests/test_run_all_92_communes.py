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
from processors.m0_m5 import process_0_m5
from processors.m0_m6 import process_0_m6
from processors.m1_m1 import process_1_m1
from processors.m2_m1 import process_2_m1

from exporters.excel_exporter import (
    export_0_m1_test,
    export_0_m5_test,
    export_0_m6_test,
    export_1_m1_test,
    export_2_m1_test
)

def test_pipeline_92():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU CHẠY KIỂM THỬ TỔNG HỢP TOÀN BỘ NGUỒN 92 XÃ ===")
    t0 = time.time()
    
    scan_res = scan_data_directory(data_dir)
    valid_files = [f for f in scan_res['files'] if f['is_valid']]
    print(f"1. Quét dữ liệu: Tổng số file = {len(scan_res['files'])}, File hợp lệ = {len(valid_files)}")
    assert len(valid_files) == 92, f"Kỳ vọng 92 xã hợp lệ, thực tế: {len(valid_files)}"
    
    # 1. BIỂU 0.M1
    print("\n--- [1/5] TỔNG HỢP BIỂU 0.M1 (92 XÃ) ---")
    t_start = time.time()
    res_0_m1 = process_0_m1(valid_files)
    out_0_m1 = export_0_m1_test(res_0_m1, template_path, outputs_dir)
    print(f"  + Hoàn thành sau {time.time()-t_start:.2f}s -> {os.path.basename(out_0_m1)}")
    # Kiểm tra một số chỉ tiêu then chốt của 0.M1
    agg_0_m1 = res_0_m1['aggregated_cells']
    print(f"  + Tổng số trẻ 5 tuổi (Cột I, Hàng 5): {agg_0_m1.get('I5', 0):,}")
    print(f"  + Tổng số trẻ MN toàn tỉnh (Cột K, Hàng 5): {agg_0_m1.get('K5', 0):,}")
    print(f"  + Tỷ lệ huy động 5 tuổi (Hàng 16, Cột I): {agg_0_m1.get('I16', 0):.2f}%")
    print(f"  + Tỷ lệ chuyên cần 5 tuổi (Hàng 22, Cột I): {agg_0_m1.get('I22', 0):.2f}%")
    assert agg_0_m1.get('K5', 0) > 0, "Tổng số trẻ MN K5 phải > 0"
    
    # 2. BIỂU 0.M5
    print("\n--- [2/5] TỔNG HỢP BIỂU 0.M5 (92 XÃ) ---")
    t_start = time.time()
    res_0_m5 = process_0_m5(valid_files)
    out_0_m5 = export_0_m5_test(res_0_m5, template_path, outputs_dir)
    print(f"  + Hoàn thành sau {time.time()-t_start:.2f}s -> {os.path.basename(out_0_m5)}")
    agg_0_m5 = res_0_m5['aggregated_cells']
    print(f"  + Tổng chi ngân sách tỉnh (E8): {agg_0_m5.get('E8', 0):,.2f} Tr.đg")
    print(f"  + Tổng kinh phí GDMN (D8): {agg_0_m5.get('D8', 0):,.2f} Tr.đg")
    print(f"  + Tỉ lệ chi HĐ chuyên môn (Hàng 13): {agg_0_m5.get('E13', 0):.2f}%")
    assert agg_0_m5.get('E8', 0) > 0, "Tổng chi ngân sách E8 phải > 0"
    
    # 3. BIỂU 0.M6
    print("\n--- [3/5] TỔNG HỢP BIỂU 0.M6 (92 XÃ) ---")
    t_start = time.time()
    res_0_m6 = process_0_m6(valid_files)
    out_0_m6 = export_0_m6_test(res_0_m6, template_path, outputs_dir)
    print(f"  + Hoàn thành sau {time.time()-t_start:.2f}s -> {os.path.basename(out_0_m6)}")
    agg_0_m6 = res_0_m6['aggregated_cells']
    print(f"  + Tổng số trẻ khuyết tật MN toàn tỉnh (C13): {agg_0_m6.get('C13', 0):,}")
    print(f"  + Trẻ khuyết tật được can thiệp (L13): {agg_0_m6.get('L13', 0):,}")
    print(f"  + Tỷ lệ can thiệp khuyết tật MN (M13): {agg_0_m6.get('M13', 0):.2f}%")
    # Đối soát liên biểu: Trẻ 5 tuổi khuyết tật C12 so với I8 của 0.M1
    val_c12 = agg_0_m6.get('C12', 0)
    val_0_m1_i8 = agg_0_m1.get('I8', 0)
    print(f"  + Đối soát chéo 0.M6 C12 ({val_c12}) với 0.M1 I8 ({val_0_m1_i8}): {'KHỚP 100%' if val_c12 == val_0_m1_i8 else f'LỆCH: {val_c12} != {val_0_m1_i8}'}")
    
    # 4. BIỂU 1.M1
    print("\n--- [4/5] TỔNG HỢP BIỂU 1.M1 (92 XÃ) ---")
    t_start = time.time()
    res_1_m1 = process_1_m1(valid_files)
    out_1_m1 = export_1_m1_test(res_1_m1, template_path, outputs_dir)
    print(f"  + Hoàn thành sau {time.time()-t_start:.2f}s -> {os.path.basename(out_1_m1)}")
    agg_1_m1 = res_1_m1['aggregated_cells']
    print(f"  + Tổng số trẻ 6 tuổi toàn tỉnh (F8): {agg_1_m1.get('F8', 0):,}")
    print(f"  + Tổng số trẻ 6-10 tuổi (K8): {agg_1_m1.get('K8', 0):,}")
    print(f"  + Tiêu chí Hàng 58 (% Trẻ 6 tuổi vào lớp 1): {agg_1_m1.get('G58', 0):.2f}%")
    print(f"  + Tiêu chí Hàng 61 (% Trẻ 11t HTCTTH): {agg_1_m1.get('G61', 0):.2f}%")
    print(f"  + Tiêu chí Hàng 62 (% Trẻ 14t HTCTTH): {agg_1_m1.get('G62', 0):.2f}%")
    assert agg_1_m1.get('F8', 0) > 0, "Tổng số trẻ 6 tuổi F8 phải > 0"
    
    # 5. BIỂU 2.M1
    print("\n--- [5/5] TỔNG HỢP BIỂU 2.M1 (92 XÃ) ---")
    t_start = time.time()
    res_2_m1 = process_2_m1(valid_files)
    out_2_m1 = export_2_m1_test(res_2_m1, template_path, outputs_dir)
    print(f"  + Hoàn thành sau {time.time()-t_start:.2f}s -> {os.path.basename(out_2_m1)}")
    agg_2_m1 = res_2_m1['aggregated_cells']
    print(f"  + Tổng số TTN 15 tuổi (Hàng 7, N7): {agg_2_m1.get('N7', 0):,}")
    print(f"  + Tiêu chí H55 (% TN THCS hàng năm): {agg_2_m1.get('H55', 0):.2f}%")
    print(f"  + Tiêu chí H58 (% TTN 15-18t tốt nghiệp THCS): {agg_2_m1.get('H58', 0):.2f}%")
    print(f"  + Tiêu chí H59 (% TTN tốt nghiệp THCS đang học tiếp): {agg_2_m1.get('H59', 0):.2f}%")
    assert agg_2_m1.get('N7', 0) > 0, "Tổng số TTN 15 tuổi N7 phải > 0"
    
    total_duration = time.time() - t0
    print(f"\n=== TỔNG HỢP TOÀN BỘ 5 BIỂU TRÊN 92 XÃ THÀNH CÔNG TRONG {total_duration:.2f} GIÂY ===")

if __name__ == '__main__':
    test_pipeline_92()
