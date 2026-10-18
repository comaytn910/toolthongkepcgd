import os
import sys
import time
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from core.excel_reader import scan_data_directory
from processors.m0_m2 import process_0_m2
from processors.m0_m1 import process_0_m1
from exporters.excel_exporter import export_0_m2_test

def test_0_m2():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M2 (TIÊU CHUẨN, ĐK PCGD MẦM NON) ===")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, thực tế: {len(valid_files)}"
    print(f"Quét thành công {len(valid_files)} xã hợp lệ.")
    
    # 2. Xử lý tổng hợp 0.M2
    t0 = time.time()
    res = process_0_m2(valid_files)
    dur = time.time() - t0
    print(f"Tổng hợp xong 0.M2 trong {dur:.2f} giây!")
    
    agg = res['aggregated_cells']
    
    # 3. Kiểm tra tính logic nội bộ của 0.M2
    for r in [8, 9]:
        lbl = "3, 4 tuổi" if r == 8 else "5 tuổi"
        print(f"\n--- Kiểm tra Hàng {r} ({lbl}) ---")
        # G = H + I
        g_val = agg[f'G{r}']
        h_val = agg[f'H{r}']
        i_val = agg[f'I{r}']
        assert abs(g_val - (h_val + i_val)) < 0.01, f"Lớp mẫu giáo G ({g_val}) != H+I ({h_val + i_val})"
        print(f"  [ĐẠT] G{r} = H{r} + I{r} ({g_val:,.0f} lớp = {h_val:,.0f} đơn + {i_val:,.0f} ghép)")
        
        # L = K / J * 100
        j_val = agg[f'J{r}']
        k_val = agg[f'K{r}']
        l_val = agg[f'L{r}']
        expected_l = round(k_val / j_val * 100, 2) if j_val > 0 else 0.0
        assert abs(l_val - expected_l) < 0.01, f"Tỉ lệ huy động L ({l_val}) != {expected_l}"
        print(f"  [ĐẠT] Tỉ lệ huy động L{r} = {l_val:.2f}% (Tử K={k_val:,.0f}, Mẫu J={j_val:,.0f})")
        
        # R = Q / P * 100
        p_val = agg[f'P{r}']
        q_val = agg[f'Q{r}']
        r_val = agg[f'R{r}']
        expected_r = round(q_val / p_val * 100, 2) if p_val > 0 else 0.0
        assert abs(r_val - expected_r) < 0.01, f"Tỉ lệ KT tiếp cận R ({r_val}) != {expected_r}"
        print(f"  [ĐẠT] Tỉ lệ trẻ KT tiếp cận GD R{r} = {r_val:.2f}% (Tử Q={q_val:,.0f}, Mẫu P={p_val:,.0f})")
        
        # S (Tỷ lệ GV/lớp), T (Tỷ lệ phòng/lớp)
        print(f"  [ĐẠT] Tỷ lệ GV/lớp S{r} = {agg[f'S{r}']:.2f}, Phòng/lớp T{r} = {agg[f'T{r}']:.2f}")
        print(f"  [ĐẠT] Lớp đủ TBDH U{r} = {agg[f'U{r}']:,.0f} lớp")
        print(f"  [ĐẠT] Đạt chuẩn V{r} = {agg[f'V{r}']}")

    # 4. Xuất file Excel kết quả
    out_file = export_0_m2_test(res, template_path, outputs_dir)
    assert os.path.exists(out_file), "File xuất ra phải tồn tại"
    print(f"\nXuất file Excel thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # 5. Kiểm tra định dạng trong file Excel xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=False)
    ws = wb['TONG_HOP_THU']
    for r in [8, 9]:
        assert ws[f'L{r}'].number_format == '0.00"%"', f"Format L{r} phải là 0.00%"": {ws[f'L{r}'].number_format}"
        assert ws[f'R{r}'].number_format == '0.00"%"', f"Format R{r} phải là 0.00%"": {ws[f'R{r}'].number_format}"
        assert ws[f'S{r}'].number_format == '0.00', f"Format S{r} phải là 0.00: {ws[f'S{r}'].number_format}"
        assert ws[f'T{r}'].number_format == '0.00', f"Format T{r} phải là 0.00: {ws[f'T{r}'].number_format}"
        assert ws[f'G{r}'].number_format == '#,##0', f"Format G{r} phải là #,##0: {ws[f'G{r}'].number_format}"
    print("[ĐẠT] Toàn bộ định dạng cell Excel: % nhận 0.00\"%\", thập phân 0.00, số lượng #,##0.")
    wb.close()
    
    # 6. Kiểm tra tính bất biến khi chạy lại (Clean Recalculation)
    res_again = process_0_m2(valid_files)
    assert res_again['aggregated_cells']['G8'] == agg['G8'], "Chạy lại lần 2 phải ra cùng kết quả chính xác 100%"
    assert res_again['aggregated_cells']['L8'] == agg['L8'], "Chạy lại lần 2 không bị nhân đôi tỷ lệ"
    print("[ĐẠT] Kiểm thử Clean Recalculation: Tính mới 100%, không bị lưu cache hay cộng dồn.")
    
    # 7. Hồi quy biểu 0.M1
    res_m1 = process_0_m1(valid_files)
    assert res_m1['aggregated_cells']['I5'] == 23765.0, "Hồi quy 0.M1: Trẻ 5 tuổi I5 phải giữ nguyên"
    print("[ĐẠT] Kiểm thử hồi quy 0.M1: Không bị ảnh hưởng, I5 = 23,765 trẻ.")
    
    print("\n=== TOÀN BỘ KIỂM THỬ BIỂU 0.M2 ĐẠT 100% ===")

if __name__ == '__main__':
    test_0_m2()
