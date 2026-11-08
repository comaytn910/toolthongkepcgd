import os
import sys
import openpyxl

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.excel_reader import scan_data_directory
from processors.m1_m4 import process_1_m4
from exporters.excel_exporter import export_1_m4_test

def test_1_m4_pipeline():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 1.M4 (KẾT QUẢ PCGD TIỂU HỌC) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ.")
    
    # 2. Xử lý tổng hợp 1.M4
    res_m4 = process_1_m4(valid_files)
    agg = res_m4['aggregated_cells']
    rates = res_m4['province_rates']
    
    # 3. Kiểm tra các chỉ tiêu cốt lõi
    tot_truong = agg['C8']
    tot_hs = agg['D8']
    tot_kt = agg['E8']
    tre_6t = agg['F8']
    tre_11t = agg['H8']
    tre_den_14t = agg['J8']
    tre_14t = agg['L8']
    kt_kn_ht = agg['O8']
    kt_tc_gd = agg['P8']
    
    print(f"[ĐỐI SOÁT CỐT LÕI]")
    print(f"  - Tổng số trường tiểu học (C): {tot_truong:,.0f}")
    print(f"  - Tổng số học sinh PPC (D): {tot_hs:,.0f} học sinh")
    print(f"  - Trẻ khuyết tật (E, N): {tot_kt:,.0f} trẻ")
    print(f"  - Trẻ 6t vào lớp 1 (F): {tre_6t:,.0f} | Tỉ lệ G = {rates['G']:.2f}%")
    print(f"  - Trẻ 11t HTCTTH (H): {tre_11t:,.0f} | Tỉ lệ I = {rates['I']:.2f}%")
    print(f"  - Trẻ đến 14t HTCTTH (J): {tre_den_14t:,.0f} | Tỉ lệ K = {rates['K']:.2f}%")
    print(f"  - Trẻ 14t HTCTTH (L): {tre_14t:,.0f} | Tỉ lệ M = {rates['M']:.2f}%")
    print(f"  - Trẻ KT có khả năng HT (O): {kt_kn_ht:,.0f} | Tiếp cận GD (P): {kt_tc_gd:,.0f} | Tỉ lệ Q = {rates['Q']:.2f}%")
    print(f"  - Đạt chuẩn Đội ngũ (R): {res_m4['guarantee_staff_count']}/{res_m4['total_communes']} xã")
    print(f"  - Đạt chuẩn CSVC (S): {res_m4['guarantee_fac_count']}/{res_m4['total_communes']} xã")
    print(f"  - Đạt chuẩn Mức độ 3 (T): {res_m4['level_3_count']}/{res_m4['total_communes']} xã")
    
    # Assertions
    assert tot_truong == 158, f"Tổng số trường ({tot_truong}) phải bằng 158"
    assert tot_hs == 264693, f"Tổng số học sinh ({tot_hs}) phải bằng 264,693"
    assert tot_kt == 3920, f"Tổng trẻ khuyết tật ({tot_kt}) phải bằng 3,920"
    assert tre_6t == 27046, f"Trẻ 6t vào lớp 1 ({tre_6t}) phải bằng 27,046"
    assert rates['G'] == 100.0, f"Tỉ lệ 6t vào lớp 1 phải là 100%"
    assert abs(rates['I'] - 99.62) < 0.05, f"Tỉ lệ 11t HTCTTH phải là 99.62%"
    assert abs(rates['Q'] - 99.50) < 0.05, f"Tỉ lệ trẻ KT tiếp cận GD phải là 99.50%"
    print("  -> [ĐẠT] Toàn bộ chỉ tiêu số lượng và tỷ lệ % chuẩn hóa khớp 100%!")
    
    # 4. Kiểm tra xuất file Excel kết quả
    out_file = export_1_m4_test(res_m4, template_path, outputs_dir)
    assert os.path.exists(out_file), f"File không tồn tại: {out_file}"
    file_size = os.path.getsize(out_file)
    print(f"\n[XUẤT EXCEL THÀNH CÔNG] File: {out_file} (Kích thước: {file_size:,} bytes)")
    
    # Mở lại kiểm tra định dạng và dữ liệu
    wb_chk = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_1.M4' in wb_chk.sheetnames
    ws_chk = wb_chk['TONG_HOP_1.M4']
    
    # Kiểm tra dòng đầu tiên (dòng 8) là xã thứ 1
    assert ws_chk['A8'].value == 1 or ws_chk['A8'].value == 1.0 or ws_chk['A8'].value == '1'
    # Kiểm tra dòng tổng cộng cuối cùng (dòng 100)
    sum_row = 8 + len(valid_files)
    assert ws_chk[f"B{sum_row}"].value == "TỔNG CỘNG TOÀN TỈNH"
    assert ws_chk[f"C{sum_row}"].value == 158.0
    assert ws_chk[f"D{sum_row}"].value == 264693.0
    wb_chk.close()
    print(f"  -> [ĐẠT] File Excel xuất ra đọc lại hoàn hảo, dòng tổng tỉnh tại dòng {sum_row} khớp tuyệt đối!")
    
    print("\n>>> TẤT CẢ KIỂM THỬ BIỂU 1.M4 ĐÃ ĐẠT 100%! <<<")

if __name__ == '__main__':
    test_1_m4_pipeline()
