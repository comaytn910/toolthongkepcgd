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
from processors.m0_m3 import process_0_m3
from processors.m0_m2 import process_0_m2
from processors.m0_m1 import process_0_m1
from exporters.excel_exporter import export_0_m3_test

def test_0_m3():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M3 (ĐỘI NGŨ CBQL, GV, NV MN) ===")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, thực tế: {len(valid_files)}"
    print(f"Quét thành công {len(valid_files)} xã hợp lệ.")
    
    # 2. Xử lý tổng hợp 0.M3
    t0 = time.time()
    res = process_0_m3(valid_files)
    dur = time.time() - t0
    print(f"Tổng hợp xong 0.M3 trong {dur:.2f} giây!")
    
    agg = res['aggregated_cells']
    
    # 3. Kiểm tra tính logic nội bộ của 0.M3
    print("\n--- [Kiểm tra Hàng 8: Toàn trường & Lớp 3, 4 tuổi] ---")
    tot_c8 = agg['C8']
    tot_e8 = agg['E8']
    tot_f8 = agg['F8']
    tot_h8 = agg['H8']
    print(f"  + Tổng CBQL+GV+NV toàn tỉnh (C8): {tot_c8:,.0f} người")
    print(f"  + Trong đó: CBQL (E8) = {tot_e8:,.0f}, Giáo viên (F8) = {tot_f8:,.0f}, Nhân viên (H8) = {tot_h8:,.0f}")
    assert tot_c8 > 0, "Tổng CBQL+GV+NV C8 phải > 0"
    
    # Kiểm tra ô G8 không bị lỗi cộng dồn ngày tháng (46235.0)
    g8_val = agg['G8']
    print(f"  + Tỉ lệ GV/lớp toàn trường (G8): {g8_val:.2f} (Đã lọc sạch lỗi serial date 46235.0)")
    assert g8_val < 10.0, f"G8 bị lỗi serial date: {g8_val}"
    
    # Lớp 3, 4 tuổi:
    tot_j8 = agg['J8']
    tot_k8 = agg['K8']
    tot_n8 = agg['N8']
    tot_q8 = agg['Q8']
    tot_s8 = agg['S8']
    print(f"  + Lớp 3, 4 tuổi: Số lớp (J8) = {tot_j8:,.0f}, Tổng GV (K8) = {tot_k8:,.0f}, Tỉ lệ GV/lớp (N8) = {tot_n8:.2f}")
    print(f"  + Tỉ lệ GV đạt chuẩn trở lên (Q8) = {tot_q8:.2f}%, Đạt chuẩn nghề nghiệp (S8) = {tot_s8:.2f}%")
    assert 0 <= tot_q8 <= 100.0, "Tỉ lệ Q8 phải trong khoảng 0-100%"
    assert 0 <= tot_s8 <= 100.0, "Tỉ lệ S8 phải trong khoảng 0-100%"
    
    print("\n--- [Kiểm tra Hàng 9: Lớp 5 tuổi] ---")
    tot_j9 = agg['J9']
    tot_k9 = agg['K9']
    tot_n9 = agg['N9']
    tot_q9 = agg['Q9']
    tot_s9 = agg['S9']
    print(f"  + Lớp 5 tuổi: Số lớp (J9) = {tot_j9:,.0f}, Tổng GV (K9) = {tot_k9:,.0f}, Tỉ lệ GV/lớp (N9) = {tot_n9:.2f}")
    print(f"  + Tỉ lệ GV đạt chuẩn trở lên (Q9) = {tot_q9:.2f}%, Đạt chuẩn nghề nghiệp (S9) = {tot_s9:.2f}%")
    assert 0 <= tot_q9 <= 100.0, "Tỉ lệ Q9 phải trong khoảng 0-100%"
    assert 0 <= tot_s9 <= 100.0, "Tỉ lệ S9 phải trong khoảng 0-100%"

    # 4. Xuất file Excel kết quả
    out_file = export_0_m3_test(res, template_path, outputs_dir)
    assert os.path.exists(out_file), "File xuất ra phải tồn tại"
    print(f"\nXuất file Excel thành công: {out_file} ({os.path.getsize(out_file):,} bytes)")
    
    # 5. Kiểm tra định dạng trong file Excel xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=False)
    ws = wb['TONG_HOP_THU']
    assert ws['Q8'].number_format == '0.00"%"', f"Format Q8 phải là 0.00%"": {ws['Q8'].number_format}"
    assert ws['S8'].number_format == '0.00"%"', f"Format S8 phải là 0.00%"": {ws['S8'].number_format}"
    assert ws['G8'].number_format == '0.00', f"Format G8 phải là 0.00: {ws['G8'].number_format}"
    assert ws['N8'].number_format == '0.00', f"Format N8 phải là 0.00: {ws['N8'].number_format}"
    assert ws['C8'].number_format == '#,##0', f"Format C8 phải là #,##0: {ws['C8'].number_format}"
    print("[ĐẠT] Toàn bộ định dạng cell Excel: % nhận 0.00\"%\", thập phân 0.00, số lượng #,##0.")
    wb.close()
    
    # 6. Kiểm tra tính bất biến khi chạy lại (Clean Recalculation)
    res_again = process_0_m3(valid_files)
    assert res_again['aggregated_cells']['C8'] == agg['C8'], "Chạy lại lần 2 phải ra cùng kết quả chính xác 100%"
    assert res_again['aggregated_cells']['Q8'] == agg['Q8'], "Chạy lại lần 2 không bị nhân đôi tỷ lệ"
    print("[ĐẠT] Kiểm thử Clean Recalculation: Tính mới 100%, không bị lưu cache hay cộng dồn.")
    
    # 7. Hồi quy biểu 0.M1 và 0.M2
    res_m1 = process_0_m1(valid_files)
    assert res_m1['aggregated_cells']['I5'] == 23765.0, "Hồi quy 0.M1: Trẻ 5 tuổi I5 phải giữ nguyên"
    res_m2 = process_0_m2(valid_files)
    assert res_m2['aggregated_cells']['G8'] == 1843.0, "Hồi quy 0.M2: Tổng lớp G8 phải giữ nguyên"
    print("[ĐẠT] Kiểm thử hồi quy 0.M1 và 0.M2: Không bị ảnh hưởng, kết quả giữ nguyên tuyệt đối.")
    
    print("\n=== TOÀN BỘ KIỂM THỬ BIỂU 0.M3 ĐẠT 100% ===")

if __name__ == '__main__':
    test_0_m3()
