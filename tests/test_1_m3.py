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
from processors.m1_m3 import process_1_m3
from exporters.excel_exporter import export_1_m3_test

def test_1_m3_pipeline():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 1.M3 (CSVC CẤP TIỂU HỌC) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ.")
    
    # 2. Xử lý tổng hợp 1.M3
    res_m3 = process_1_m3(valid_files)
    agg = res_m3['aggregated_cells']
    
    # 3. Kiểm tra các chỉ tiêu cốt lõi
    tot_diem = agg['C8']
    tot_lop = agg['D8']
    tot_lop_ghep = agg['E8']
    phong_kc = agg['F8']
    phong_bkc = agg['G8']
    phong_tam = agg['H8']
    phong_muon = agg['I8']
    tot_phong = res_m3['total_rooms']
    tl_ph_lop = agg['J8']
    
    print(f"[ĐỐI SOÁT CỐT LÕI]")
    print(f"  - Tổng số Điểm trường (C8): {tot_diem:,.0f}")
    print(f"  - Tổng số Lớp học Tiểu học (D8): {tot_lop:,.0f} lớp")
    print(f"  - Số lớp ghép (E8): {tot_lop_ghep:,.0f} lớp")
    print(f"  - Phòng học kiên cố (F8): {phong_kc:,.0f}")
    print(f"  - Phòng học bán kiên cố (G8): {phong_bkc:,.0f}")
    print(f"  - Phòng tạm (H8): {phong_tam:,.0f} | Thuê/mượn (I8): {phong_muon:,.0f}")
    print(f"  - Tổng số phòng học toàn tỉnh: {tot_phong:,.0f} phòng")
    print(f"  - Tỉ lệ Phòng / Lớp toàn tỉnh (J8): {tl_ph_lop:.2f}%")
    
    # Kiểm tra tính logic
    assert tot_diem == 360, f"Tổng điểm trường ({tot_diem}) phải bằng 360"
    assert tot_lop == 4984, f"Tổng số lớp ({tot_lop}) phải bằng 4,984"
    assert tot_phong == 5079, f"Tổng số phòng ({tot_phong}) phải bằng 5,079"
    assert abs(tl_ph_lop - 101.91) < 0.05, f"Tỉ lệ phòng/lớp ({tl_ph_lop}) phải xấp xỉ 101.91%"
    print("  -> [ĐẠT] Toàn bộ chỉ tiêu Điểm trường, Số lớp, Phòng học và Tỉ lệ khớp chuẩn 100%!")
    
    # Kiểm tra xử lý cộng trùng lặp
    assert len(res_m3['discrepancies']) == 1, "Phải phát hiện đúng 1 xã có chèn dòng cộng trùng lặp"
    disc = res_m3['discrepancies'][0]
    assert 'Tích Lương' in disc['commune_name']
    assert disc['d8_raw'] == 262.0 and disc['d_clean'] == 131.0
    print(f"  -> [ĐẠT] Phát hiện và chuẩn hóa chính xác lỗi cộng trùng Phường Tích Lương (D8: 262 -> 131 lớp)!")
    
    # 4. Kiểm tra các phòng chức năng và diện tích
    print(f"\n[PHÒNG CHỨC NĂNG & CÔNG TRÌNH PHỤ TRỢ]")
    print(f"  - Phòng Hiệu trưởng (K8): {agg['K8']:,.0f} | Phó Hiệu trưởng (L8): {agg['L8']:,.0f}")
    print(f"  - Phòng họp (P8): {agg['P8']:,.0f} phòng ({agg['Q8']:,.1f} m2)")
    print(f"  - Thư viện (R8): {agg['R8']:,.0f} phòng ({agg['S8']:,.1f} m2)")
    print(f"  - Thiết bị (T8): {agg['T8']:,.0f} phòng ({agg['U8']:,.1f} m2)")
    print(f"  - Sân chơi (Z8): {agg['Z8']:,.0f} sân ({agg['AA8']:,.1f} m2)")
    print(f"  - Sân tập (AB8): {agg['AB8']:,.0f} sân ({agg['AC8']:,.1f} m2)")
    
    # 5. Kiểm tra xuất file Excel kết quả
    out_file = export_1_m3_test(res_m3, template_path, outputs_dir)
    assert os.path.exists(out_file), f"File không tồn tại: {out_file}"
    file_size = os.path.getsize(out_file)
    print(f"\n[XUẤT EXCEL THÀNH CÔNG] File: {out_file} (Kích thước: {file_size:,} bytes)")
    
    # Mở lại kiểm tra định dạng và dữ liệu
    wb_chk = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb_chk.sheetnames
    assert 'CHI_TIET_THEO_XA' in wb_chk.sheetnames
    assert 'DOI_CHIEU_D_NGUON' in wb_chk.sheetnames
    ws_chk = wb_chk['TONG_HOP_THU']
    assert ws_chk['D8'].value == 4984.0
    assert ws_chk['J8'].value == 101.91
    wb_chk.close()
    print("  -> [ĐẠT] File Excel xuất ra đọc lại hoàn hảo, đầy đủ 3 sheet!")
    
    print("\n>>> TẤT CẢ KIỂM THỬ BIỂU 1.M3 ĐÃ ĐẠT 100%! <<<")

if __name__ == '__main__':
    test_1_m3_pipeline()
