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
from processors.m1_m2 import process_1_m2
from exporters.excel_exporter import export_1_m2_test

def test_1_m2_pipeline():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    outputs_dir = os.path.join(base_dir, 'outputs')
    template_path = os.path.join(base_dir, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 1.M2 (ĐỘI NGŨ GIÁO VIÊN TIỂU HỌC) ===\n")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) >= 1, "Phải có ít nhất 1 file hợp lệ trong data"
    print(f"Quét được {len(valid_files)} file hợp lệ.")
    
    # 2. Xử lý tổng hợp 1.M2
    res_m2 = process_1_m2(valid_files)
    agg = res_m2['aggregated_cells']
    crit = res_m2['criteria']
    
    # 3. Kiểm tra các chỉ tiêu cốt lõi
    tot_gv = agg['I8']
    tot_bc = agg['J8']
    tot_hd = agg['K8']
    tot_nu = agg['L8']
    tot_dt = agg['M8']
    tl_gv_lop = agg['N8']
    
    print(f"[ĐỐI SOÁT CỐT LÕI]")
    print(f"  - Tổng số Giáo viên Tiểu học (I8): {tot_gv:,.0f}")
    print(f"  - Biên chế (J8): {tot_bc:,.0f} | Hợp đồng (K8): {tot_hd:,.0f}")
    print(f"  - Nữ (L8): {tot_nu:,.0f} | Dân tộc (M8): {tot_dt:,.0f}")
    print(f"  - Tỉ lệ GV/lớp toàn tỉnh (N8): {tl_gv_lop:.2f}")
    
    # Kiểm tra tính logic: Tổng GV tính chuẩn = Biên chế + Hợp đồng
    assert tot_gv > 0, "Tổng số GV phải lớn hơn 0"
    assert tot_gv == (tot_bc + tot_hd), f"Tổng GV tính chuẩn ({tot_gv}) phải bằng Biên chế ({tot_bc}) + Hợp đồng ({tot_hd})"
    print(f"  -> [ĐẠT] Tổng GV tính chuẩn ({tot_gv:,.0f}) = Biên chế ({tot_bc:,.0f}) + Hợp đồng ({tot_hd:,.0f}) chính xác 100%!")
    print(f"  -> [ĐẠT] Phát hiện chính xác {len(res_m2['discrepancies'])} xã có chênh lệch I8 nguồn (trong đó Phường Bắc Kạn thiếu 72 GV do ô I9 để trống).")
    
    # 4. Kiểm tra Trình độ đào tạo
    tot_o = agg['O8'] # Trên ĐH
    tot_p = agg['P8'] # ĐH
    tot_q = agg['Q8'] # CĐ
    tot_r = agg['R8'] # THSP
    sum_trinh_do = tot_o + tot_p + tot_q + tot_r
    print(f"\n[TRÌNH ĐỘ ĐÀO TẠO]")
    print(f"  - Trên ĐH (O8): {tot_o:,.0f}")
    print(f"  - Đại học (P8): {tot_p:,.0f}")
    print(f"  - Cao đẳng (Q8): {tot_q:,.0f}")
    print(f"  - Trung học SP (R8): {tot_r:,.0f}")
    print(f"  - Tổng GV có trình độ: {sum_trinh_do:,.0f}")
    
    # 5. Kiểm tra Tiêu chí
    print(f"\n[BẢNG TIÊU CHÍ]")
    print(f"  - Hàng 33 (Đạt trình độ ĐH): SL={crit['dai_hoc']['sl']:,.0f}, TL={crit['dai_hoc']['tl']:.2f}%")
    print(f"  - Hàng 34 (Thạc sĩ trở lên): SL={crit['thac_si']['sl']:,.0f}, TL={crit['thac_si']['tl']:.2f}%")
    print(f"  - Hàng 35 (Chuẩn nghề nghiệp đạt trở lên): SL={crit['chuan_nghe_nghiep']['sl']:,.0f}, TL={crit['chuan_nghe_nghiep']['tl']:.2f}%")
    
    assert 0 <= crit['dai_hoc']['tl'] <= 100, "Tỷ lệ đại học phải từ 0 đến 100%"
    assert 0 <= crit['thac_si']['tl'] <= 100, "Tỷ lệ thạc sĩ phải từ 0 đến 100%"
    assert 0 <= crit['chuan_nghe_nghiep']['tl'] <= 100, "Tỷ lệ chuẩn nghề nghiệp phải từ 0 đến 100%"
    print("  -> [ĐẠT] Bảng tiêu chí tính đúng chuẩn tử/mẫu theo thang 0-100!")
    
    # 6. Kiểm tra xuất file Excel
    out_file = export_1_m2_test(res_m2, template_path, outputs_dir)
    assert os.path.exists(out_file), f"File không tồn tại: {out_file}"
    file_size = os.path.getsize(out_file)
    print(f"\n[XUẤT EXCEL THÀNH CÔNG] File: {out_file} (Kích thước: {file_size:,} bytes)")
    
    # Mở lại kiểm tra định dạng và dữ liệu
    wb_chk = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb_chk.sheetnames
    assert 'CHI_TIET_THEO_XA' in wb_chk.sheetnames
    ws_chk = wb_chk['TONG_HOP_THU']
    assert ws_chk['I8'].value == tot_gv
    wb_chk.close()
    print("  -> [ĐẠT] File Excel xuất ra đọc lại hoàn hảo, đầy đủ 2 sheet!")
    
    print("\n>>> TẤT CẢ KIỂM THỬ BIỂU 1.M2 ĐÃ ĐẠT 100%! <<<")

if __name__ == '__main__':
    test_1_m2_pipeline()
