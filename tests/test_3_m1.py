"""
Kiểm thử độc lập Biểu 3.M1: Thống kê đạt chuẩn XMC Mức 1 (15-25 tuổi)
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

import os
import sys
import glob
import time
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from processors.m3_m1 import process_3_m1_single_file, aggregate_3_m1
from exporters.excel_exporter import export_3_m1_test


def test_3_m1():
    print("=" * 60)
    print("BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 3.M1 (CHUẨN XMC MỨC 1 15-25T)")
    print("=" * 60)
    
    start_time = time.time()
    files = sorted(glob.glob('data/*.xlsx'))
    assert len(files) == 92, f"Kỳ vọng 92 files, thực tế: {len(files)}"
    
    # 1. Đọc và xử lý từng xã
    communes_data = []
    for f in files:
        c_name = os.path.basename(f).split(' - ')[0]
        data = process_3_m1_single_file(f, c_name)
        assert data is not None, f"Không đọc được sheet 3.M1 từ {c_name}"
        communes_data.append(data)
        
    assert len(communes_data) == 92, f"Kỳ vọng 92 xã, thực tế: {len(communes_data)}"
    print(f"-> Test 1 PASS: Đọc thành công 92/92 xã.")
    
    # 2. Tổng hợp cấp tỉnh
    result = aggregate_3_m1(communes_data)
    assert result['total_communes'] == 92
    assert result['sum_c_dan_so'] == 244515.0, f"Tổng C dân số sai: {result['sum_c_dan_so']}"
    assert result['sum_d_biet_chu_m1'] == 244352.0, f"Tổng D biết chữ M1 sai: {result['sum_d_biet_chu_m1']}"
    assert result['sum_f_biet_chu_m2'] == 244230.0, f"Tổng F biết chữ M2 sai: {result['sum_f_biet_chu_m2']}"
    
    # Kiểm tra tỷ lệ cấp tỉnh
    expected_rate_m1 = round(244352.0 / 244515.0 * 100, 2)
    expected_rate_m2 = round(244230.0 / 244515.0 * 100, 2)
    assert result['rate_m1_prov'] == expected_rate_m1, f"Tỷ lệ M1 sai: {result['rate_m1_prov']} vs {expected_rate_m1}"
    assert result['rate_m2_prov'] == expected_rate_m2, f"Tỷ lệ M2 sai: {result['rate_m2_prov']} vs {expected_rate_m2}"
    print(f"-> Test 2 PASS: Tổng hợp cấp tỉnh chuẩn xác 100% (Dân số: {result['sum_c_dan_so']:,.0f}, M1: {result['rate_m1_prov']}%, M2: {result['rate_m2_prov']}%).")
    
    # 3. Xuất file Excel
    os.makedirs('outputs', exist_ok=True)
    out_path = 'outputs/TONG_HOP_THU_3.M1_PCGD_2026.xlsx'
    saved_path = export_3_m1_test(result, 'MẪU CHUẨN - PCGD 2026(ok).xlsx', out_path)
    assert os.path.exists(saved_path), f"File không tồn tại: {saved_path}"
    print(f"-> Test 3 PASS: Xuất file Excel thành công tại {saved_path}.")
    
    # 4. Kiểm tra cấu trúc file Excel xuất ra
    wb = openpyxl.load_workbook(saved_path, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames, "Thiếu sheet TONG_HOP_THU"
    assert 'CHI_TIET_92_XA' in wb.sheetnames, "Thiếu sheet CHI_TIET_92_XA"
    
    ws_th = wb['TONG_HOP_THU']
    # Tìm dòng Cộng
    total_val = None
    for r in range(1, ws_th.max_row + 1):
        v1 = str(ws_th.cell(r, 2).value or '')
        if 'cộng toàn tỉnh' in v1.lower():
            total_val = ws_th.cell(r, 3).value
            rate_m1_val = ws_th.cell(r, 5).value
            break
            
    assert total_val == 244515.0, f"Giá trị ô tổng Cột C trong sheet TONG_HOP_THU sai: {total_val}"
    assert rate_m1_val == 99.93, f"Giá trị ô tỷ lệ M1 sai: {rate_m1_val}"
    wb.close()
    print(f"-> Test 4 PASS: File Excel kiểm tra ngược dữ liệu khớp 100%.")
    
    elapsed = time.time() - start_time
    print(f"\nTOÀN BỘ 4/4 TESTS ĐẠT 100% TRONG {elapsed:.2f} GIÂY!")
    print("=" * 60)


if __name__ == '__main__':
    test_3_m1()
