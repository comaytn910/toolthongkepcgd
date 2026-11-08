import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m2_m4 import process_2_m4
from exporters.excel_exporter import export_2_m4_test

def test_2_m4_aggregation():
    data_dir = Path("data")
    files = sorted([f for f in data_dir.glob("*.xlsx") if not f.name.startswith("~$")])
    
    valid_files = []
    for idx, f in enumerate(files, 1):
        valid_files.append({
            'filepath': str(f),
            'filename': f.name,
            'commune_name': f.stem,
            'commune_stt': idx
        })
        
    print(f"Bắt đầu kiểm thử Biểu 2.M4 trên {len(valid_files)} xã...")
    result = process_2_m4(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 2.M4")
    
    # 2. Kiểm tra số lượng trường THCS/liên cấp phát hiện
    assert result['total_schools'] == 158, f"Expected 158 schools, got {result['total_schools']}"
    print(f"PASS 2: Phát hiện và trích xuất chi tiết {result['total_schools']} trường cấp THCS toàn tỉnh")
    
    # 3. Kiểm tra các chỉ tiêu phòng học và tỷ lệ phòng/lớp
    p_cnt = result['province_counts']
    assert p_cnt['C'] == 125.0, f"Expected C=125, got {p_cnt['C']}"
    assert p_cnt['D'] == 3332.0, f"Expected D=3332, got {p_cnt['D']}"
    assert p_cnt['E'] == 3159.0, f"Expected E=3159, got {p_cnt['E']}"
    assert p_cnt['F'] == 85.0, f"Expected F=85, got {p_cnt['F']}"
    assert p_cnt['G'] == 3.0, f"Expected G=3, got {p_cnt['G']}"
    assert result['province_tot_rooms'] == 3247.0, f"Expected tot_rooms=3247, got {result['province_tot_rooms']}"
    assert round(result['province_rate_h'], 2) == 97.45, f"Expected H=97.45%, got {result['province_rate_h']}"
    print(f"PASS 3: Lớp học THCS = {p_cnt['D']:,.0f}, Phòng học = {result['province_tot_rooms']:,.0f} (KC: {p_cnt['E']:,.0f}, Bán KC: {p_cnt['F']:,.0f}, Tạm: {p_cnt['G']:,.0f}), Tỉ lệ Ph/Lớp = {result['province_rate_h']:.2f}%")
    
    # 4. Kiểm tra phòng chức năng và diện tích
    assert p_cnt['I'] == 168.0, f"Expected I=168, got {p_cnt['I']}"
    assert p_cnt['J'] == 302.0, f"Expected J=302, got {p_cnt['J']}"
    assert p_cnt['K'] == 227.0, f"Expected K=227, got {p_cnt['K']}"
    assert p_cnt['L'] == 276.0, f"Expected L=276, got {p_cnt['L']}"
    assert p_cnt['M'] == 273.0, f"Expected M=273, got {p_cnt['M']}"
    assert p_cnt['N'] == 299.0, f"Expected N=299, got {p_cnt['N']}"
    assert p_cnt['O'] == 521.0, f"Expected O=521, got {p_cnt['O']}"
    assert p_cnt['P'] == 27768.0, f"Expected P=27768, got {p_cnt['P']}"
    assert p_cnt['U'] == 305.0, f"Expected U=305, got {p_cnt['U']}"
    assert p_cnt['V'] == 1269714.0, f"Expected V=1269714, got {p_cnt['V']}"
    assert p_cnt['W'] == 277.0, f"Expected W=277, got {p_cnt['W']}"
    assert p_cnt['X'] == 406400.0, f"Expected X=406400, got {p_cnt['X']}"
    print("PASS 4: Các chỉ tiêu phòng chức năng, phòng thí nghiệm, sân chơi, bãi tập khớp chuẩn xác 100%")
    
    # 5. Xuất Excel và kiểm tra cấu trúc
    out_file = export_2_m4_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 30000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 5: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    assert 'CHI_TIET_TRUONG' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['B12'].value == "TỔNG CỘNG TOÀN TỈNH"
    assert ws_th['C12'].value == 125.0
    assert ws_th['D12'].value == 3332.0
    assert ws_th['E12'].value == 3159.0
    assert ws_th['F12'].value == 85.0
    assert ws_th['G12'].value == 3.0
    assert ws_th['H12'].value == 97.45
    print("PASS 6: Sheet TONG_HOP_THU hiển thị chuẩn xác Hàng 12 dòng tổng cộng toàn tỉnh")
    
    ws_dt = wb['CHI_TIET_92_XA']
    assert ws_dt.max_row >= 96
    print("PASS 7: Sheet CHI_TIET_92_XA đầy đủ 92 xã và dòng tổng toàn tỉnh")
    
    ws_sch = wb['CHI_TIET_TRUONG']
    assert ws_sch.max_row >= 162
    print("PASS 8: Sheet CHI_TIET_TRUONG đầy đủ 158 trường cấp THCS toàn tỉnh")
    
    wb.close()
    print("=== TẤT CẢ 8/8 ĐIỀU KIỆN KIỂM THỬ BIỂU 2.M4 ĐÃ ĐẠT 100% ===")

if __name__ == "__main__":
    test_2_m4_aggregation()
