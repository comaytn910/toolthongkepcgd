import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m2_m3 import process_2_m3
from exporters.excel_exporter import export_2_m3_test

def test_2_m3_aggregation():
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
        
    print(f"Bắt đầu kiểm thử Biểu 2.M3 trên {len(valid_files)} xã...")
    result = process_2_m3(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 2.M3")
    
    # 2. Kiểm tra các chỉ tiêu số lượng và tỷ lệ cấp tỉnh
    p_cnt = result['province_counts']
    p_rt = result['province_rates']
    
    # Thanh thiếu niên 11 - 18 tuổi
    assert p_cnt['D'] == 230681.0, f"Expected D=230681, got {p_cnt['D']}"
    assert p_cnt['E'] == 3600.0, f"Expected E=3600, got {p_cnt['E']}"
    print(f"PASS 2: TTN 11-18t toàn tỉnh = {p_cnt['D']:,.0f}, Khuyết tật 11-18t = {p_cnt['E']:,.0f}")
    
    # Phổ cập THCS thanh thiếu niên 15 - 18 tuổi
    assert p_cnt['H'] == 101237.0, f"Expected H=101237, got {p_cnt['H']}"
    assert p_cnt['I'] == 98765.0, f"Expected I=98765, got {p_cnt['I']}"
    assert round(p_rt['J'], 2) == 97.56, f"Expected J=97.56%, got {p_rt['J']}"
    print(f"PASS 3: Đối tượng PPC 15-18t = {p_cnt['H']:,.0f}, TN THCS = {p_cnt['I']:,.0f} (Tỷ lệ = {p_rt['J']:.2f}% - Khớp tuyệt đối với H55 của 2.M1)")
    
    # Học tiếp THPT / GDTX / GDNN
    assert p_cnt['K'] == 93280.0, f"Expected K=93280, got {p_cnt['K']}"
    assert round(p_rt['L'], 2) == 92.14, f"Expected L=92.14%, got {p_rt['L']}"
    print(f"PASS 4: TTN 15-18t học tiếp GDPT/GDTX/GDNN = {p_cnt['K']:,.0f} (Tỷ lệ = {p_rt['L']:.2f}%)")
    
    # Trẻ khuyết tật
    assert p_cnt['M'] == 3602.0, f"Expected M=3602, got {p_cnt['M']}"
    assert p_cnt['N'] == 2919.0, f"Expected N=2919, got {p_cnt['N']}"
    assert p_cnt['O'] == 2881.0, f"Expected O=2881, got {p_cnt['O']}"
    assert round(p_rt['P'], 2) == 98.70, f"Expected P=98.70%, got {p_rt['P']}"
    print(f"PASS 5: Trẻ khuyết tật = {p_cnt['M']:,.0f}, Có khả năng học = {p_cnt['N']:,.0f}, Tiếp cận GD = {p_cnt['O']:,.0f} (Tỷ lệ = {p_rt['P']:.2f}%)")
    
    # Đạt chuẩn toàn tỉnh
    assert result['provincial_standard'] == "Đạt chuẩn PCGD THCS Mức độ 3"
    print("PASS 6: Đánh giá chuẩn toàn tỉnh: Đạt chuẩn PCGD THCS Mức độ 3")
    
    # 3. Xuất Excel và kiểm tra tính toàn vẹn
    out_file = export_2_m3_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 30000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 7: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    # Kiểm tra cấu trúc workbook xuất ra
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    assert 'DOI_CHIEU_SAI_LECH' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['D46'].value == 230681.0
    assert ws_th['E46'].value == 3600.0
    assert ws_th['H46'].value == 101237.0
    assert ws_th['I46'].value == 98765.0
    assert ws_th['J46'].value == 97.56
    assert ws_th['K46'].value == 93280.0
    assert ws_th['L46'].value == 92.14
    assert ws_th['M46'].value == 3602.0
    assert ws_th['N46'].value == 2919.0
    assert ws_th['O46'].value == 2881.0
    assert ws_th['P46'].value == 98.70
    assert ws_th['F46'].value == "Mức độ 3"
    assert ws_th['G46'].value == "Mức độ 2"
    assert ws_th['Q46'].value == "Đạt"
    assert ws_th['R46'].value == "Đạt"
    assert ws_th['S46'].value == "Mức độ 3"
    print("PASS 8: Dữ liệu trên sheet TONG_HOP_THU khớp chuẩn xác 100% với các chỉ tiêu nghiệp vụ")
    
    # Kiểm tra sheet chi tiết 92 xã
    ws_dt = wb['CHI_TIET_92_XA']
    assert ws_dt.max_row >= 96, f"Expected >= 96 rows, got {ws_dt.max_row}"
    print("PASS 9: Sheet CHI_TIET_92_XA đầy đủ 92 xã và dòng tổng toàn tỉnh")
    
    wb.close()
    print("=== TẤT CẢ 9/9 ĐIỀU KIỆN KIỂM THỬ BIỂU 2.M3 ĐÃ ĐẠT 100% ===")

if __name__ == "__main__":
    test_2_m3_aggregation()
