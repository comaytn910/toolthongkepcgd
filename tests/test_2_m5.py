import os
import sys
from pathlib import Path
import openpyxl

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent.parent))

from processors.m2_m5 import process_2_m5
from exporters.excel_exporter import export_2_m5_test

def test_2_m5_aggregation():
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
        
    print(f"Bắt đầu kiểm thử Biểu 2.M5 trên {len(valid_files)} xã...")
    result = process_2_m5(valid_files)
    
    # 1. Kiểm tra số lượng xã
    assert result['total_communes_count'] == 92, f"Expected 92 communes, got {result['total_communes_count']}"
    print("PASS 1: Đủ 92/92 xã có dữ liệu Biểu 2.M5")
    
    # 2. Kiểm tra số lượng trường THCS phát hiện
    assert result['total_schools'] == 157, f"Expected 157 schools, got {result['total_schools']}"
    print(f"PASS 2: Phát hiện và trích xuất chi tiết {result['total_schools']} trường cấp THCS toàn tỉnh")
    
    # 3. Kiểm tra các chỉ tiêu CBQL và Giáo viên
    p_cnt = result['province_counts']
    assert p_cnt['F'] == 147.0, f"Expected F=147, got {p_cnt['F']}"
    assert p_cnt['G'] == 297.0, f"Expected G=297, got {p_cnt['G']}"
    assert p_cnt['H'] == 6379.0, f"Expected H=6379, got {p_cnt['H']}"
    assert p_cnt['I'] == 5340.0, f"Expected I=5340, got {p_cnt['I']}"
    assert p_cnt['J'] == 1008.0, f"Expected J=1008, got {p_cnt['J']}"
    assert p_cnt['K'] == 5059.0, f"Expected K=5059, got {p_cnt['K']}"
    assert p_cnt['L'] == 2555.0, f"Expected L=2555, got {p_cnt['L']}"
    assert result['prov_ratio_gv_class'] == 1.91, f"Expected ratio=1.91, got {result['prov_ratio_gv_class']}"
    print(f"PASS 3: Hiệu trưởng = {p_cnt['F']:,.0f}, P.Hiệu trưởng = {p_cnt['G']:,.0f}, Tổng GV = {p_cnt['H']:,.0f} (BC: {p_cnt['I']:,.0f}, HĐ: {p_cnt['J']:,.0f}, Nữ: {p_cnt['K']:,.0f}), Tỉ lệ GV/lớp = {result['prov_ratio_gv_class']:.2f}")
    
    # 4. Kiểm tra trình độ đào tạo và Bảng tiêu chí
    assert p_cnt['N'] == 526.0, f"Expected N=526, got {p_cnt['N']}"
    assert p_cnt['O'] == 5833.0, f"Expected O=5833, got {p_cnt['O']}"
    assert p_cnt['P'] == 28.0, f"Expected P=28, got {p_cnt['P']}"
    assert p_cnt['Q'] == 12.0, f"Expected Q=12, got {p_cnt['Q']}"
    
    crit = result['criteria']
    assert crit['dat_chuan']['sl'] == 6359.0
    assert crit['dat_chuan']['tl'] == 99.69
    assert crit['tren_chuan']['sl'] == 526.0
    assert crit['tren_chuan']['tl'] == 8.25
    assert crit['chuan_nghe_nghiep']['sl'] == 6183.0
    assert crit['chuan_nghe_nghiep']['tl'] == 96.93
    print(f"PASS 4: Bảng tiêu chí chuẩn xác: Đạt chuẩn = {crit['dat_chuan']['sl']:,.0f} ({crit['dat_chuan']['tl']:.2f}%), Trên chuẩn = {crit['tren_chuan']['sl']:,.0f} ({crit['tren_chuan']['tl']:.2f}%), Chuẩn NN = {crit['chuan_nghe_nghiep']['sl']:,.0f} ({crit['chuan_nghe_nghiep']['tl']:.2f}%)")
    
    # 5. Kiểm tra phát hiện 8 xã sai lệch nguồn
    assert len(result['discrepancies']) == 8, f"Expected 8 discrepancies, got {len(result['discrepancies'])}"
    comm_disc = [d['commune'] for d in result['discrepancies']]
    assert any("Bạch Thông" in c for c in comm_disc)
    assert any("Phủ Thông" in c for c in comm_disc)
    assert any("Trung Thành" in c for c in comm_disc)
    assert any("Phổ Yên" in c for c in comm_disc)
    print(f"PASS 5: Phát hiện và giải trình chuẩn xác 8 xã sai lệch nguồn giữa Tổng số GV và Biên chế + Hợp đồng")
    
    # 6. Xuất Excel và kiểm tra cấu trúc
    out_file = export_2_m5_test(result)
    assert os.path.exists(out_file), f"Output file does not exist: {out_file}"
    fsize = os.path.getsize(out_file)
    assert fsize > 50000, f"Output file size too small: {fsize} bytes"
    print(f"PASS 6: Xuất Excel thành công -> {out_file} ({fsize/1024:.1f} KB)")
    
    wb = openpyxl.load_workbook(out_file, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    assert 'CHI_TIET_TRUONG' in wb.sheetnames
    assert 'DOI_CHIEU_SAI_LECH' in wb.sheetnames
    
    ws_th = wb['TONG_HOP_THU']
    assert ws_th['B12'].value == "TỔNG CỘNG TOÀN TỈNH"
    assert ws_th['F12'].value == 147.0
    assert ws_th['G12'].value == 297.0
    assert ws_th['H12'].value == 6379.0
    assert ws_th['I12'].value == 5340.0
    assert ws_th['J12'].value == 1008.0
    assert ws_th['M12'].value == 1.91
    assert ws_th['I15'].value == 6359.0
    assert ws_th['J15'].value == 99.69
    assert ws_th['I16'].value == 526.0
    assert ws_th['J16'].value == 8.25
    assert ws_th['I17'].value == 6183.0
    assert ws_th['J17'].value == 96.93
    print("PASS 7: Sheet TONG_HOP_THU hiển thị chuẩn xác Hàng 12 dòng tổng cộng toàn tỉnh và Bảng tiêu chí Hàng 15-17")
    
    ws_dt = wb['CHI_TIET_92_XA']
    assert ws_dt.max_row >= 96
    print("PASS 8: Sheet CHI_TIET_92_XA đầy đủ 92 xã và dòng tổng toàn tỉnh")
    
    ws_sch = wb['CHI_TIET_TRUONG']
    assert ws_sch.max_row >= 161
    print("PASS 9: Sheet CHI_TIET_TRUONG đầy đủ 157 trường cấp THCS toàn tỉnh")
    
    wb.close()
    print("=== TẤT CẢ 9/9 ĐIỀU KIỆN KIỂM THỬ BIỂU 2.M5 ĐÃ ĐẠT 100% ===")

if __name__ == "__main__":
    test_2_m5_aggregation()
