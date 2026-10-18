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
from processors.m0_m4 import process_0_m4
from processors.m0_m3 import process_0_m3
from processors.m0_m2 import process_0_m2
from processors.m0_m1 import process_0_m1
from exporters.excel_exporter import export_0_m4_test

def test_0_m4():
    data_dir = os.path.join(BASE_DIR, 'data')
    outputs_dir = os.path.join(BASE_DIR, 'outputs')
    template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
    
    print("=== BẮT ĐẦU KIỂM THỬ ĐỘC LẬP BIỂU 0.M4 (CƠ SỞ VẬT CHẤT, TBDH LỚP MẪU GIÁO) ===")
    
    # 1. Quét dữ liệu
    scan = scan_data_directory(data_dir)
    valid_files = [f for f in scan['files'] if f['is_valid']]
    assert len(valid_files) == 92, f"Phải đủ 92 xã, thực tế: {len(valid_files)}"
    print(f"Quét thành công {len(valid_files)} xã hợp lệ.")
    
    # 2. Xử lý tổng hợp 0.M4
    t0 = time.time()
    res = process_0_m4(valid_files)
    dur = time.time() - t0
    print(f"Tổng hợp xong 0.M4 trong {dur:.2f} giây!")
    
    agg = res['aggregated_cells']
    
    # 3. Kiểm tra tính logic nội bộ của 0.M4
    print("\n--- [Kiểm tra Dòng 9 & 11: Mầm non 3, 4 tuổi] ---")
    c9 = agg['C9']
    c11 = agg['C11']
    c13 = agg['C13']
    print(f"  + Số trường MN công lập (C9): {c9:,.0f}, Cơ sở ĐL (C11): {c11:,.0f} -> Tổng C13: {c13:,.0f}")
    assert c13 == c9 + c11, "C13 phải bằng C9 + C11"
    
    h13 = agg['H13']
    k13 = agg['K13']
    l13 = agg['L13']
    m13 = agg['M13']
    n13 = agg['N13']
    o13 = agg['O13']
    p13 = agg['P13']
    print(f"  + Lớp 3, 4 tuổi: Số lớp (H13) = {h13:,.0f}, Tổng phòng học (K13) = {k13:,.0f}")
    print(f"  + Tỷ lệ phòng/lớp cấp tỉnh (L13): {l13:.2f} phòng/lớp")
    print(f"  + Chia theo loại: Kiên cố (M13) = {m13:,.0f}, Bán kiên cố (N13) = {n13:,.0f}, Tạm/nhờ (O13) = {o13:,.0f}")
    print(f"  + Lớp đủ TBDH (P13): {p13:,.0f}")
    assert h13 > 0, "Số lớp H13 phải > 0"
    assert k13 > 0, "Số phòng K13 phải > 0"
    assert 0.8 <= l13 <= 1.5, f"Tỷ lệ phòng/lớp L13={l13} ngoài biên độ bình thường (0.8 - 1.5)"
    
    print("\n--- [Kiểm tra Dòng 10 & 12: Mầm non 5 tuổi] ---")
    h14 = agg['H14']
    k14 = agg['K14']
    l14 = agg['L14']
    m14 = agg['M14']
    n14 = agg['N14']
    o14 = agg['O14']
    p14 = agg['P14']
    print(f"  + Lớp 5 tuổi: Số lớp (H14) = {h14:,.0f}, Tổng phòng học (K14) = {k14:,.0f}")
    print(f"  + Tỷ lệ phòng/lớp cấp tỉnh (L14): {l14:.2f} phòng/lớp")
    print(f"  + Chia theo loại: Kiên cố (M14) = {m14:,.0f}, Bán kiên cố (N14) = {n14:,.0f}, Tạm/nhờ (O14) = {o14:,.0f}")
    print(f"  + Lớp đủ TBDH (P14): {p14:,.0f}")
    assert h14 > 0, "Số lớp H14 phải > 0"
    assert k14 > 0, "Số phòng K14 phải > 0"
    assert 0.8 <= l14 <= 1.5, f"Tỷ lệ phòng/lớp L14={l14} ngoài biên độ bình thường (0.8 - 1.5)"

    print("\n--- [Kiểm tra hạ tầng phụ trợ: Vệ sinh, Nước sạch, Bếp ăn, Sân chơi] ---")
    q13, r13 = agg['Q13'], agg['R13']
    s13, t13 = agg['S13'], agg['T13']
    u13, v13 = agg['U13'], agg['V13']
    w13, x13 = agg['W13'], agg['X13']
    print(f"  + Khu vệ sinh 3, 4t: Tổng (Q13) = {q13:,.0f}, Đạt chuẩn (R13) = {r13:,.0f}")
    print(f"  + CT Nước sạch 3, 4t: Tổng (S13) = {s13:,.0f}, Đạt chuẩn (T13) = {t13:,.0f}")
    print(f"  + Bếp ăn 3, 4t: Tổng (U13) = {u13:,.0f}, Đạt chuẩn (V13) = {v13:,.0f}")
    print(f"  + Sân chơi 3, 4t: Sân chơi (W13) = {w13:,.0f}, Sân có đồ chơi (X13) = {x13:,.0f}")
    assert r13 <= q13, "Khu VS đạt chuẩn phải <= Tổng SL"
    assert t13 <= s13, "Nước sạch đạt chuẩn phải <= Tổng SL"
    assert v13 <= u13, "Bếp ăn đạt chuẩn phải <= Tổng SL"
    assert x13 <= w13, "Sân có đồ chơi phải <= Tổng sân chơi"

    # 4. Xuất file Excel kết quả
    print("\n--- [Kiểm tra Xuất file Excel kết quả] ---")
    out_file = export_0_m4_test(res, template_path, outputs_dir)
    assert os.path.exists(out_file), f"File không tồn tại: {out_file}"
    file_size = os.path.getsize(out_file)
    print(f"File xuất thành công: {out_file} ({file_size:,} bytes)")
    assert file_size > 30000, "Kích thước file quá nhỏ"

    # Kiểm tra format trong Excel đã xuất
    wb_out = openpyxl.load_workbook(out_file, data_only=True)
    ws_th = wb_out['TONG_HOP_THU']
    assert ws_th['L13'].number_format == '0.00', f"Định dạng L13 phải là 0.00, thực tế: {ws_th['L13'].number_format}"
    assert ws_th['L14'].number_format == '0.00', f"Định dạng L14 phải là 0.00, thực tế: {ws_th['L14'].number_format}"
    assert ws_th['C10'].value == 'x', "C10 phải là 'x'"
    assert ws_th['C14'].value == 'x', "C14 phải là 'x'"
    assert ws_th['L13'].value == l13, f"Giá trị L13 trong Excel ({ws_th['L13'].value}) phải bằng {l13}"
    assert ws_th['L14'].value == l14, f"Giá trị L14 trong Excel ({ws_th['L14'].value}) phải bằng {l14}"

    # Kiểm tra sheet chi tiết
    assert 'CHI_TIET_THEO_XA' in wb_out.sheetnames, "Phải có sheet CHI_TIET_THEO_XA"
    ws_dt = wb_out['CHI_TIET_THEO_XA']
    assert ws_dt.max_row >= 92 * 4 + 4, f"Sheet chi tiết phải có đủ {92*4} dòng dữ liệu, thực tế: {ws_dt.max_row}"
    wb_out.close()

    # 5. Đối soát liên biểu với 0.M2 và 0.M3
    print("\n--- [Đối soát liên biểu với 0.M2 và 0.M3] ---")
    res_m3 = process_0_m3(valid_files)
    res_m2 = process_0_m2(valid_files)
    
    m3_j8 = res_m3['aggregated_cells']['J8'] # Lớp 3, 4t
    m3_j9 = res_m3['aggregated_cells']['J9'] # Lớp 5t
    m2_g8 = res_m2['aggregated_cells']['G8'] # Lớp 3, 4t
    m2_g9 = res_m2['aggregated_cells']['G9'] # Lớp 5t
    m2_u8 = res_m2['aggregated_cells']['U8'] # Lớp đủ TBDH 3, 4t
    m2_u9 = res_m2['aggregated_cells']['U9'] # Lớp đủ TBDH 5t
    
    print(f"  + Lớp 3, 4 tuổi: 0.M4 H13 = {h13:,.0f} | 0.M3 J8 = {m3_j8:,.0f} | 0.M2 G8 = {m2_g8:,.0f}")
    print(f"  + Lớp 5 tuổi:    0.M4 H14 = {h14:,.0f} | 0.M3 J9 = {m3_j9:,.0f} | 0.M2 G9 = {m2_g9:,.0f}")
    print(f"  + Lớp đủ TBDH:   0.M4 P13 = {p13:,.0f} vs 0.M2 U8 = {m2_u8:,.0f} (3, 4t)")
    print(f"                   0.M4 P14 = {p14:,.0f} vs 0.M2 U9 = {m2_u9:,.0f} (5t)")
    
    # Kiểm tra sai lệch trong biên độ cho phép do số liệu báo cáo giữa các biểu ở cấp xã
    diff_h13_j8 = abs(h13 - m3_j8)
    diff_h14_j9 = abs(h14 - m3_j9)
    print(f"  => Chênh lệch lớp 3,4t giữa 0.M4 và 0.M3: {diff_h13_j8:.0f} lớp ({diff_h13_j8/h13*100:.2f}%)")
    print(f"  => Chênh lệch lớp 5t giữa 0.M4 và 0.M3:   {diff_h14_j9:.0f} lớp ({diff_h14_j9/h14*100:.2f}%)")
    assert diff_h13_j8 / h13 < 0.05, "Chênh lệch số lớp 3,4t giữa các biểu vượt quá 5%"
    assert diff_h14_j9 / h14 < 0.05, "Chênh lệch số lớp 5t giữa các biểu vượt quá 5%"

    # 6. Kiểm tra hồi quy 0.M1
    print("\n--- [Kiểm tra Hồi quy Biểu 0.M1] ---")
    res_m1 = process_0_m1(valid_files)
    assert res_m1['aggregated_cells']['I5'] == 23765.0, "Hồi quy 0.M1 I5 thất bại"
    assert res_m1['aggregated_cells']['K5'] == 110547.0, "Hồi quy 0.M1 K5 thất bại"
    print("  + Hồi quy 0.M1: 100% Khớp chuẩn tuyệt đối (I5 = 23,765; K5 = 110,547)!")

    print("\n=======================================================")
    print(">>> TẤT CẢ KIỂM THỬ BIỂU 0.M4 ĐẠT 100% HOÀN HẢO! <<<")
    print("=======================================================")

if __name__ == '__main__':
    test_0_m4()
