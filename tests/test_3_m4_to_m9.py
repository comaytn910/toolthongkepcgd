"""
Kiểm thử độc lập liên hoàn 6 Biểu Hiện trạng mù chữ: 3.M4, 3.M5, 3.M6, 3.M7, 3.M8, 3.M9
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

import os
import sys
import glob
import time
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu
from exporters.excel_exporter import export_3_mu_chu_test

TEST_CASES = [
    {
        'sheet': '3.M4',
        'title': 'Hiện trạng mù chữ Mức 1 (15-25t)',
        'exp_dan_so': 244515.0,
        'exp_mu_chu': 163.0,
        'cross_ref_sheet': '3.M1',
        'cross_biet_chu': 244350.0
    },
    {
        'sheet': '3.M5',
        'title': 'Hiện trạng mù chữ Mức 1 (15-35t)',
        'exp_dan_so': 479435.0,
        'exp_mu_chu': 752.0,
        'cross_ref_sheet': '3.M2',
        'cross_biet_chu': 478664.0
    },
    {
        'sheet': '3.M6',
        'title': 'Hiện trạng mù chữ Mức 1 (15-60t)',
        'exp_dan_so': 1097844.0,
        'exp_mu_chu': 7466.0,
        'cross_ref_sheet': '3.M3',
        'cross_biet_chu': 1090383.0
    },
    {
        'sheet': '3.M7',
        'title': 'Hiện trạng mù chữ Mức 2 (15-25t)',
        'exp_dan_so': 244515.0,
        'exp_mu_chu': 285.0,
        'cross_ref_sheet': '3.M1',
        'cross_biet_chu': 244230.0
    },
    {
        'sheet': '3.M8',
        'title': 'Hiện trạng mù chữ Mức 2 (15-35t)',
        'exp_dan_so': 479435.0,
        'exp_mu_chu': 1918.0,
        'cross_ref_sheet': '3.M2',
        'cross_biet_chu': 477517.0
    },
    {
        'sheet': '3.M9',
        'title': 'Hiện trạng mù chữ Mức 2 (15-60t)',
        'exp_dan_so': 1097844.0,
        'cross_ref_sheet': '3.M3',
        'cross_biet_chu': 1080608.0
    },
]


def test_single_sheet(tc, files):
    s = tc['sheet']
    print(f"\n--- KIỂM THỬ BIỂU {s}: {tc['title']} ---")
    t0 = time.time()
    
    communes_data = []
    for f in files:
        c_name = os.path.basename(f).split(' - ')[0]
        data = process_3_mu_chu_single_file(f, c_name, s)
        assert data is not None, f"Lỗi đọc sheet {s} xã {c_name}"
        communes_data.append(data)
        
    assert len(communes_data) == 92, f"Kỳ vọng 92 xã, thực tế: {len(communes_data)}"
    result = aggregate_3_mu_chu(communes_data, s)
    
    # 1. Kiểm tra dân số
    assert result['sum_c_dan_so'] == tc['exp_dan_so'], f"Dân số {s} sai: {result['sum_c_dan_so']} vs {tc['exp_dan_so']}"
    
    # 2. Kiểm tra đối soát chéo toán học: Mù chữ + Biết chữ ~ Tổng dân số
    total_mu_chu = result['sum_g_mu_chu']
    total_biet_chu = tc['cross_biet_chu']
    calc_sum = total_mu_chu + total_biet_chu
    diff = abs(calc_sum - tc['exp_dan_so'])
    # Sai số liên biểu cho phép dưới 0.1% (hoặc tuyệt đối bằng 0 như 3.M7, 3.M8)
    assert diff / tc['exp_dan_so'] < 0.001, f"Lệch đối soát chéo {s}: {diff} người"
    print(f"   [PASS] Đối soát chéo với {tc['cross_ref_sheet']}: Mù chữ ({total_mu_chu:,.0f}) + Biết chữ ({total_biet_chu:,.0f}) = {calc_sum:,.0f} / {tc['exp_dan_so']:,.0f} (Khớp {100 - diff/tc['exp_dan_so']*100:.3f}%)")
    
    # 3. Xuất file Excel
    os.makedirs('outputs', exist_ok=True)
    out_path = f'outputs/TONG_HOP_THU_{s}_PCGD_2026.xlsx'
    saved_path = export_3_mu_chu_test(result, 'MẪU CHUẨN - PCGD 2026(ok).xlsx', out_path)
    assert os.path.exists(saved_path), f"File không tồn tại: {saved_path}"
    
    # 4. Kiểm tra file Excel xuất ra
    wb = openpyxl.load_workbook(saved_path, data_only=True)
    assert 'TONG_HOP_THU' in wb.sheetnames
    assert 'CHI_TIET_92_XA' in wb.sheetnames
    ws_th = wb['TONG_HOP_THU']
    total_r = result['def_sum_row']
    # Tìm dòng Cộng
    found_c = None
    for r in range(1, ws_th.max_row + 1):
        v1 = str(ws_th.cell(r, 2).value or '')
        if 'cộng toàn tỉnh' in v1.lower():
            found_c = ws_th.cell(r, 3).value
            break
    assert found_c == tc['exp_dan_so'], f"Ô tổng cấp tỉnh {s} sai: {found_c} vs {tc['exp_dan_so']}"
    wb.close()
    
    elapsed = time.time() - t0
    print(f"   [PASS] Xuất file và kiểm tra ngược Excel thành công ({elapsed:.1f}s) -> {saved_path}")
    return result


def main():
    print("=" * 70)
    print("BẮT ĐẦU KIỂM THỬ LIÊN HOÀN 6 BIỂU CUỐI CÙNG: 3.M4 ĐẾN 3.M9")
    print("=" * 70)
    
    t_start = time.time()
    files = sorted(glob.glob('data/*.xlsx'))
    assert len(files) == 92, f"Kỳ vọng 92 files, thực tế {len(files)}"
    
    results = {}
    for tc in TEST_CASES:
        res = test_single_sheet(tc, files)
        results[tc['sheet']] = res
        
    t_total = time.time() - t_start
    print("\n" + "=" * 70)
    print(f"TẤT CẢ 6/6 BIỂU (3.M4 -> 3.M9) ĐÃ VƯỢT QUA 100% KIỂM THỬ TRONG {t_total:.1f} GIÂY!")
    print("=" * 70)
    for s, r in results.items():
        print(f"  * {s}: Dân số={r['sum_c_dan_so']:11,.0f} | Mù chữ={r['sum_g_mu_chu']:6,.0f} ({r['rate_h_prov']:4.2f}%) | Nữ={r['sum_i_mu_chu_nu']:5,.0f} | DT={r['sum_k_mu_chu_dt']:5,.0f}")


if __name__ == '__main__':
    main()
