"""
Kiểm thử toàn diện luồng xử lý và xuất Excel cho khối 3.M (3.M1 đến 3.M9)
Hệ thống Tổng hợp số liệu PCGD - XMC tỉnh Thái Nguyên năm 2026
"""

import os
import sys
import glob
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

from processors.m3_m1 import process_3_m1_single_file, aggregate_3_m1
from processors.m3_m2 import process_3_m2_single_file, aggregate_3_m2
from processors.m3_m3 import process_3_m3_single_file, aggregate_3_m3
from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu
from exporters.excel_exporter import (
    export_3_m1_test,
    export_3_m2_test,
    export_3_m3_test,
    export_3_m4_test,
    export_3_m5_test,
    export_3_m6_test,
    export_3_m7_test,
    export_3_m8_test,
    export_3_m9_test,
    export_3_mu_chu_test
)

TEMPLATE_PATH = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
OUTPUTS_DIR = os.path.join(BASE_DIR, 'outputs')
DATA_DIR = os.path.join(BASE_DIR, 'data')

def verify_exported_excel(out_path: str, sheet_code: str, expected_title_part: str, is_standard_type: bool):
    assert os.path.exists(out_path), f"File không tồn tại: {out_path}"
    file_size = os.path.getsize(out_path)
    assert file_size > 5000, f"File quá nhỏ ({file_size} bytes): {out_path}"
    
    wb = openpyxl.load_workbook(out_path, data_only=False)
    sheet_names = wb.sheetnames
    assert 'TONG_HOP_THU' in sheet_names, f"Thiếu sheet TONG_HOP_THU trong {out_path}"
    assert 'CHI_TIET_92_XA' in sheet_names, f"Thiếu sheet CHI_TIET_92_XA trong {out_path}"
    
    ws_th = wb['TONG_HOP_THU']
    a1_val = str(ws_th['A1'].value or '')
    assert sheet_code in a1_val, f"Tiêu đề A1 không chứa mã biểu {sheet_code}: {a1_val}"
    assert expected_title_part.lower() in a1_val.lower(), f"Tiêu đề A1 không chứa '{expected_title_part}': {a1_val}"
    
    # Tìm dòng tổng cộng
    total_r = None
    for r in range(1, ws_th.max_row + 1):
        v1 = str(ws_th.cell(r, 2).value or '').lower()
        v0 = str(ws_th.cell(r, 1).value or '').strip()
        if 'cộng toàn tỉnh' in v1 or v0 == '@':
            total_r = r
            break
            
    assert total_r is not None, f"Không tìm thấy dòng tổng cộng trong {sheet_code}"
    
    # Kiểm tra định dạng %
    if is_standard_type:
        # Cột E và G là tỷ lệ
        fmt_e = ws_th.cell(total_r, 5).number_format
        fmt_g = ws_th.cell(total_r, 7).number_format
        assert '%' in fmt_e, f"Định dạng cột E không có %: {fmt_e}"
        assert '%' in fmt_g, f"Định dạng cột G không có %: {fmt_g}"
    else:
        # Cột H (cột 8) là tỷ lệ mù chữ
        fmt_h = ws_th.cell(total_r, 8).number_format
        assert '%' in fmt_h, f"Định dạng cột H không có %: {fmt_h}"
        
    # Kiểm tra sheet chi tiết
    ws_dt = wb['CHI_TIET_92_XA']
    assert ws_dt.max_row >= 93, f"Sheet chi tiết không đủ 92 xã: max_row={ws_dt.max_row}"
    
    wb.close()
    return file_size


def run_pipeline_test():
    print("=" * 80)
    print("KIỂM THỬ TOÀN DIỆN LUỒNG XỬ LÝ VÀ XUẤT EXCEL KHỐI 3.M (3.M1 ĐẾN 3.M9)")
    print("=" * 80)
    
    files = sorted(glob.glob(os.path.join(DATA_DIR, '*.xlsx')))
    assert len(files) == 92, f"Kỳ vọng 92 files, thực tế: {len(files)}"
    
    test_results = []
    
    # Danh sách cấu hình 9 biểu
    configs = [
        {
            'code': '3.M1',
            'name': 'Chuẩn XMC Mức 1 (15-25t)',
            'is_standard': True,
            'expected_title': 'MỨC ĐỘ 1 (15-25 TUỔI)',
            'proc_single': process_3_m1_single_file,
            'proc_agg': aggregate_3_m1,
            'exporter': export_3_m1_test
        },
        {
            'code': '3.M2',
            'name': 'Chuẩn XMC Mức 2 (15-35t)',
            'is_standard': True,
            'expected_title': 'MỨC ĐỘ 2 (15-35 TUỔI)',
            'proc_single': process_3_m2_single_file,
            'proc_agg': aggregate_3_m2,
            'exporter': export_3_m2_test
        },
        {
            'code': '3.M3',
            'name': 'Chuẩn XMC Chung toàn xã (15-60t)',
            'is_standard': True,
            'expected_title': 'TOÀN XÃ (15-60 TUỔI)',
            'proc_single': process_3_m3_single_file,
            'proc_agg': aggregate_3_m3,
            'exporter': export_3_m3_test
        },
        {
            'code': '3.M4',
            'name': 'Hiện trạng mù chữ M1 (15-25t)',
            'is_standard': False,
            'expected_title': 'MỨC 1 (15-25 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M4'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M4'),
            'exporter': export_3_m4_test
        },
        {
            'code': '3.M5',
            'name': 'Hiện trạng mù chữ M1 (15-35t)',
            'is_standard': False,
            'expected_title': 'MỨC 1 (15-35 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M5'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M5'),
            'exporter': export_3_m5_test
        },
        {
            'code': '3.M6',
            'name': 'Hiện trạng mù chữ M1 (15-60t)',
            'is_standard': False,
            'expected_title': 'MỨC 1 (15-60 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M6'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M6'),
            'exporter': export_3_m6_test
        },
        {
            'code': '3.M7',
            'name': 'Hiện trạng mù chữ M2 (15-25t)',
            'is_standard': False,
            'expected_title': 'MỨC 2 (15-25 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M7'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M7'),
            'exporter': export_3_m7_test
        },
        {
            'code': '3.M8',
            'name': 'Hiện trạng mù chữ M2 (15-35t)',
            'is_standard': False,
            'expected_title': 'MỨC 2 (15-35 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M8'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M8'),
            'exporter': export_3_m8_test
        },
        {
            'code': '3.M9',
            'name': 'Hiện trạng mù chữ M2 (15-60t)',
            'is_standard': False,
            'expected_title': 'MỨC 2 (15-60 TUỔI)',
            'proc_single': lambda f, c: process_3_mu_chu_single_file(f, c, '3.M9'),
            'proc_agg': lambda d: aggregate_3_mu_chu(d, '3.M9'),
            'exporter': export_3_m9_test
        },
    ]
    
    t_total_start = time.time()
    
    for cfg in configs:
        sc = cfg['code']
        print(f"\n--- Đang kiểm thử biểu: {sc} - {cfg['name']} ---")
        t0 = time.time()
        
        # 1. Đọc 92 xã
        c_data = []
        for f in files:
            c_name = os.path.basename(f).split(' - ')[0]
            d = cfg['proc_single'](f, c_name)
            assert d is not None, f"Không đọc được {sc} từ {c_name}"
            c_data.append(d)
            
        t_read = time.time() - t0
        print(f"  + Đọc 92 xã: {t_read:.2f}s (đủ {len(c_data)} xã)")
        
        # 2. Tổng hợp cấp tỉnh
        t1 = time.time()
        res = cfg['proc_agg'](c_data)
        t_agg = time.time() - t1
        
        if cfg['is_standard']:
            info_str = f"DS={res['sum_c_dan_so']:,.0f}, M1={res['sum_d_biet_chu_m1']:,.0f} ({res['rate_m1_prov']:.2f}%), M2={res['sum_f_biet_chu_m2']:,.0f} ({res['rate_m2_prov']:.2f}%), Đạt M1={res['count_dat_m1']}/92, Đạt M2={res['count_dat_m2']}/92"
        else:
            info_str = f"DS={res['sum_c_dan_so']:,.0f}, Mù chữ={res['sum_g_mu_chu']:,.0f} ({res['rate_h_prov']:.2f}%), Nữ={res['sum_i_mu_chu_nu']:,.0f} ({res['rate_j_prov']:.2f}%), DT={res['sum_k_mu_chu_dt']:,.0f} ({res['rate_l_prov']:.2f}%)"
            
        print(f"  + Tổng hợp cấp tỉnh: {t_agg:.2f}s -> {info_str}")
        
        # 3. Xuất file Excel
        t2 = time.time()
        out_file = os.path.join(OUTPUTS_DIR, f"TONG_HOP_THU_{sc}_PCGD_2026.xlsx")
        
        # Xóa file cũ nếu có để đảm bảo file mới tạo
        if os.path.exists(out_file):
            try:
                os.remove(out_file)
            except Exception:
                pass
                
        out_saved = cfg['exporter'](res, TEMPLATE_PATH, out_file)
        t_export = time.time() - t2
        print(f"  + Xuất file Excel: {t_export:.2f}s -> {os.path.basename(out_saved)}")
        
        # 4. Đọc ngược file vừa xuất để xác minh
        t3 = time.time()
        f_size = verify_exported_excel(out_saved, sc, cfg['expected_title'], cfg['is_standard'])
        t_verify = time.time() - t3
        print(f"  + Mở đọc xác minh: {t_verify:.2f}s -> OK (Kích thước: {f_size:,} bytes, Đủ sheets & đúng định dạng %)")
        
        test_results.append({
            'code': sc,
            'name': cfg['name'],
            'status': 'PASS',
            'time_total': round(t_read + t_agg + t_export + t_verify, 2),
            'file_size': f_size,
            'summary': info_str
        })

    # Kiểm thử bổ sung: Gọi trực tiếp export_3_mu_chu_test cho 3.M4 để xác nhận
    print("\n--- Kiểm thử bổ sung: Lời gọi trực tiếp export_3_mu_chu_test ---")
    c_data_test = [process_3_mu_chu_single_file(f, os.path.basename(f).split(' - ')[0], '3.M4') for f in files]
    res_test = aggregate_3_mu_chu(c_data_test, '3.M4')
    out_direct = os.path.join(OUTPUTS_DIR, "TONG_HOP_THU_3.M4_DIRECT_TEST.xlsx")
    saved_direct = export_3_mu_chu_test(res_test, TEMPLATE_PATH, out_direct)
    assert os.path.exists(saved_direct), "File export_3_mu_chu_test trực tiếp không tồn tại!"
    print(f"  + export_3_mu_chu_test trực tiếp THÀNH CÔNG: {saved_direct}")
    try:
        os.remove(saved_direct)
    except Exception:
        pass

    t_all = time.time() - t_total_start
    print("\n" + "=" * 80)
    print(f"KẾT QUẢ KIỂM THỬ KHỐI 3.M: 9/9 BIỂU ĐẠT CHUẨN 100% TRONG {t_all:.2f} GIÂY")
    print("=" * 80)
    for r in test_results:
        print(f"[{r['status']}] {r['code']:<5} | {r['name']:<35} | {r['time_total']}s | {r['summary']}")
    print("=" * 80)
    return test_results

if __name__ == '__main__':
    run_pipeline_test()
