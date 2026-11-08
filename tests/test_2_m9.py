import os
import sys
import unittest
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
from processors.m2_m9 import process_2_m9
from exporters.excel_exporter import export_2_m9_test

class TestProcess2M9(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_2_m9(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_province_counts(self):
        """Kiểm tra số lượng các chỉ tiêu cấp tỉnh"""
        p_cnt = self.result['province_counts']
        
        # C: Đối tượng PC 15-21
        self.assertEqual(p_cnt['C'], 166166.0)
        # D: Tổng số phải PC 15-18
        self.assertEqual(p_cnt['D'], 100561.0)
        # E: Tổng số TN THCS 15-18
        self.assertEqual(p_cnt['E'], 98111.0)
        # G, H: Còn học, Bỏ học THCS
        self.assertEqual(p_cnt['G'], 807.0)
        self.assertEqual(p_cnt['H'], 2270.0)
        
        # I, J: Lớp 9, TN THCS năm qua
        self.assertEqual(p_cnt['I'], 26233.0)
        self.assertEqual(p_cnt['J'], 26148.0)
        
        # L, M, N: Phân luồng sau THCS năm qua
        self.assertEqual(p_cnt['L'], 19192.0)
        self.assertEqual(p_cnt['M'], 6226.0)
        self.assertEqual(p_cnt['N'], 815.0)
        # Tổng phân luồng phải khớp 100% với HS lớp 9
        self.assertEqual(p_cnt['L'] + p_cnt['M'] + p_cnt['N'], p_cnt['I'])
        
        # O, P, Q, R, S: Đối tượng 18-21
        self.assertEqual(p_cnt['O'], 28861.0)
        self.assertEqual(p_cnt['P'], 21134.0)
        self.assertEqual(p_cnt['Q'], 3686.0)
        self.assertEqual(p_cnt['R'], 490.0)
        self.assertEqual(p_cnt['S'], 3294.0)

    def test_03_province_rates(self):
        """Kiểm tra các tỷ lệ % tính chuẩn cấp tỉnh"""
        p_rt = self.result['province_rates']
        
        # F: Tỷ lệ 15-18 TN THCS = 97.56% (khớp 100% với 2.M1 H55=97.56% và 2.M3 J46=97.56%)
        self.assertEqual(p_rt['F'], 97.56, "Tỷ lệ 15-18 TN THCS phải là 97.56%")
        
        # K: Tỷ lệ TN THCS năm qua = 99.68%
        self.assertEqual(p_rt['K'], 99.68, "Tỷ lệ TN THCS năm qua phải là 99.68%")
        
        # T: Tỷ lệ vào TCCN = 1.70%
        self.assertEqual(p_rt['T'], 1.70)
        
        # U: Tỷ lệ vào Dạy nghề = 11.41%
        self.assertEqual(p_rt['U'], 11.41)
        
        # V: Tỷ lệ có bằng TN THPT, BTTHPT, TCCN, Dạy nghề = 99.11%
        self.assertEqual(p_rt['V'], 99.11)

    def test_04_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và đọc lại các sheet"""
        out_file = export_2_m9_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_file), "File Excel kết quả phải tồn tại")
        
        wb = openpyxl.load_workbook(out_file, data_only=True)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_92_XA', wb.sheetnames)
        
        ws_th = wb['TONG_HOP_THU']
        self.assertEqual(ws_th['D48'].value, 100561.0)
        self.assertEqual(ws_th['E48'].value, 98111.0)
        self.assertEqual(ws_th['F48'].value, 97.56)
        self.assertEqual(ws_th['K48'].value, 99.68)
        self.assertEqual(ws_th['V48'].value, 99.11)
        
        wb.close()
        print(f"\n[OK] Xuất Excel 2.M9 thành công: {out_file}")

if __name__ == '__main__':
    unittest.main()
