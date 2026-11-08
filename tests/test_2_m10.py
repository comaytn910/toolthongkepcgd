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
from processors.m2_m10 import process_2_m10
from exporters.excel_exporter import export_2_m10_test

class TestProcess2M10(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_2_m10(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_province_counts(self):
        """Kiểm tra các chỉ tiêu số lượng cấp tỉnh"""
        p_cnt = self.result['province_counts']
        
        # C, D: Đối tượng PC 18-21
        self.assertEqual(p_cnt['C'], 88368.0)
        self.assertEqual(p_cnt['D'], 86483.0)
        
        # E, F, G: Có bằng TN
        self.assertEqual(p_cnt['E'], 68041.0)
        self.assertEqual(p_cnt['F'], 3312.0)
        self.assertEqual(p_cnt['G'], 4522.0)
        sum_efg = p_cnt['E'] + p_cnt['F'] + p_cnt['G']
        self.assertEqual(sum_efg, 75875.0, "Tổng số người 18-21 có bằng tốt nghiệp phải là 75,875")
        
        # I, J: Lớp 12, TN THPT năm qua
        self.assertEqual(p_cnt['I'], 20917.0)
        self.assertEqual(p_cnt['J'], 20605.0)
        
        # L, M: Cuối cấp GDTX/TCCN, TN GDTX/TCCN năm qua
        self.assertEqual(p_cnt['L'], 2617.0)
        self.assertEqual(p_cnt['M'], 2684.0)
        
        # O, P, Q: Đi học sau THPT
        self.assertEqual(p_cnt['O'], 29781.0)
        self.assertEqual(p_cnt['P'], 5079.0)
        self.assertEqual(p_cnt['Q'], 757.0)
        
        # R, S: Đào tạo nghề
        self.assertEqual(p_cnt['R'], 5901.0)
        self.assertEqual(p_cnt['S'], 4613.0)

    def test_03_province_rates(self):
        """Kiểm tra các tỷ lệ % tính chuẩn cấp tỉnh"""
        p_rt = self.result['province_rates']
        
        # H: Tỷ lệ 18-21 có bằng tốt nghiệp = 87.73%
        self.assertEqual(p_rt['H'], 87.73, "Tỷ lệ 18-21 có bằng tốt nghiệp phải là 87.73%")
        
        # K: Tỷ lệ TN THPT năm qua = 98.51%
        self.assertEqual(p_rt['K'], 98.51, "Tỷ lệ TN THPT năm qua phải là 98.51%")
        
        # N: Tỷ lệ TN GDTX/TCCN năm qua = 102.56%
        self.assertEqual(p_rt['N'], 102.56)
        
        # T: Tỷ lệ có bằng nghề = 78.17%
        self.assertEqual(p_rt['T'], 78.17)

    def test_04_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và đọc lại các sheet"""
        out_file = export_2_m10_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_file), "File Excel kết quả phải tồn tại")
        
        wb = openpyxl.load_workbook(out_file, data_only=True)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_92_XA', wb.sheetnames)
        
        ws_th = wb['TONG_HOP_THU']
        self.assertEqual(ws_th['D47'].value, 86483.0)
        self.assertEqual(ws_th['E47'].value, 68041.0)
        self.assertEqual(ws_th['H47'].value, 87.73)
        self.assertEqual(ws_th['K47'].value, 98.51)
        self.assertEqual(ws_th['T47'].value, 78.17)
        
        wb.close()
        print(f"\n[OK] Xuất Excel 2.M10 thành công: {out_file}")

if __name__ == '__main__':
    unittest.main()
