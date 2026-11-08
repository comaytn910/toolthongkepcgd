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
from processors.m1_m6 import process_1_m6
from exporters.excel_exporter import export_1_m6_test

class TestProcess1M6(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_1_m6(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_grand_totals(self):
        """Kiểm tra tổng số học sinh học ngoài xã (trong tỉnh) và ngoài tỉnh"""
        agg = self.result['aggregated_cells']
        
        # Phần I (AR27): Ngoài xã trong tỉnh
        self.assertEqual(agg['AR27'], 9165.0, "Tổng học sinh học ngoài xã trong tỉnh phải là 9,165")
        self.assertEqual(self.result['total_part1'], 9165.0)
        
        # Phần II (AR48): Ngoài tỉnh
        self.assertEqual(agg['AR48'], 1209.0, "Tổng học sinh học ngoài tỉnh phải là 1,209")
        self.assertEqual(self.result['total_part2'], 1209.0)
        
        # Tổng cộng (AR49) = AR27 + AR48
        expected_grand = 9165.0 + 1209.0
        self.assertEqual(agg['AR49'], 10374.0, "Tổng cộng học sinh đi học ngoài địa bàn phải là 10,374")
        self.assertEqual(self.result['total_grand'], expected_grand)

    def test_03_age_breakdown(self):
        """Kiểm tra phân tích độ tuổi: nhóm 6-10t chiếm tuyệt đại đa số"""
        by_age = self.result['province_by_age']
        
        # Tổng 6-10 tuổi
        sum_6_10 = sum(by_age[f"{a} Tuổi"]['total'] for a in range(6, 11))
        self.assertEqual(sum_6_10, 10328.0, "Nhóm 6-10 tuổi phải là 10,328 hs")
        
        # 6 tuổi
        self.assertEqual(by_age['6 Tuổi']['total'], 2214.0)
        # 7 tuổi
        self.assertEqual(by_age['7 Tuổi']['total'], 2243.0)
        # 8 tuổi
        self.assertEqual(by_age['8 Tuổi']['total'], 1996.0)
        # 9 tuổi
        self.assertEqual(by_age['9 Tuổi']['total'], 2114.0)
        # 10 tuổi
        self.assertEqual(by_age['10 Tuổi']['total'], 1761.0)

    def test_04_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và các định dạng chuẩn"""
        out_path = export_1_m6_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_path), "File xuất phải tồn tại")
        self.assertTrue(os.path.getsize(out_path) > 10000, "Dung lượng file phải hợp lý")
        
        wb = openpyxl.load_workbook(out_path)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_THEO_XA', wb.sheetnames)
        self.assertIn('DANH_SACH_TRUONG_NGOAI', wb.sheetnames)
        
        ws = wb['TONG_HOP_THU']
        self.assertEqual(ws['AR27'].value, 9165.0)
        self.assertEqual(ws['AR48'].value, 1209.0)
        self.assertEqual(ws['AR49'].value, 10374.0)
        
        self.assertEqual(ws['AR49'].number_format, '#,##0', "Format số lượng phải là #,##0")
        
        wb.close()
        print(f"\n[OK] Đã xuất file thành công: {out_path} ({os.path.getsize(out_path):,} bytes)")

if __name__ == '__main__':
    unittest.main()
