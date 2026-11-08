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
from processors.m2_m7 import process_2_m7
from exporters.excel_exporter import export_2_m7_test

class TestProcess2M7(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_2_m7(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_grand_totals(self):
        """Kiểm tra tổng số học sinh học ngoài xã (trong tỉnh) và ngoài tỉnh cấp THCS"""
        agg = self.result['aggregated_cells']
        
        # Phần I (AM21): Ngoài xã trong tỉnh
        self.assertEqual(agg['AM21'], 3627.0, "Tổng học sinh học ngoài xã trong tỉnh phải là 3,627")
        self.assertEqual(self.result['total_part1'], 3627.0)
        
        # Phần II (AM39): Ngoài tỉnh
        self.assertEqual(agg['AM39'], 317.0, "Tổng học sinh học ngoài tỉnh phải là 317")
        self.assertEqual(self.result['total_part2'], 317.0)
        
        # Tổng cộng (AM40) = AM21 + AM39
        expected_grand = 3627.0 + 317.0
        self.assertEqual(agg['AM40'], 3944.0, "Tổng cộng học sinh THCS đi học ngoài địa bàn phải là 3,944")
        self.assertEqual(self.result['total_grand'], expected_grand)

    def test_03_age_breakdown(self):
        """Kiểm tra cơ cấu độ tuổi: nhóm 11-14 tuổi (đúng độ tuổi THCS) chiếm tuyệt đại đa số"""
        by_age = self.result['age_breakdown']
        
        # 10 tuổi: 67
        self.assertEqual(by_age['10 tuổi']['total'], 67.0)
        # 11 tuổi: 1,030
        self.assertEqual(by_age['11 tuổi']['total'], 1030.0)
        # 12 tuổi: 982
        self.assertEqual(by_age['12 tuổi']['total'], 982.0)
        # 13 tuổi: 978
        self.assertEqual(by_age['13 tuổi']['total'], 978.0)
        # 14 tuổi: 903
        self.assertEqual(by_age['14 tuổi']['total'], 903.0)
        # 15 tuổi: 12
        self.assertEqual(by_age['15 tuổi']['total'], 12.0)
        # 16 tuổi: 2
        self.assertEqual(by_age['16 tuổi']['total'], 2.0)
        # 17 tuổi: 0
        self.assertEqual(by_age['17 tuổi']['total'], 0.0)
        # 18 tuổi: 0
        self.assertEqual(by_age['18 tuổi']['total'], 0.0)
        
        # Tổng nhóm đúng độ tuổi THCS (11-14t)
        sum_11_14 = by_age['11 tuổi']['total'] + by_age['12 tuổi']['total'] + by_age['13 tuổi']['total'] + by_age['14 tuổi']['total']
        self.assertEqual(sum_11_14, 3893.0, "Nhóm 11-14 tuổi phải đạt 3,893 học sinh (chiếm 98.7%)")

    def test_04_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và đọc lại các sheet"""
        out_file = export_2_m7_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_file), "File Excel kết quả phải tồn tại")
        
        wb = openpyxl.load_workbook(out_file, data_only=True)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_THEO_XA', wb.sheetnames)
        self.assertIn('DANH_SACH_TRUONG_NGOAI', wb.sheetnames)
        
        ws_th = wb['TONG_HOP_THU']
        self.assertEqual(ws_th['AM21'].value, 3627.0)
        self.assertEqual(ws_th['AM39'].value, 317.0)
        self.assertEqual(ws_th['AM40'].value, 3944.0)
        
        wb.close()
        print(f"\n[OK] Xuất Excel 2.M7 thành công: {out_file}")

if __name__ == '__main__':
    unittest.main()
