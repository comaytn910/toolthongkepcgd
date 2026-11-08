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
from processors.m2_m8 import process_2_m8
from exporters.excel_exporter import export_2_m8_test

class TestProcess2M8(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_2_m8(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_grand_totals(self):
        """Kiểm tra tổng số học sinh ngoại địa bàn đến học tại xã cấp THCS"""
        agg = self.result['aggregated_cells']
        
        # Phần I (AM24): Học sinh xã ngoài trong tỉnh đến học
        self.assertEqual(agg['AM24'], 10538.0, "Tổng học sinh xã ngoài trong tỉnh đến học phải là 10,538")
        self.assertEqual(self.result['total_part1'], 10538.0)
        
        # Phần II (AM38): Học sinh ngoài tỉnh đến học
        self.assertEqual(agg['AM38'], 427.0, "Tổng học sinh ngoài tỉnh đến học phải là 427")
        self.assertEqual(self.result['total_part2'], 427.0)
        
        # Tổng cộng (AM39) = AM24 + AM38
        expected_grand = 10538.0 + 427.0
        self.assertEqual(agg['AM39'], 10965.0, "Tổng cộng học sinh ngoại địa bàn đến học cấp THCS phải là 10,965")
        self.assertEqual(self.result['total_grand'], expected_grand)

    def test_03_age_breakdown(self):
        """Kiểm tra phân tích độ tuổi: nhóm 11-14 tuổi (đúng độ tuổi THCS) chiếm tuyệt đại đa số"""
        by_age = self.result['age_breakdown']
        
        # 10 tuổi: 9
        self.assertEqual(by_age['10 tuổi']['total'], 9.0)
        # 11 tuổi: 2,645
        self.assertEqual(by_age['11 tuổi']['total'], 2645.0)
        # 12 tuổi: 2,614
        self.assertEqual(by_age['12 tuổi']['total'], 2614.0)
        # 13 tuổi: 2,664
        self.assertEqual(by_age['13 tuổi']['total'], 2664.0)
        # 14 tuổi: 2,933
        self.assertEqual(by_age['14 tuổi']['total'], 2933.0)
        # 15 tuổi: 44
        self.assertEqual(by_age['15 tuổi']['total'], 44.0)
        # 16 tuổi: 11
        self.assertEqual(by_age['16 tuổi']['total'], 11.0)
        # 17 tuổi: 1
        self.assertEqual(by_age['17 tuổi']['total'], 1.0)
        # 18 tuổi: 0
        self.assertEqual(by_age['18 tuổi']['total'], 0.0)
        
        # Tổng nhóm đúng độ tuổi THCS (11-14t)
        sum_11_14 = by_age['11 tuổi']['total'] + by_age['12 tuổi']['total'] + by_age['13 tuổi']['total'] + by_age['14 tuổi']['total']
        self.assertEqual(sum_11_14, 10856.0, "Nhóm 11-14 tuổi phải đạt 10,856 học sinh (chiếm 99.0%)")

    def test_04_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và đọc lại các sheet"""
        out_file = export_2_m8_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_file), "File Excel kết quả phải tồn tại")
        
        wb = openpyxl.load_workbook(out_file, data_only=True)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_THEO_XA', wb.sheetnames)
        self.assertIn('DANH_SACH_TRUONG_NGOAI', wb.sheetnames)
        self.assertIn('DOI_CHIEU_SAI_LECH', wb.sheetnames)
        
        ws_th = wb['TONG_HOP_THU']
        self.assertEqual(ws_th['AM24'].value, 10538.0)
        self.assertEqual(ws_th['AM38'].value, 427.0)
        self.assertEqual(ws_th['AM39'].value, 10965.0)
        
        wb.close()
        print(f"\n[OK] Xuất Excel 2.M8 thành công: {out_file}")

if __name__ == '__main__':
    unittest.main()
