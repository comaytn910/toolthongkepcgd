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
from processors.m1_m5 import process_1_m5
from exporters.excel_exporter import export_1_m5_test

class TestProcess1M5(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data_dir = os.path.join(BASE_DIR, 'data')
        cls.template_path = os.path.join(BASE_DIR, 'MẪU CHUẨN - PCGD 2026(ok).xlsx')
        cls.output_dir = os.path.join(BASE_DIR, 'outputs')
        
        scan_res = scan_data_directory(cls.data_dir)
        cls.valid_files = [f for f in scan_res['files'] if f['is_valid']]
        cls.result = process_1_m5(cls.valid_files)

    def test_01_commune_count(self):
        """Kiểm tra số lượng xã hợp lệ"""
        self.assertEqual(self.result['total_communes_count'], 92, "Phải quét đủ 92 xã")
        self.assertEqual(len(self.result['commune_names']), 92)

    def test_02_grand_totals(self):
        """Kiểm tra tổng số trẻ khuyết tật và các cột dạng tật toàn tỉnh"""
        agg = self.result['aggregated_cells']
        
        # 8 dạng tật hàng 17
        d17 = agg['D17'] # Vận động
        e17 = agg['E17'] # Nghe, nói
        f17 = agg['F17'] # Nhìn
        g17 = agg['G17'] # Thần kinh, tâm thần
        h17 = agg['H17'] # Trí tuệ
        i17 = agg['I17'] # Tự kỷ
        j17 = agg['J17'] # Học tập
        k17 = agg['K17'] # Khác
        
        self.assertEqual(d17, 436.0, "Vận động D17 phải bằng 436")
        self.assertEqual(e17, 237.0, "Nghe nói E17 phải bằng 237")
        self.assertEqual(f17, 125.0, "Nhìn F17 phải bằng 125")
        self.assertEqual(g17, 267.0, "Thần kinh tâm thần G17 phải bằng 267")
        self.assertEqual(h17, 2304.0, "Trí tuệ H17 phải bằng 2,304")
        self.assertEqual(i17, 54.0, "Tự kỷ I17 phải bằng 54")
        self.assertEqual(j17, 104.0, "Học tập J17 phải bằng 104")
        self.assertEqual(k17, 270.0, "Khác K17 phải bằng 270")
        
        # Tổng tật C17 = SUM(D:K)
        expected_c17 = d17 + e17 + f17 + g17 + h17 + i17 + j17 + k17
        self.assertEqual(agg['C17'], 3797.0, "Tổng số trẻ khuyết tật C17 phải bằng 3,797")
        self.assertEqual(agg['C17'], expected_c17, "C17 phải bằng đúng tổng D17..K17")
        
        # Tiếp cận GD L17
        self.assertEqual(agg['L17'], 3307.0, "Tiếp cận GD L17 phải bằng 3,307")
        
        # Tỷ lệ tiếp cận GD M17 = L17/C17*100
        expected_rate = round(3307.0 / 3797.0 * 100.0, 2)
        self.assertEqual(agg['M17'], 87.10, "Tỷ lệ tiếp cận GD M17 phải bằng 87.10%")
        self.assertEqual(agg['M17'], expected_rate)

    def test_03_age_breakdown(self):
        """Kiểm tra chi tiết từng độ tuổi từ 6 đến 14"""
        agg = self.result['aggregated_cells']
        expected_c_by_age = {
            8: 262.0,  # 6 tuổi (2020)
            9: 378.0,  # 7 tuổi (2019)
            10: 397.0, # 8 tuổi (2018)
            11: 458.0, # 9 tuổi (2017)
            12: 390.0, # 10 tuổi (2016)
            13: 459.0, # 11 tuổi (2015)
            14: 456.0, # 12 tuổi (2014)
            15: 474.0, # 13 tuổi (2013)
            16: 523.0  # 14 tuổi (2012)
        }
        for r, exp_c in expected_c_by_age.items():
            self.assertEqual(agg[f"C{r}"], exp_c, f"Tổng trẻ KT hàng {r} phải là {exp_c}")
            # Cột C phải bằng tổng D:K
            sum_dk = sum(agg[f"{c}{r}"] for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'])
            self.assertEqual(agg[f"C{r}"], sum_dk, f"Hàng {r}: C phải bằng tổng D:K")

    def test_04_yen_phong_discrepancy(self):
        """Kiểm tra phát hiện sai lệch dòng tổng 17 nguồn của xã Yên Phong"""
        discrepancies = self.result['h17_discrepancies']
        yp_found = any('Yên Phong' in d['commune'] for d in discrepancies)
        self.assertTrue(yp_found, "Phải phát hiện sai lệch ở xã Yên Phong do gõ đè số 0 vào hàng 17")

    def test_05_export_excel(self):
        """Kiểm tra xuất file Excel kết quả và các định dạng chuẩn"""
        out_path = export_1_m5_test(self.result, self.template_path, self.output_dir)
        self.assertTrue(os.path.exists(out_path), "File xuất phải tồn tại")
        self.assertTrue(os.path.getsize(out_path) > 10000, "Dung lượng file phải hợp lý")
        
        wb = openpyxl.load_workbook(out_path)
        self.assertIn('TONG_HOP_THU', wb.sheetnames)
        self.assertIn('CHI_TIET_THEO_XA', wb.sheetnames)
        self.assertIn('DOI_CHIEU_H17_NGUON', wb.sheetnames)
        
        ws = wb['TONG_HOP_THU']
        self.assertEqual(ws['C17'].value, 3797.0)
        self.assertEqual(ws['L17'].value, 3307.0)
        self.assertEqual(ws['M17'].value, 87.10)
        
        # Định dạng tỷ lệ M17 phải là 0.00"%"
        self.assertEqual(ws['M17'].number_format, '0.00"%"', "Format cột M phải là 0.00\"%\"")
        self.assertEqual(ws['C17'].number_format, '#,##0', "Format cột C phải là #,##0")
        
        wb.close()
        print(f"\n[OK] Đã xuất file thành công: {out_path} ({os.path.getsize(out_path):,} bytes)")

if __name__ == '__main__':
    unittest.main()
