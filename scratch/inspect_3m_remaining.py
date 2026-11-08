import sys
import glob
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

files = sorted(glob.glob('data/*.xlsx'))

# Chọn 5 xã để khảo sát cấu trúc của 3.M4..3.M9
test_files = files[:5]

for s in ['3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
    print(f"\n{'='*30} KHẢO SÁT {s} {'='*30}")
    for f in test_files:
        c_name = f.split('\\')[-1].split(' - ')[0]
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        if s not in wb.sheetnames:
            print(f"{c_name}: THIẾU SHEET {s}")
            wb.close()
            continue
        ws = wb[s]
        # Tìm các dòng có chữ Tổng, Cộng hoặc các dòng từ 30 trở đi
        r_idx = 0
        found = False
        for row in ws.iter_rows(values_only=True):
            r_idx += 1
            if not row or len(row) < 3:
                continue
            text_join = ' '.join(str(x) for x in row[:3] if x is not None).lower()
            if any(k in text_join for k in ['tổng', 'cộng', '@']):
                print(f"{c_name} R{r_idx}: {[str(x)[:15] if x is not None else '' for x in row[:15]]}")
                found = True
        if not found:
            # In dòng cuối cùng có dữ liệu
            last_r = 0
            last_row = None
            r_idx = 0
            for row in ws.iter_rows(values_only=True):
                r_idx += 1
                if any(x is not None for x in row[:10]):
                    last_r = r_idx
                    last_row = row
            print(f"{c_name}: KHÔNG CÓ TỪ KHÓA, Dòng cuối R{last_r}: {[str(x)[:15] if x is not None else '' for x in last_row[:15]] if last_row else 'None'}")
        wb.close()
