import sys
import glob
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

files = sorted(glob.glob('data/*.xlsx'))

# Kiểm tra các xã có dat_m1=False và dat_m2=False trong inspect trước
for f in files[7:15]:
    commune_name = f.split('\\')[-1].split(' - ')[0]
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb['3.M1']
    for r in range(40, min(52, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, 10)]
        if any(v is not None for v in row_vals):
            print(f"{commune_name} R{r}: {row_vals}")
