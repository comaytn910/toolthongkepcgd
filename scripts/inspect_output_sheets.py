import glob
import openpyxl
import os

outputs = glob.glob('outputs/TONG_HOP_THU_*.xlsx')
outputs.sort()

print(f"Total output files: {len(outputs)}")
for f in outputs:
    bn = os.path.basename(f)
    try:
        wb = openpyxl.load_workbook(f, read_only=True)
        print(f"{bn} -> sheets: {wb.sheetnames}")
        wb.close()
    except Exception as e:
        print(f"{bn} -> ERROR: {e}")
