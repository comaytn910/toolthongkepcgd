import os
import sys
import openpyxl

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

wb = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=False)
ws = wb['1.M2']
for c in range(1, 34):
    col = openpyxl.utils.get_column_letter(c)
    val = ws.cell(9, c).value
    h4 = ws.cell(4, c).value
    h5 = ws.cell(5, c).value
    h6 = ws.cell(6, c).value
    print(f"{col}9: val={val} | H4={h4} | H5={h5} | H6={h6}")
