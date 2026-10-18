import openpyxl

wb = openpyxl.load_workbook('data/01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx', data_only=False)
ws = wb['2.M1']

import openpyxl
import os

os.makedirs('scratch', exist_ok=True)
with open('scratch/inspect_2_m1.txt', 'w', encoding='utf-8') as f:
    f.write('--- ROWS 4 to 10 ---\n')
    for r in range(4, 11):
        row_vals = {openpyxl.utils.get_column_letter(c): ws.cell(r, c).value for c in range(1, 30) if ws.cell(r, c).value is not None}
        f.write(f'R{r}: {row_vals}\n')

    f.write('\n--- ROWS 11 to 50 ---\n')
    for r in range(11, 51):
        tt = ws.cell(r, 2).value
        name = ws.cell(r, 3).value
        formula_z = ws.cell(r, 26).value # Col Z = 26
        # also print a few sample cells
        f.write(f'R{r}: TT={tt} | {name} | Z={formula_z}\n')

    f.write('\n--- ROWS 51 to 65 ---\n')
    for r in range(51, 65):
        vals = {openpyxl.utils.get_column_letter(c): ws.cell(r, c).value for c in range(1, 28) if ws.cell(r, c).value is not None}
        f.write(f'R{r}: {vals}\n')

print("Inspection completed.")
