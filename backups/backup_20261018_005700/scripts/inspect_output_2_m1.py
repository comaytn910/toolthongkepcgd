import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('outputs/TONG_HOP_THU_2.M1_PCGD_2026.xlsx', data_only=False)
print('Sheet names:', wb.sheetnames)
ws = wb['TONG_HOP_THU']
print('G3 title:', ws['G3'].value)
print('G4 note:', ws['G4'].value)

for r in [55, 56, 57, 58, 59]:
    cg = ws[f'G{r}']
    ch = ws[f'H{r}']
    a_val = str(ws[f'A{r}'].value or '').replace('\n', ' ')
    print(f'Hàng {r} ({a_val}):')
    print(f'   G{r} (SL): val={cg.value}, format={cg.number_format}')
    print(f'   H{r} (TL): val={ch.value}, format={ch.number_format}')

ws_z = wb['DOI_CHIEU_Z_NGUON']
print('\nDOI_CHIEU_Z_NGUON preview:')
for r in range(4, 12):
    row_vals = [ws_z.cell(r, c).value for c in range(1, 9)]
    print(f'R{r}: {row_vals}')

wb.close()
