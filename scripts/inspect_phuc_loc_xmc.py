import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

fp = r"D:\DATA PCGD\1810\32_32. Xã Phúc Lộc - Nhập hoàn chỉnh PCGD 2026.xlsx"
wb = openpyxl.load_workbook(fp, data_only=True)
wb_raw = openpyxl.load_workbook(fp, data_only=False)

print("=== SOI CHI TIẾT TẠI FILE XÃ PHÚC LỘC ===")

# 1. 3.M1 vs 3.M4
ws1 = wb['3.M1']
ws1_r = wb_raw['3.M1']
ws4 = wb['3.M4']
ws4_r = wb_raw['3.M4']

print("\n--- 3.M1: Hàng tổng ---")
for r in range(1, ws1.max_row + 1):
    v0 = str(ws1.cell(r, 1).value or '')
    v1 = str(ws1.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M1 tại Hàng {r}:")
        print(f"  C{r} (Dân số): val={ws1.cell(r, 3).value}, formula={ws1_r.cell(r, 3).value}")
        print(f"  D{r} (Biết chữ M1): val={ws1.cell(r, 4).value}, formula={ws1_r.cell(r, 4).value}")
        print(f"  F{r} (Biết chữ M2): val={ws1.cell(r, 6).value}, formula={ws1_r.cell(r, 6).value}")
        
print("\n--- 3.M4: Hàng tổng ---")
for r in range(1, ws4.max_row + 1):
    v0 = str(ws4.cell(r, 1).value or '')
    v1 = str(ws4.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M4 tại Hàng {r}:")
        print(f"  C{r} (Dân số): val={ws4.cell(r, 3).value}, formula={ws4_r.cell(r, 3).value}")
        print(f"  G{r} (Mù chữ M1): val={ws4.cell(r, 7).value}, formula={ws4_r.cell(r, 7).value}")

# 2. 3.M2 vs 3.M5
ws2 = wb['3.M2']
ws2_r = wb_raw['3.M2']
ws5 = wb['3.M5']
ws5_r = wb_raw['3.M5']
print("\n--- 3.M2: Hàng tổng ---")
for r in range(1, ws2.max_row + 1):
    v0 = str(ws2.cell(r, 1).value or '')
    v1 = str(ws2.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M2 tại Hàng {r}: C={ws2.cell(r, 3).value} (form: {ws2_r.cell(r, 3).value}), D={ws2.cell(r, 4).value} (form: {ws2_r.cell(r, 4).value})")
        
print("\n--- 3.M5: Hàng tổng ---")
for r in range(1, ws5.max_row + 1):
    v0 = str(ws5.cell(r, 1).value or '')
    v1 = str(ws5.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M5 tại Hàng {r}: C={ws5.cell(r, 3).value} (form: {ws5_r.cell(r, 3).value}), G={ws5.cell(r, 7).value} (form: {ws5_r.cell(r, 7).value})")

# 3. 3.M3 vs 3.M6
ws3 = wb['3.M3']
ws3_r = wb_raw['3.M3']
ws6 = wb['3.M6']
ws6_r = wb_raw['3.M6']
print("\n--- 3.M3: Hàng tổng ---")
for r in range(1, ws3.max_row + 1):
    v0 = str(ws3.cell(r, 1).value or '')
    v1 = str(ws3.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M3 tại Hàng {r}: C={ws3.cell(r, 3).value} (form: {ws3_r.cell(r, 3).value}), D={ws3.cell(r, 4).value} (form: {ws3_r.cell(r, 4).value})")
        
print("\n--- 3.M6: Hàng tổng ---")
for r in range(1, ws6.max_row + 1):
    v0 = str(ws6.cell(r, 1).value or '')
    v1 = str(ws6.cell(r, 2).value or '').lower()
    if v0 == '@' or 'cộng' in v1:
        print(f"Dòng tổng 3.M6 tại Hàng {r}: C={ws6.cell(r, 3).value} (form: {ws6_r.cell(r, 3).value}), G={ws6.cell(r, 7).value} (form: {ws6_r.cell(r, 7).value})")

wb.close()
wb_raw.close()
