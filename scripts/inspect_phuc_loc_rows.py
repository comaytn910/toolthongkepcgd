import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
fp = r"D:\DATA PCGD\1810\32_32. Xã Phúc Lộc - Nhập hoàn chỉnh PCGD 2026.xlsx"
wb = openpyxl.load_workbook(fp, data_only=True)
wb_raw = openpyxl.load_workbook(fp, data_only=False)

for s in ['3.M4', '3.M5', '3.M6']:
    ws = wb[s]
    ws_r = wb_raw[s]
    print(f"\n=== SHEET {s} PHÚC LỘC (Hàng 42 đến 50) ===")
    for r in range(42, 51):
        v_a = ws.cell(r, 1).value
        v_b = ws.cell(r, 2).value
        v_c = ws.cell(r, 3).value
        v_g = ws.cell(r, 7).value
        v_c_form = ws_r.cell(r, 3).value
        v_g_form = ws_r.cell(r, 7).value
        print(f"  H{r:<2} | A={v_a!r} | B={v_b!r} | C={v_c} (form: {v_c_form}) | G={v_g} (form: {v_g_form})")

wb.close()
wb_raw.close()
