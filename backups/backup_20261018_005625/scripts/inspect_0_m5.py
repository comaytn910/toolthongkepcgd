import openpyxl
import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

template_path = 'MẪU CHUẨN - PCGD 2026(ok).xlsx'
wb = openpyxl.load_workbook(template_path, data_only=False)
ws = wb['0.M5']

print("=== 0.M5 IN TEMPLATE ===")
print("Dimensions:", ws.dimensions)
for r in range(1, 45):
    row_vals = []
    has_content = False
    for c in range(1, 15):
        cell = ws.cell(row=r, column=c)
        v = cell.value
        f = cell.number_format
        if v is not None:
            has_content = True
        row_vals.append(f"{cell.coordinate}: {repr(v)} (fmt={f})")
    if has_content:
        print(f"Row {r:02d}: " + " | ".join(row_vals[:10]))

wb.close()

commune_files = glob.glob('data/*.xlsx')
for cf in commune_files:
    print(f"\n=== 0.M5 IN COMMUNE: {os.path.basename(cf)} ===")
    wb_c = openpyxl.load_workbook(cf, data_only=False)
    if '0.M5' in wb_c.sheetnames:
        ws_c = wb_c['0.M5']
        wb_c_val = openpyxl.load_workbook(cf, data_only=True)
        ws_c_val = wb_c_val['0.M5']
        for r in range(1, 42):
            row_items = []
            has_content = False
            for col_idx in range(1, 12):
                coord = ws_c.cell(row=r, column=col_idx).coordinate
                f_val = ws_c[coord].value
                d_val = ws_c_val[coord].value
                fmt = ws_c[coord].number_format
                if f_val is not None or d_val is not None:
                    has_content = True
                row_items.append(f"{coord}: val={d_val!r}, f={f_val!r}, fmt={fmt}")
            if has_content:
                print(f"Row {r:02d}: " + " | ".join(row_items[:8]))
        wb_c_val.close()
    else:
        print("Sheet 0.M5 NOT FOUND!")
    wb_c.close()
