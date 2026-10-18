import os
import openpyxl
import json

template_path = 'MẪU CHUẨN - PCGD 2026(ok).xlsx'
wb_template = openpyxl.load_workbook(template_path, data_only=False)
sheet_tmpl = wb_template['0.M4']

files = [f for f in os.listdir('data') if f.endswith('.xlsx') and not f.startswith('~$')]
files.sort()

wb_c1 = openpyxl.load_workbook(os.path.join('data', files[0]), data_only=False)
sheet_c1 = wb_c1['0.M4']

wb_c2 = openpyxl.load_workbook(os.path.join('data', files[1]), data_only=False)
sheet_c2 = wb_c2['0.M4']

info = {
    'file_1': files[0],
    'file_2': files[1],
    'template': {'max_row': sheet_tmpl.max_row, 'max_col': sheet_tmpl.max_column, 'rows': {}},
    'commune_1': {'max_row': sheet_c1.max_row, 'max_col': sheet_c1.max_column, 'rows': {}},
    'commune_2': {'max_row': sheet_c2.max_row, 'max_col': sheet_c2.max_column, 'rows': {}}
}

for r in range(1, 25):
    r_tmpl = {}
    r_c1 = {}
    r_c2 = {}
    for c in range(1, 30):
        col_letter = openpyxl.utils.get_column_letter(c)
        v_t = sheet_tmpl.cell(r, c).value
        v_1 = sheet_c1.cell(r, c).value
        v_2 = sheet_c2.cell(r, c).value
        if v_t is not None:
            r_tmpl[col_letter] = str(v_t)
        if v_1 is not None:
            r_c1[col_letter] = str(v_1)
        if v_2 is not None:
            r_c2[col_letter] = str(v_2)
    if r_tmpl:
        info['template']['rows'][r] = r_tmpl
    if r_c1:
        info['commune_1']['rows'][r] = r_c1
    if r_c2:
        info['commune_2']['rows'][r] = r_c2

with open('0_m4_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(info, f, ensure_ascii=False, indent=2)

print("Exported 0_m4_analysis.json successfully.")
