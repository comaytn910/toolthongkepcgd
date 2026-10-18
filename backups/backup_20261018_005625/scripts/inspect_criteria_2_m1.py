import openpyxl

with open('scratch/criteria_2_m1.txt', 'w', encoding='utf-8') as out:
    for fname in ['01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx', '02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx']:
        wb = openpyxl.load_workbook(f'data/{fname}', data_only=False)
        ws = wb['2.M1']
        wb_v = openpyxl.load_workbook(f'data/{fname}', data_only=True)
        ws_v = wb_v['2.M1']
        out.write(f'=== {fname} ===\n')
        for r in range(51, 60):
            name = ws.cell(r, 1).value
            f_g = ws.cell(r, 7).value
            v_g = ws_v.cell(r, 7).value
            f_h = ws.cell(r, 8).value
            v_h = ws_v.cell(r, 8).value
            out.write(f'R{r}: {name} | G_f={f_g} G_v={v_g} | H_f={f_h} H_v={v_h}\n')

print("Criteria dumped.")
