import openpyxl

with open('scratch/calc_2_m1.txt', 'w', encoding='utf-8') as out:
    for fname in ['01. Xã Yên Trạch - Nhập hoàn chỉnh PCGD 2026.xlsx', '02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx']:
        wb = openpyxl.load_workbook(f'data/{fname}', data_only=True)
        ws = wb['2.M1']
        
        # G55: sum(N35:Y37)
        g55 = sum(ws.cell(r, c).value or 0 for r in range(35, 38) for c in range(14, 26))
        z13 = ws.cell(13, 26).value
        h55 = round(g55 / z13 * 100, 2) if z13 else 0
        
        # G56: sum(N40:Y42)
        g56 = sum(ws.cell(r, c).value or 0 for r in range(40, 43) for c in range(14, 26))
        h56 = round(g56 / z13 * 100, 2) if z13 else 0
        
        # G57: sum(F12:Y12)
        g57 = sum(ws.cell(12, c).value or 0 for c in range(6, 26))
        denom57 = sum(ws.cell(11, c).value or 0 for c in [6, 8, 10, 12, 14, 17, 20, 23])
        h57 = round(g57 / denom57 * 100, 2) if denom57 else 0
        
        # G58: sum(F15:Y17, F20:Y22, F25:Y27, F30:Y32, F40:Y42)
        rows_g58 = list(range(15, 18)) + list(range(20, 23)) + list(range(25, 28)) + list(range(30, 33)) + list(range(40, 43))
        g58 = sum(ws.cell(r, c).value or 0 for r in rows_g58 for c in range(6, 26))
        denom58 = sum(ws.cell(13, c).value or 0 for c in [6, 8, 10, 12, 14, 17, 20, 23])
        h58 = round(g58 / denom58 * 100, 2) if denom58 else 0
        
        # G59: sum(W35:Y37)
        g59 = sum(ws.cell(r, c).value or 0 for r in range(35, 38) for c in range(23, 26))
        w13 = ws.cell(13, 23).value
        h59 = round(g59 / w13 * 100, 2) if w13 else 0
        
        out.write(f'{fname}:\n')
        out.write(f'  G55={g55} (source: {ws.cell(55, 7).value}), H55={h55}% (source: {ws.cell(55, 8).value})\n')
        out.write(f'  G56={g56} (source: {ws.cell(56, 7).value}), H56={h56}% (source: {ws.cell(56, 8).value})\n')
        out.write(f'  G57={g57} (source: {ws.cell(57, 7).value}), Denom57={denom57}, H57={h57}% (source: {ws.cell(57, 8).value})\n')
        out.write(f'  G58={g58} (source: {ws.cell(58, 7).value}), Denom58={denom58}, H58={h58}% (source: {ws.cell(58, 8).value})\n')
        out.write(f'  G59={g59} (source: {ws.cell(59, 7).value}), W13={w13}, H59={h59}% (source: {ws.cell(59, 8).value})\n')

print("Calculated criteria written.")
