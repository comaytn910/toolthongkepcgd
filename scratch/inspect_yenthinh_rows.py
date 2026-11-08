import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('data/02_02. Xã Yên Thịnh - Nhập hoàn chỉnh PCGD 2026.xlsx', data_only=True)

for s in ['3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']:
    ws = wb[s]
    print(f"\n{'='*20} Yên Thịnh -> {s} {'='*20}")
    for r in range(1, 55):
        vals = [ws.cell(r, c).value for c in range(1, 15)]
        if any(v is not None for v in vals):
            v_str = ' '.join(str(x) for x in vals[:3] if x is not None)
            if any(k in v_str.lower() for k in ['tổng', 'cộng', '@', 'xóm', 'thôn', '1', '2']) or r <= 8 or r >= 45:
                print(f"R{r:02d}: {[str(v)[:15] if v is not None else '' for v in vals]}")
