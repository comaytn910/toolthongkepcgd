import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=False)

sheets = ['3.M1', '3.M2', '3.M3', '3.M4', '3.M5', '3.M6', '3.M7', '3.M8', '3.M9']

for s in sheets:
    if s in wb.sheetnames:
        ws = wb[s]
        print(f"\n{'='*30} {s} (max_row={ws.max_row}, max_col={ws.max_column}) {'='*30}")
        # Tìm các dòng tiêu đề và cấu trúc
        for r in range(1, min(12, ws.max_row + 1)):
            vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
            if any(v is not None for v in vals):
                clean_vals = [str(v).replace('\n', ' ')[:20] if v is not None else '' for v in vals]
                print(f"R{r:02d}: {clean_vals}")
        
        # Tìm dòng Cộng hoặc các dòng đặc biệt
        print("--- CÁC DÒNG ĐẶC BIỆT / CUỐI ---")
        for r in range(12, min(ws.max_row + 1, 60)):
            val_a = str(ws.cell(r, 1).value or '')
            val_b = str(ws.cell(r, 2).value or '')
            if any(kw in val_a.lower() or kw in val_b.lower() for kw in ['cộng', 'chuẩn', 'tổng', 'đạt', '@']):
                vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
                clean_vals = [str(v).replace('\n', ' ')[:20] if v is not None else '' for v in vals]
                print(f"Special R{r:02d}: {clean_vals}")
    else:
        print(f"{s}: KHÔNG TÌM THẤY")
