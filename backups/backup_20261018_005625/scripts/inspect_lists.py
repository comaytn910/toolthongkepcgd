import openpyxl
import sys

def inspect_specific_sheets(file_path):
    sys.stdout.reconfigure(encoding='utf-8')
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    for s_name in ['0.M3', '0.M7', '1.M6', '1.M7', '2.M7', '2.M8']:
        ws = wb[s_name]
        print(f"\n=================== SHEET: {s_name} ===================")
        for r in range(1, 15):
            vals = [f"{openpyxl.utils.get_column_letter(c)}{r}={repr(ws.cell(row=r, column=c).value)}" for c in range(1, 12) if ws.cell(row=r, column=c).value is not None]
            if vals:
                print(f"R{r:02d}:", " | ".join(vals[:6]))

if __name__ == '__main__':
    inspect_specific_sheets('MẪU CHUẨN - PCGD 2026(ok).xlsx')
