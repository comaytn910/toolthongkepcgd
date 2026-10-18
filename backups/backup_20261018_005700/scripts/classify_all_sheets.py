import openpyxl
from openpyxl.utils import get_column_letter
import json
import sys

def inspect_sheet_structure(ws, sheet_name):
    # Find real max row and col
    real_max_r = 0
    real_min_r = 9999
    real_max_c = 0
    real_min_c = 9999
    
    val_count = 0
    formulas = []
    text_cells = []
    number_cells = []
    
    for r in range(1, min(ws.max_row + 1, 150)):
        for c in range(1, min(ws.max_column + 1, 150)):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                val_count += 1
                if r > real_max_r: real_max_r = r
                if r < real_min_r: real_min_r = r
                if c > real_max_c: real_max_c = c
                if c < real_min_c: real_min_c = c
                
                coord = f"{get_column_letter(c)}{r}"
                val_str = str(val).strip()
                if val_str.startswith('='):
                    formulas.append((coord, val_str))
                elif isinstance(val, (int, float)):
                    number_cells.append((coord, val))
                else:
                    text_cells.append((coord, val_str))
                    
    # Scan for "Tổng cộng", "Cộng", "Tổng"
    total_rows = []
    for coord, txt in text_cells:
        if any(w in txt.lower() for w in ['tổng cộng', 'tổng số', 'cộng']):
            r = int(''.join(filter(str.isdigit, coord)))
            if r not in total_rows:
                total_rows.append(r)
                
    # Detect title
    title = ""
    for r in range(1, min(real_max_r + 1, 6)):
        for c in range(1, min(real_max_c + 1, 8)):
            v = ws.cell(row=r, column=c).value
            if v and isinstance(v, str) and len(v.strip()) > 8 and not v.startswith('='):
                title = v.strip().replace('\n', ' ')
                break
        if title: break

    # Check headers
    header_rows = []
    for r in range(max(1, real_min_r), min(real_max_r + 1, 12)):
        # check if row has multiple text cells indicating column headers
        texts_in_row = [ws.cell(row=r, column=c).value for c in range(real_min_c, real_max_c + 1) if isinstance(ws.cell(row=r, column=c).value, str)]
        if len(texts_in_row) >= 3:
            header_rows.append(r)

    return {
        'sheet': sheet_name,
        'title': title[:120],
        'real_range': f"{get_column_letter(real_min_c)}{real_min_r}:{get_column_letter(real_max_c)}{real_max_r}" if val_count > 0 else "EMPTY",
        'real_min_row': real_min_r if val_count > 0 else 0,
        'real_max_row': real_max_r if val_count > 0 else 0,
        'real_min_col': real_min_c if val_count > 0 else 0,
        'real_max_col': real_max_c if val_count > 0 else 0,
        'val_count': val_count,
        'formula_count': len(formulas),
        'number_count': len(number_cells),
        'header_rows': header_rows,
        'total_rows': total_rows[:5],
        'sample_formulas': formulas[:5]
    }

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    wb = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=False)
    results = []
    for s in wb.sheetnames:
        res = inspect_sheet_structure(wb[s], s)
        results.append(res)
        
    with open('all_sheets_structure.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Analyzed {len(results)} sheets.")

if __name__ == '__main__':
    main()
