import openpyxl
from openpyxl.utils import get_column_letter
import json
import sys

def analyze_sheet_0_m1(file_path):
    wb_formula = openpyxl.load_workbook(file_path, data_only=False)
    wb_data = openpyxl.load_workbook(file_path, data_only=True)
    
    ws_f = wb_formula['0.M1']
    ws_d = wb_data['0.M1']
    
    # Merged ranges
    merged_ranges = [str(m) for m in ws_f.merged_cells.ranges]
    
    cells_data = []
    
    for r in range(1, 41):
        for c in range(1, 12): # Cols A to K
            cell_f = ws_f.cell(row=r, column=c)
            cell_d = ws_d.cell(row=r, column=c)
            col_letter = get_column_letter(c)
            coord = f"{col_letter}{r}"
            
            raw_val = cell_f.value
            eval_val = cell_d.value
            
            # Fill color
            fill_color = None
            if cell_f.fill and cell_f.fill.fill_type:
                start_c = cell_f.fill.start_color
                if start_c:
                    fill_color = start_c.rgb or start_c.theme or start_c.indexed
                    
            # Formula check
            is_formula = isinstance(raw_val, str) and raw_val.startswith('=')
            
            cells_data.append({
                'row': r,
                'col': c,
                'coord': coord,
                'raw_val': str(raw_val) if raw_val is not None else None,
                'eval_val': str(eval_val) if eval_val is not None else None,
                'is_formula': is_formula,
                'data_type': cell_f.data_type,
                'fill_color': str(fill_color) if fill_color else None
            })
            
    # Rows 1 to 40 semantic structure
    rows_structure = []
    for r in range(1, 41):
        r_cells = [x for x in cells_data if x['row'] == r]
        # label
        label_a = ws_f.cell(row=r, column=1).value
        label_b = ws_f.cell(row=r, column=2).value
        code_c = ws_f.cell(row=r, column=3).value
        
        row_info = {
            'row': r,
            'label_A': str(label_a) if label_a is not None else "",
            'label_B': str(label_b) if label_b is not None else "",
            'code_C': str(code_c) if code_c is not None else "",
            'cells': {f"{get_column_letter(c['col'])}": (c['raw_val'], c['eval_val'], c['fill_color']) for c in r_cells}
        }
        rows_structure.append(row_info)
        
    return {
        'merged_ranges': merged_ranges,
        'cells': cells_data,
        'rows': rows_structure
    }

if __name__ == '__main__':
    res = analyze_sheet_0_m1('MẪU CHUẨN - PCGD 2026(ok).xlsx')
    with open('0_m1_analysis.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print("0.M1 detailed analysis complete.")
