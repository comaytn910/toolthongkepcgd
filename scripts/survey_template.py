import openpyxl
from openpyxl.utils import get_column_letter
import json
import sys

def analyze_workbook(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=False)
    
    # 1. Sheet classification
    business_sheets = {
        '0.M': [f'0.M{i}' for i in range(1, 8)],
        '1.M': [f'1.M{i}' for i in range(1, 8)],
        '2.M': [f'2.M{i}' for i in range(1, 11)],
        '3.M': [f'3.M{i}' for i in range(1, 10)]
    }
    all_business = []
    for grp in business_sheets.values():
        all_business.extend(grp)
        
    supporting_sheets = ['ĐỐI_SOÁT', 'TỔNG_QUAN', 'CAU_HINH_DOI_SOAT', 'DS_SAO_LUU_20261012_180009_1147']
    
    sheet_inventory = []
    
    for s_name in wb.sheetnames:
        ws = wb[s_name]
        is_biz = s_name in all_business
        is_supp = s_name in supporting_sheets or not is_biz
        
        # Find real boundaries
        min_r, max_r = 999999, 0
        min_c, max_c = 999999, 0
        val_count = 0
        formula_count = 0
        
        # Scan cells that have value
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None:
                    val_count += 1
                    r, c = cell.row, cell.column
                    if r < min_r: min_r = r
                    if r > max_r: max_r = r
                    if c < min_c: min_c = c
                    if c > max_c: max_c = c
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        formula_count += 1
                        
        if val_count == 0:
            min_r, max_r, min_c, max_c = 0, 0, 0, 0
            
        # Find title (first non-empty text in first 5 rows)
        title = ""
        for r in range(1, min(max_r + 1, 10)):
            for c in range(1, min(max_c + 1, 10)):
                v = ws.cell(row=r, column=c).value
                if v and isinstance(v, str) and len(v.strip()) > 5:
                    if not v.startswith('='):
                        title = v.strip().replace('\n', ' ')
                        break
            if title:
                break
                
        sheet_inventory.append({
            'name': s_name,
            'is_business': is_biz,
            'is_supporting': is_supp,
            'openpyxl_max_row': ws.max_row,
            'openpyxl_max_col': ws.max_column,
            'real_min_row': min_r,
            'real_max_row': max_r,
            'real_min_col': min_c,
            'real_max_col': max_c,
            'real_max_col_letter': get_column_letter(max_c) if max_c > 0 else '',
            'val_count': val_count,
            'formula_count': formula_count,
            'title': title[:100]
        })
        
    return sheet_inventory

if __name__ == '__main__':
    file_path = 'MẪU CHUẨN - PCGD 2026(ok).xlsx'
    inv = analyze_workbook(file_path)
    with open('inventory_result.json', 'w', encoding='utf-8') as f:
        json.dump(inv, f, ensure_ascii=False, indent=2)
    print(f'Inventory completed for {len(inv)} sheets.')
