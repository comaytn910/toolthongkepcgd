import openpyxl
import sys

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    wb = openpyxl.load_workbook('MẪU CHUẨN - PCGD 2026(ok).xlsx', data_only=False)
    ws = wb['0.M1']

    print(f"{'Row':<4} | {'Mã':<4} | {'Tên chỉ tiêu':<42} | {'D(0t)':<10} | {'E(1t)':<10} | {'F(2t)':<10} | {'G(3t)':<10} | {'H(4t)':<10} | {'I(5t)':<10} | {'J(6t)':<10} | {'K(Tổng)':<35}")
    print('-'*150)

    for r in range(3, 30):
        code = str(ws[f'C{r}'].value or '')
        if code.endswith('.0'): code = code[:-2]
        name_a = str(ws[f'A{r}'].value or '').strip()
        name_b = str(ws[f'B{r}'].value or '').strip()
        name = (name_a + ' - ' + name_b).strip(' -')
        
        vals = {}
        for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            cell = ws[f'{c}{r}']
            v = cell.value
            color_tag = ''
            if cell.fill and cell.fill.start_color:
                rgb = str(cell.fill.start_color.rgb or cell.fill.start_color.theme or '')
                if 'FFFF' in rgb or 'FFEB' in rgb or 'FFF2' in rgb:
                    color_tag = '[Y]'
            val_str = str(v) if v is not None else ''
            if val_str.startswith('='):
                vals[c] = f'f:{val_str}'
            elif val_str == 'x':
                vals[c] = 'x'
            else:
                vals[c] = f'{val_str}{color_tag}'
                
        print(f"R{r:<3} | {code:<4} | {name[:42]:<42} | {vals['D']:<10} | {vals['E']:<10} | {vals['F']:<10} | {vals['G']:<10} | {vals['H']:<10} | {vals['I']:<10} | {vals['J']:<10} | {vals['K']:<35}")

if __name__ == '__main__':
    main()
