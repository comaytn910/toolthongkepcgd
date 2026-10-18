import os
import sys
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path: sys.path.insert(0, BASE_DIR)
if sys.stdout.encoding != 'utf-8':
    try: sys.stdout.reconfigure(encoding='utf-8')
    except Exception: pass
from core.fast_excel import fast_read_sheet

data_dir = os.path.join(BASE_DIR, 'data')
files = [os.path.join(data_dir, f) for f in sorted(os.listdir(data_dir)) if f.endswith('.xlsx') and not f.startswith('~$')]

print("Kiểm tra công thức ô Q8 và Q9 ở 5 xã đầu tiên:")
for f in files[:5]:
    c_name = os.path.basename(f).split(' - ')[0]
    cells = fast_read_sheet(f, '0.M3')
    q8_f = cells.get('Q8', {}).get('formula')
    q8_v = cells.get('Q8', {}).get('val')
    k8 = cells.get('K8', {}).get('val')
    o8 = cells.get('O8', {}).get('val')
    p8 = cells.get('P8', {}).get('val')
    f8 = cells.get('F8', {}).get('val')
    print(f"Xã: {c_name}")
    print(f"  Hàng 8: K8={k8}, O8={o8}, P8={p8}, O+P={float(o8 or 0)+float(p8 or 0)}, F8={f8}, Q8_val={q8_v}, Q8_formula={q8_f}")
    
    q9_f = cells.get('Q9', {}).get('formula')
    q9_v = cells.get('Q9', {}).get('val')
    k9 = cells.get('K9', {}).get('val')
    o9 = cells.get('O9', {}).get('val')
    p9 = cells.get('P9', {}).get('val')
    print(f"  Hàng 9: K9={k9}, O9={o9}, P9={p9}, O+P={float(o9 or 0)+float(p9 or 0)}, Q9_val={q9_v}, Q9_formula={q9_f}")
