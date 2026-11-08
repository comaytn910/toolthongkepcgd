import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.fast_excel import fast_read_sheet
import openpyxl

fpath = "data/41_41. Xã Nghinh Tường - Nhập hoàn chỉnh PCGD 2026.xlsx"
cells = fast_read_sheet(fpath, "1.M7")

ALL_COLS = [openpyxl.utils.get_column_letter(c) for c in range(3, 45)] # C đến AR

print("Row 7 (Tong I+II):")
for col in ALL_COLS:
    v = cells.get(f"{col}7", {}).get('val')
    if v is not None and v != 0 and v != '0':
        print(f"  {col}7 = {v}")

print("\nRow 34 (Cong I):")
for col in ALL_COLS:
    v = cells.get(f"{col}34", {}).get('val')
    if v is not None and v != 0 and v != '0':
        print(f"  {col}34 = {v}")

print("\nRow 68 (Cong II):")
for col in ALL_COLS:
    v = cells.get(f"{col}68", {}).get('val')
    if v is not None and v != 0 and v != '0':
        print(f"  {col}68 = {v}")
