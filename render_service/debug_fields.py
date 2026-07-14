"""Debug: read ALL rows from _Fields sheet via COM and via XML."""
import sys, os, win32com.client
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from render_service.excel_cluster_reader import read_fields

wb = "[V3.1_Sample]アンケート用紙.xlsx"

# Method 1: ExcelClusterReader
print("=== ExcelClusterReader ===")
fields = read_fields(wb)
print(f"Fields: {len(fields)}")
for f in fields:
    print(f"  {f.addr}: ratios=({f.ratio_left:.6f},{f.ratio_top:.6f},{f.ratio_right:.6f},{f.ratio_bottom:.6f})")

# Method 2: Direct COM dump
print("\n=== Direct COM Dump ===")
excel = win32com.client.Dispatch("Excel.Application")
excel.DisplayAlerts = False
excel.Visible = False
try:
    wb_com = excel.Workbooks.Open(os.path.abspath(wb))
    
    sheet = None
    for i in range(1, wb_com.Sheets.Count + 1):
        if wb_com.Sheets(i).Name == "_Fields":
            sheet = wb_com.Sheets(i)
            print(f"_Fields sheet: Visible={sheet.Visible}")
            break
    
    if sheet:
        # Try multiple ways to count rows
        last_row_a = sheet.UsedRange.Rows.Count
        print(f"UsedRange.Rows.Count = {last_row_a}")
        
        # Try finding last row from column A specifically
        try:
            last_row_b = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
            print(f"xlUp from bottom in col A = {last_row_b}")
        except:
            pass
        
        # Read all cells from column A
        for row in range(1, 20):
            val = sheet.Cells(row, 1).Value
            if val:
                vals = []
                for col in range(1, 8):
                    v = sheet.Cells(row, col).Value
                    vals.append(str(v)[:60] if v is not None else "None")
                print(f"  Row {row}: {' | '.join(vals)}")
            else:
                print(f"  Row {row}: (empty)")
finally:
    excel.Quit()
