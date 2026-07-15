"""Test basic Excel COM operations."""
import sys, os
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

import win32com.client

DOCS = r"C:\Users\MCF-JO~1\Documents"
FORM = os.path.join(DOCS, "FormTest - Copy.xlsx")
JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n=== Testing {wb_name} ===")
    if not os.path.exists(wb_path):
        print(f"  NOT FOUND: {wb_path}")
        continue
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    
    try:
        wb = excel.Workbooks.Open(os.path.abspath(wb_path))
        print(f"  Opened: {wb.Name}")
        
        for ws_idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(ws_idx)
            print(f"  Sheet {ws_idx}: {ws.Name}")
            
            used = ws.UsedRange
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            print(f"    UsedRange: rows=1..{lr}, cols=1..{lc}")
            
            # Try reading first 5 cells
            comment_count = 0
            for row in range(1, min(lr + 1, 5)):
                for col in range(1, min(lc + 1, 5)):
                    try:
                        cell = ws.Cells(row, col)
                        comment = cell.Comment
                        if comment is not None:
                            comment_count += 1
                            text = str(comment.Text())
                            print(f"    [{row},{col}] Comment: {text[:60]}...")
                    except Exception as e:
                        print(f"    [{row},{col}] Error: {e}")
            
            print(f"    Total comments in scanned area: {comment_count}")
        
        wb.Close(False)
        print(f"  Closed OK")
    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        excel.Quit()
        print(f"  Excel quit OK")

print("\nDone")
