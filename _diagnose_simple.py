"""
Simple diagnostic to avoid segfault - minimal COM operations.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

DOCS = os.path.expanduser("~\\Documents")
paths = [
    ("Sample A", os.path.join(DOCS, "Sample A.xlsx")),
    ("Text by HandWriting", os.path.join(DOCS, "Text by HandWriting.xlsx")),
]

import pythoncom
import win32com.client
import traceback

for label, path in paths:
    print(f"\n=== {label} ===")
    print(f"Path: {path}")
    print(f"Exists: {os.path.isfile(path)}")
    
    if not os.path.isfile(path):
        print(f"  FILE NOT FOUND — skip")
        continue
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        try:
            wb = excel.Workbooks.Open(os.path.abspath(path))
            
            # Minimal: just count sheets and get names
            try:
                count = wb.Sheets.Count
                print(f"  Sheets.Count = {count}")
                for i in range(1, min(count + 1, 5)):  # first 5
                    try:
                        s = wb.Sheets(i)
                        print(f"    Sheet {i}: '{s.Name}' vis={s.Visible}")
                    except Exception as e:
                        print(f"    Sheet {i}: ERROR {e}")
            except Exception as e:
                print(f"  Sheets failed: {e}")
            
            # Try worksheets
            try:
                count = wb.Worksheets.Count
                print(f"  Worksheets.Count = {count}")
            except Exception as e:
                print(f"  Worksheets failed: {e}")
            
            # Try reading comments from first visible sheet
            try:
                ws = wb.Worksheets(1)
                cc = ws.Comments.Count
                print(f"  Sheet1.Comments.Count = {cc}")
                if cc > 0:
                    used = ws.UsedRange
                    print(f"  UsedRange: {used.Rows.Count}r x {used.Columns.Count}c")
            except Exception as e:
                print(f"  Comments check failed: {e}")
            
            wb.Close(False)
            
        except Exception as e:
            print(f"  OPEN ERROR: {type(e).__name__}: {e}")
        finally:
            try:
                excel.Quit()
            except:
                pass
    finally:
        pythoncom.CoUninitialize()
