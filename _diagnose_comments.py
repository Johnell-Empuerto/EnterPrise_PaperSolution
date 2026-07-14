"""
Diagnostic script: Test COM comment reading on FormTest - Copy.xlsx.
This bypasses the full pipeline to isolate the comment reading issue.
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

WORKBOOK = os.path.abspath("FormTest - Copy.xlsx")
print(f"Workbook: {WORKBOOK}")
print(f"File exists: {os.path.isfile(WORKBOOK)}")

import pythoncom
import win32com.client

pythoncom.CoInitialize()
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(WORKBOOK)
        
        print(f"\n=== Sheet inventory ===")
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            print(f"  Sheet {i}: name='{ws.Name}' visible={ws.Visible}")
        
        # Check _Fields sheet
        print(f"\n=== _Fields sheet check ===")
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            if ws.Name == "_Fields":
                used = ws.UsedRange
                rows = used.Rows.Count
                cols = used.Columns.Count
                print(f"  _Fields: UsedRange rows={rows}, cols={cols}")
                # Check if any non-header row has data
                has_data = False
                for r in range(2, rows + 1):
                    for c in range(1, cols + 1):
                        val = ws.Cells(r, c).Value
                        if val is not None:
                            has_data = True
                            print(f"  Row {r}, Col {c}: value='{val}'")
                if not has_data:
                    print(f"  _Fields has NO data rows (rows={rows}, but all empty)")
        
        # Check comments on each visible data sheet
        print(f"\n=== Comment reading diagnostics ===")
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            if ws.Visible != -1:  # skip hidden
                print(f"  Sheet '{ws.Name}' is hidden, skipping")
                continue
            if ws.Name == "_Fields":
                continue
            
            print(f"\n  --- Sheet '{ws.Name}' ---")
            
            # Method 1: ws.Comments
            try:
                comments = ws.Comments
                if comments is None:
                    print(f"  [Method 1] ws.Comments = None")
                else:
                    count = comments.Count
                    print(f"  [Method 1] ws.Comments count = {count}")
                    if count > 0:
                        for ci in range(1, count + 1):
                            try:
                                c = comments(ci)
                                addr = c.Parent.Address(False, False)
                                text = str(c.Text())
                                print(f"    Comment on {addr}: text='{text[:100]}'")
                            except Exception as e:
                                print(f"    Error reading comment {ci}: {e}")
            except Exception as e:
                print(f"  [Method 1] ws.Comments FAILED: {e}")
            
            # Method 2: Per-cell cell.Comment via UsedRange
            print(f"  [Method 2] Per-cell cell.Comment scan...")
            try:
                used = ws.UsedRange
                found = 0
                for r in range(1, used.Rows.Count + 1):
                    for c in range(1, used.Columns.Count + 1):
                        cell = ws.Cells(r, c)
                        try:
                            comment = cell.Comment
                            if comment is not None:
                                addr = cell.Address(False, False)
                                text = str(comment.Text())
                                print(f"    cell.Comment on {addr}: text='{text[:100]}'")
                                found += 1
                        except:
                            pass
                if found == 0:
                    print(f"    No cell.Comment found via per-cell scan")
            except Exception as e:
                print(f"  [Method 2] Per-cell scan FAILED: {e}")
            
            # Method 3: Check if used range has content
            print(f"  [Method 3] UsedRange dimensions...")
            try:
                used = ws.UsedRange
                print(f"    Rows={used.Rows.Count}, Cols={used.Columns.Count}")
                print(f"    Address={used.Address(False, False)}")
                for r in range(1, min(used.Rows.Count + 1, 20)):  # first 20 rows
                    row_vals = []
                    for c in range(1, min(used.Columns.Count + 1, 5)):  # first 5 cols
                        val = ws.Cells(r, c).Value
                        if val is not None:
                            row_vals.append(f"Col{c}={val}")
                    if row_vals:
                        print(f"    Row {r}: {', '.join(row_vals)}")
            except Exception as e:
                print(f"    Error: {e}")
        
        print(f"\n=== Diagnosis complete ===")
        
    finally:
        excel.Quit()
finally:
    pythoncom.CoUninitialize()
