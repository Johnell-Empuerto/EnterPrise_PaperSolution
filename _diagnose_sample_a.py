"""
Diagnose Sample A COM sheet access.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

DOCS = os.path.expanduser("~\\Documents")
SAMPLES = [os.path.join(DOCS, "Sample A.xlsx"),
           os.path.join(DOCS, "Text by HandWriting.xlsx")]

import pythoncom
import win32com.client
import traceback

for path in SAMPLES:
    label = os.path.basename(path)
    print(f"\n{'='*60}")
    print(f"Diagnosing: {label}")
    print(f"Path: {path}")
    
    if not os.path.isfile(path):
        print(f"  FILE NOT FOUND")
        continue
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(path))
            
            # Test 1: wb.Sheets
            print(f"\n  Test 1: wb.Sheets")
            try:
                count = wb.Sheets.Count
                print(f"    Sheets.Count = {count}")
                for i in range(1, count + 1):
                    s = wb.Sheets(i)
                    print(f"    Sheet {i}: name='{s.Name}' visible={s.Visible}")
            except Exception as e:
                print(f"    ERROR: {e}")
                traceback.print_exc()
            
            # Test 2: wb.Worksheets
            print(f"\n  Test 2: wb.Worksheets")
            try:
                count = wb.Worksheets.Count
                print(f"    Worksheets.Count = {count}")
                for i in range(1, count + 1):
                    s = wb.Worksheets(i)
                    print(f"    Worksheet {i}: name='{s.Name}' visible={s.Visible}")
            except Exception as e:
                print(f"    ERROR: {e}")
                traceback.print_exc()
            
            # Test 3: Check for _Fields sheet
            print(f"\n  Test 3: Check _Fields")
            for i in range(1, wb.Sheets.Count + 1):
                try:
                    s = wb.Sheets(i)
                    if s.Name == "_Fields":
                        print(f"    _Fields found! UsedRange rows={s.UsedRange.Rows.Count}")
                except:
                    pass
            try:
                for i in range(1, wb.Worksheets.Count + 1):
                    s = wb.Worksheets(i)
                    if s.Name == "_Fields":
                        print(f"    _Fields (via Worksheets) found!")
            except:
                pass
            
            # Test 4: Comments via Range().Comment on visible sheet
            print(f"\n  Test 4: Per-cell comment scan on visible sheets")
            for i in range(1, wb.Sheets.Count + 1):
                try:
                    ws = wb.Sheets(i)
                except:
                    continue
                if ws.Visible != -1:
                    continue
                if ws.Name == "_Fields":
                    continue
                
                print(f"    Sheet '{ws.Name}': checking for comments...")
                
                # Quick check
                try:
                    cc = ws.Comments.Count
                    print(f"      ws.Comments.Count = {cc}")
                except Exception as e:
                    print(f"      ws.Comments failed: {e}")
                    cc = 0
                
                if cc > 0:
                    try:
                        used = ws.UsedRange
                        rows = used.Rows.Count
                        cols = used.Columns.Count
                        print(f"      UsedRange: {rows}r x {cols}c")
                        for r in range(1, rows + 1):
                            for c in range(1, cols + 1):
                                from render_service.excel_cluster_reader import _col_to_letter
                                addr = f"{_col_to_letter(c)}{r}"
                                try:
                                    cell = ws.Range(addr)
                                    comment = cell.Comment
                                    if comment is not None:
                                        text = str(comment.Text())
                                        print(f"      COMMENT on {addr}: text='{text[:100]}'")
                                except:
                                    pass
                    except Exception as e:
                        print(f"      Per-cell scan error: {e}")
            
        except Exception as e:
            print(f"  ERROR opening workbook: {e}")
            traceback.print_exc()
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()
