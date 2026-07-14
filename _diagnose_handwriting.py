"""
Targeted diagnostic for Text by HandWriting.xlsx only.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pythoncom
import win32com.client

path = os.path.expanduser("~\\Documents\\Text by HandWriting.xlsx")
print(f"Testing: {path}")
print(f"File exists: {os.path.isfile(path)}")

pythoncom.CoInitialize()
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        print(f"Workbook opened successfully")
        
        # Test Sheets vs Worksheets
        print(f"\nSheets count: {wb.Sheets.Count}")
        print(f"Worksheets count: {wb.Worksheets.Count}")
        
        for i in range(1, wb.Sheets.Count + 1):
            s = wb.Sheets(i)
            print(f"  Sheet {i}: '{s.Name}' visible={s.Visible} type={type(s).__name__}")
        
        # Get first visible non-_Fields sheet
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            s = wb.Sheets(i)
            if s.Visible == -1 and s.Name != "_Fields":
                ws = s
                print(f"\nUsing sheet: '{s.Name}'")
                break
        
        if ws:
            # Check comments count
            try:
                cc = ws.Comments.Count
                print(f"ws.Comments.Count = {cc}")
            except Exception as e:
                print(f"ws.Comments failed: {e}")
                cc = 0
            
            if cc > 0:
                # Scan used range
                used = ws.UsedRange
                rows = used.Rows.Count
                cols = used.Columns.Count
                print(f"UsedRange: {rows}r x {cols}c, address={used.Address(False, False)}")
                
                # Try specific cell C5
                print(f"\nTrying Range('C5').Comment...")
                try:
                    cell = ws.Range("C5")
                    print(f"  Range C5 exists: {cell is not None}")
                    print(f"  Range C5 value: {cell.Value}")
                    comment = cell.Comment
                    print(f"  Comment: {comment}")
                    if comment is not None:
                        text = str(comment.Text())
                        print(f"  Comment text: {repr(text)}")
                    else:
                        print(f"  No comment on C5")
                except Exception as e:
                    print(f"  C5 failed: {e}")
                
                # Scan all cells
                print(f"\nPer-cell scan:")
                found = 0
                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        from render_service.excel_cluster_reader import _col_to_letter
                        addr = f"{_col_to_letter(c)}{r}"
                        try:
                            cell_range = ws.Range(addr)
                            comment = cell_range.Comment
                            if comment is not None:
                                text = str(comment.Text())
                                print(f"  COMMENT on {addr}: {repr(text)[:100]}")
                                found += 1
                        except Exception as e:
                            pass
                print(f"  Total comments found: {found}")
            else:
                print(f"No comments detected on sheet")
    except Exception as e:
        print(f"OPEN ERROR: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            excel.Quit()
        except:
            pass
finally:
    pythoncom.CoUninitialize()
