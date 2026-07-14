"""
Deep diagnostic: Test different COM comment access patterns on FormTest workbook.
"""

import os, sys
sys.stdout.reconfigure(encoding='utf-8')

WORKBOOK = os.path.abspath("FormTest - Copy.xlsx")
print(f"Workbook: {WORKBOOK}")

import pythoncom
import win32com.client

pythoncom.CoInitialize()
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(WORKBOOK)
        
        # Find Sheet1
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            s = wb.Sheets(i)
            if s.Name == "Sheet1":
                ws = s
                break
        
        if ws is None:
            print("Sheet1 not found!")
            sys.exit(1)
        
        print(f"\n=== Sheet1 comment diagnostics ===\n")
        
        # Test 1: Try to read comment from specific known cell
        print("Test 1: ws.Range(\"A1\").Comment")
        try:
            rng = ws.Range("A1")
            print(f"  Range A1 exists: {rng is not None}")
            comment = rng.Comment
            print(f"  rng.Comment type: {type(comment)}")
            print(f"  rng.Comment value: {comment}")
            if comment is not None:
                text = str(comment.Text())
                print(f"  Comment text: '{text}'")
        except Exception as e:
            print(f"  ERROR: {e}")
        
        # Test 2: Try cell.Comment instead
        print("\nTest 2: ws.Cells(1,1).Comment")
        try:
            cell = ws.Cells(1, 1)
            print(f"  Cells(1,1) type: {type(cell)}")
            print(f"  Cells(1,1) address: {cell.Address(False, False)}")
            comment = cell.Comment
            print(f"  cell.Comment type: {type(comment)}")
            print(f"  cell.Comment value: {comment}")
            if comment is not None:
                text = str(comment.Text())
                print(f"  Comment text: '{text}'")
                print(f"  repr(text): {repr(text)}")
        except Exception as e:
            print(f"  ERROR: {e}")
        
        # Test 3: Try ws.Comments iteration
        print("\nTest 3: For loop over ws.Comments")
        try:
            comments = ws.Comments
            print(f"  ws.Comments type: {type(comments)}")
            print(f"  ws.Comments count: {comments.Count if comments else 'N/A'}")
            if comments and comments.Count > 0:
                for ci, comment in enumerate(comments):
                    try:
                        print(f"  Comment #{ci}: type={type(comment)}")
                        parent = comment.Parent
                        print(f"    Parent type: {type(parent)}")
                        addr = parent.Address(False, False)
                        text = str(comment.Text())
                        print(f"    addr={addr}, text='{text[:100]}'")
                    except Exception as e:
                        print(f"    Error accessing comment #{ci}: {e}")
        except Exception as e:
            print(f"  ERROR: {e}")
        
        # Test 4: Try comments.Item
        print("\nTest 4: comments.Item(1)")
        try:
            comments = ws.Comments
            comment = comments.Item(1)
            print(f"  comments.Item(1) type: {type(comment)}")
            addr = comment.Parent.Address(False, False)
            text = str(comment.Text())
            print(f"  addr={addr}, text='{text[:100]}'")
        except Exception as e:
            print(f"  ERROR: {e}")
        
        # Test 5: Check Range("A1") value
        print("\nTest 5: Cell values around comment cells")
        for cell_ref in ["A1", "B1", "C1", "A3", "A6", "A9", "A12"]:
            try:
                rng = ws.Range(cell_ref)
                val = rng.Value
                print(f"  {cell_ref}: value={repr(val)}")
            except Exception as e:
                print(f"  {cell_ref}: ERROR: {e}")
        
        # Test 6: Try using win32com.CastTo to get proper interface
        print("\nTest 6: win32com.CastTo on Comments")
        try:
            comments = ws.Comments
            # Try accessing via indexing
            comment = comments[0]  # try 0-based indexing
            print(f"  comments[0] type: {type(comment)}")
            addr = comment.Parent.Address(False, False)
            text = str(comment.Text())
            print(f"  addr={addr}, text='{text[:100]}'")
        except Exception as e:
            print(f"  [0] ERROR: {e}")
            try:
                comment = comments(1)  # try call syntax
                print(f"  comments(1) type: {type(comment)}")
                addr = comment.Parent.Address(False, False)
                text = str(comment.Text())
                print(f"  addr={addr}, text='{text[:100]}'")
            except Exception as e2:
                print(f"  (1) ERROR: {e2}")
        
        print("\n=== Diagnostic complete ===")
        
    finally:
        excel.Quit()
finally:
    pythoncom.CoUninitialize()
