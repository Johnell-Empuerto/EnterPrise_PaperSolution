"""
Phase X.10 -- Root Cause Diagnostic: cell.Comment Failure in Single COM Session

Tests every possible cause systematically in a SINGLE subprocess,
SINGLE COM session to eliminate stale-proxy confounding factors.
"""
import os, sys, threading
sys.path.insert(0, r'C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise')
os.chdir(r'C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise')

FORMTEST = r'C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx'
HLINE = "=" * 72
results = []

def section(title):
    print(f"\n{HLINE}")
    print(f"  {title}")
    print(HLINE)

def scan_comments(ws, label):
    """Scan ALL cells for comments using the EXACT same pattern as _identify_clusters_from_comments."""
    found = []
    try:
        used = ws.UsedRange
        lr = used.Row + used.Rows.Count - 1
        lc = used.Column + used.Columns.Count - 1
        for row in range(1, lr + 1):
            for col in range(1, lc + 1):
                cell = ws.Cells(row, col)
                try:
                    comment = cell.Comment
                except Exception:
                    comment = None
                if comment is None:
                    continue
                try:
                    ma = cell.MergeArea
                except Exception:
                    ma = cell
                try:
                    addr = str(ma.Address)
                except Exception:
                    addr = str(cell.Address)
                try:
                    text = str(comment.Text())
                except Exception:
                    text = "<TEXT_FAILED>"
                found.append({
                    "addr": addr,
                    "text_preview": text[:50] if text else "",
                })
    except Exception as e:
        print(f"  SCAN FAILED: {e}")
        return []
    return found

def inspect_cell(ws, row, col, label):
    """Deep-inspect a single cell for comment properties."""
    cell = ws.Cells(row, col)
    info = {"label": label}
    print(f"  [{label}] Row={row} Col={col}")
    
    try:
        comment = cell.Comment
        info["comment_obj"] = repr(comment) if comment is not None else "None"
    except Exception as e:
        info["comment_obj"] = f"EXCEPTION: {e}"
    print(f"    comment_obj: {info['comment_obj']}")
    
    if "None" not in str(info.get("comment_obj", "")):
        try:
            info["comment_text"] = str(comment.Text())
        except Exception as e:
            info["comment_text"] = f"EXCEPTION: {e}"
        print(f"    comment_text: {info['comment_text']}")
        try:
            info["comment_author"] = str(comment.Author)
        except Exception as e:
            info["comment_author"] = f"EXCEPTION: {e}"
        print(f"    comment_author: {info['comment_author']}")
    
    try:
        ma = cell.MergeArea
        info["merge_area"] = str(ma.Address) if ma else "None"
    except Exception as e:
        info["merge_area"] = f"EXCEPTION: {e}"
    print(f"    merge_area: {info['merge_area']}")
    
    return info


# ================================================================
# MAIN: Single COM session, systematic testing
# ================================================================
import pythoncom
import win32com.client

pythoncom.CoInitialize()
print(f"CoInitialize succeeded. Thread ID: {threading.get_ident()}")

excel = win32com.client.Dispatch("Excel.Application")
excel.DisplayAlerts = False
excel.Visible = False
print(f"Excel.Application created: version={excel.Version}")

wb = excel.Workbooks.Open(os.path.abspath(FORMTEST))
ws = wb.Worksheets(1)

# --- Test A: Baseline ---
section("TEST 1: Baseline -- comments immediately after Workbook.Open()")
print(f"  Workbook.FullName: {wb.FullName}")
print(f"  Workbook.ReadOnly: {wb.ReadOnly}")
print(f"  Workbook.Sheets.Count: {wb.Sheets.Count}")

found = scan_comments(ws, "BASELINE")
print(f"  Comments found: {len(found)}")
for f in found:
    print(f"    {f['addr']}: {f['text_preview']}")

if len(found) == 0:
    print("  [CRITICAL] 0 comments found in baseline before any operations.")
else:
    print("  [OK] Comments found in baseline.")

results.append(("Baseline (after Open)", len(found)))

# --- Test B: After clearing shapes ---
section("TEST 2: After clearing shapes")
try:
    for shape in list(ws.Shapes):
        shape.Delete()
    print("  Shapes deleted")
except Exception as e:
    print(f"  Shape deletion: {e}")

found = scan_comments(ws, "AFTER SHAPES")
print(f"  Comments found: {len(found)}")
results.append(("After clearing shapes", len(found)))

# --- Test C: After clearing headers ---
section("TEST 3: After clearing headers/footers")
ws.PageSetup.CenterHeader = ""
ws.PageSetup.CenterFooter = ""
print("  Headers/footers cleared")

found = scan_comments(ws, "AFTER HEADERS")
print(f"  Comments found: {len(found)}")
results.append(("After clearing headers", len(found)))

# --- Test D: After ExportAsFixedFormat ---
section("TEST 4: After ExportAsFixedFormat")
import tempfile
tmp_dir = tempfile.mkdtemp(prefix="ple_X10_")
orig_pdf = os.path.join(tmp_dir, "original.pdf")
try:
    wb.ExportAsFixedFormat(0, os.path.abspath(orig_pdf))
    print(f"  PDF exported to: {orig_pdf}")
except Exception as e:
    print(f"  ExportAsFixedFormat FAILED: {e}")

found = scan_comments(ws, "AFTER PDF EXPORT")
print(f"  Comments found: {len(found)}")

print("  Deep inspection of known comment cells after PDF export:")
inspect_cell(ws, 1, 1, "A1")
inspect_cell(ws, 1, 3, "C1")
inspect_cell(ws, 3, 1, "A3")
inspect_cell(ws, 6, 1, "A6")
inspect_cell(ws, 9, 1, "A9")
inspect_cell(ws, 12, 1, "A12")

results.append(("After PDF export", len(found)))

# --- Test E: After SaveAs ---
section("TEST 5: After SaveAs")
save_dir = tempfile.mkdtemp(prefix="ple_X10_save_")
save_path = os.path.join(save_dir, "test.xlsx")
try:
    wb.SaveAs(os.path.abspath(save_path))
    print(f"  Saved to: {save_path}")
except Exception as e:
    print(f"  SaveAs FAILED: {e}")

found = scan_comments(ws, "AFTER SAVE")
print(f"  Comments found: {len(found)}")
print("  Deep inspection after SaveAs:")
inspect_cell(ws, 1, 1, "A1")
results.append(("After SaveAs", len(found)))

# --- Test F: Worksheet reference consistency ---
section("TEST 6: Worksheet reference consistency")
print(f"  ws.Name: {ws.Name}")
print(f"  ws.Index: {ws.Index}")
try:
    print(f"  ws.CodeName: {ws.CodeName}")
except Exception as e:
    print(f"  ws.CodeName: {e}")
ws2 = wb.Worksheets(1)
print(f"  ws2.Name (from wb.Worksheets(1)): {ws2.Name}")
print(f"  Same object? {ws is ws2}")

# --- Test G: Apartment state ---
section("TEST 7: Thread/apartment state")
print(f"  Thread ID: {threading.get_ident()}")
print(f"  CoInitialize called: YES (at start)")
print(f"  Workbook still accessible: {wb.Name}")

# --- Test H: Exact _identify_clusters code path inlined ---
section("TEST 8: Exact _identify_clusters code path (inlined)")

print("  Running exact _identify_clusters_from_comments logic in THIS session:")
found_inline = []
for ws_i in wb.Worksheets:
    try:
        used = ws_i.UsedRange
        lr = used.Row + used.Rows.Count - 1
        lc = used.Column + used.Columns.Count - 1
        for row in range(1, lr + 1):
            for col in range(1, lc + 1):
                cell = ws_i.Cells(row, col)
                try:
                    comment = cell.Comment
                except Exception:
                    comment = None
                if comment is None:
                    continue
                try:
                    ma = cell.MergeArea
                except Exception:
                    ma = cell
                try:
                    addr = str(ma.Address)
                except Exception:
                    addr = str(cell.Address)
                try:
                    text = str(comment.Text())
                except Exception:
                    text = ""
                lines = text.replace("\r\n", "\n").split("\n")
                found_inline.append({
                    "addr": addr,
                    "name": lines[0].strip() if lines else "",
                    "text_preview": text[:50],
                })
    except Exception as e:
        print(f"  Worksheet {ws_i.Name} scan failed: {e}")

print(f"  Inline scan found: {len(found_inline)} comments")
for f in found_inline:
    print(f"    {f['addr']}: {f['name']}")
results.append(("Inline scan (same session)", len(found_inline)))

# --- Summary ---
section("SUMMARY")
print(f"{'Test':40s} {'Comments':>10s}")
print("-" * 52)
for label, count in results:
    status = "OK" if count > 0 else "FAIL"
    print(f"[{status}] {label:40s} {count:>10d}")

if found_inline:
    print("\n*** ROOT CAUSE FOUND: Inlined comment code WORKS in single session.")
    print("    The previous failure was due to the indentation bug.")
    print("    The 2-session workaround can be eliminated.")
else:
    print("\n*** Inlined comment code STILL fails in single session.")
    print("    The root cause is deeper than the indentation bug.")

# Cleanup
wb.Close(False)
excel.Quit()
pythoncom.CoUninitialize()

import shutil
shutil.rmtree(tmp_dir, ignore_errors=True)
shutil.rmtree(save_dir, ignore_errors=True)
print("\nDone.")
