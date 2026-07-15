"""
Stage 2: PDF dimension comparison and _Fields sheet analysis.
Uses separate COM sessions to avoid threading issues.
"""
import sys, os, gc
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)
_ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
sys.path.insert(0, _ROOT)
DOCS = r"C:\Users\MCF-JO~1\Documents"

# ── 1. PDF Dimensions: Original workbook ──
print("=" * 60)
print("  PDF DIMENSIONS: ORIGINAL WORKBOOK")
print("=" * 60)

from render_service.pdf_converter import xlsx_to_pdf
from render_service.background_renderer import get_page_dimensions
from render_service.upload_coordinate_generator import DPI

for name, path in [("FormTest", DOCS + "\\FormTest - Copy.xlsx"), 
                   ("Japanese", DOCS + "\\[V3.1_Sample]アンケート用紙.xlsx")]:
    try:
        pdf_path = xlsx_to_pdf(path)
        w, h = get_page_dimensions(pdf_path, dpi=DPI)
        print(f"  {name}: {w}x{h} px ({w/(DPI/72):.0f}x{h/(DPI/72):.0f} pt)")
        import shutil
        try: os.unlink(pdf_path)
        except: pass
        d = os.path.dirname(pdf_path)
        if d: shutil.rmtree(d, ignore_errors=True)
    except Exception as e:
        print(f"  {name}: ERROR: {e}")
        import traceback; traceback.print_exc()

# ── 2. _Fields sheet analysis for Japanese workbook ──
print("\n" + "=" * 60)
print("  _FIELDS SHEET ANALYSIS")
print("=" * 60)

import win32com.client

for name, path in [("FormTest", DOCS + "\\FormTest - Copy.xlsx"),
                   ("Japanese", DOCS + "\\[V3.1_Sample]アンケート用紙.xlsx")]:
    print(f"\n  --- {name} ---")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        fields_sheet = None
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == "_Fields":
                fields_sheet = wb.Sheets(i)
                break
        if fields_sheet is None:
            print(f"  No _Fields sheet found")
        else:
            lr = fields_sheet.UsedRange.Rows.Count
            lc = fields_sheet.UsedRange.Columns.Count
            print(f"  _Fields sheet: {lr} rows x {lc} cols")
            for row in range(1, lr + 1):
                row_data = []
                for col in range(1, min(lc + 1, 7)):
                    try:
                        v = fields_sheet.Cells(row, col).Value
                        if v is None:
                            row_data.append("(null)")
                        else:
                            s = str(v).strip()
                            if len(s) > 40: s = s[:37] + "..."
                            row_data.append(s)
                    except Exception as e:
                        row_data.append(f"(err:{e})")
                print(f"    Row {row}: {' | '.join(row_data)}")
        wb.Close(False)
    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        excel.Quit()
    gc.collect()

# ── 3. Page setup comparison ──
print("\n" + "=" * 60)
print("  PAGE SETUP COMPARISON")
print("=" * 60)

for name, path in [("FormTest", DOCS + "\\FormTest - Copy.xlsx"),
                   ("Japanese", DOCS + "\\[V3.1_Sample]アンケート用紙.xlsx")]:
    print(f"\n  --- {name} ---")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        ps = ws.PageSetup
        print(f"  Sheet: {ws.Name}")
        for prop_name in ["Zoom", "FitToPagesWide", "FitToPagesTall",
                          "Orientation", "PaperSize",
                          "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
                          "CenterHorizontally", "CenterVertically",
                          "PageSetup.PageWidth", "PageSetup.PageHeight",
                          "Order"]:
            try:
                val = getattr(ps, prop_name) if not "." in prop_name else None
                if "." in prop_name:
                    parts = prop_name.split(".")
                    obj = ps
                    for p in parts:
                        obj = getattr(obj, p)
                    val = obj
                print(f"    {prop_name}: {val}")
            except Exception as e:
                print(f"    {prop_name}: ERROR: {e}")
        wb.Close(False)
    except Exception as e:
        print(f"  ERROR: {e}")
    finally:
        excel.Quit()
    gc.collect()

print("\nDone")
