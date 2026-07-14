# -*- coding: utf-8 -*-
"""
Phase 11J.18 — Reconstruct the Original ConMas Generation Pipeline
Forensic investigation using only database-stored artifacts.
No runtime engine modifications, no heuristics, no pixel calibration.
"""

import psycopg2, re, os, json, hashlib, shutil, struct
from io import BytesIO
from collections import defaultdict
from xml.etree import ElementTree as ET
from datetime import datetime
import fitz, zipfile  # PyMuPDF

DB_CONFIG = dict(host="127.0.0.1", port=5432, dbname="irepodb",
                 user="postgres", password="cimtops")
FID = 546
BASE_DIR = os.getcwd()

OUT = os.path.join(BASE_DIR, "Investigation_546")
PDF_DIR = os.path.join(BASE_DIR, "pdf_extracts")
os.makedirs(OUT, exist_ok=True)

# ─── Helper: checksums ───────────────────────────────────────────────────────
def sha256_of(data):
    return hashlib.sha256(data).hexdigest()

def md5_of(data):
    return hashlib.md5(data).hexdigest()

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION A — Database Artifact Integrity
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_a():
    print("=" * 70)
    print("INVESTIGATION A — Fresh Database Extraction")
    print("=" * 70)
    
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    # Fetch def_top row
    cur.execute("SELECT * FROM def_top WHERE def_top_id = %s", (FID,))
    cols = [desc[0] for desc in cur.description]
    row = cur.fetchone()
    if not row:
        print("ERROR: def_top_id=%d not found!" % FID)
        return
    
    row_dict = dict(zip(cols, row))
    
    # Extract BLOBs
    def_file_data = bytes(row_dict.get('def_file') or b'')
    bg_file_data = bytes(row_dict.get('background_image_file') or b'')
    thumb_file_data = bytes(row_dict.get('thumbnail_file') or b'')
    xml_data_str = str(row_dict.get('xml_data') or '')
    
    # Write files
    xlsx_path = os.path.join(OUT, "original.xlsx")
    with open(xlsx_path, 'wb') as f: f.write(def_file_data)
    
    pdf_path = os.path.join(OUT, "background.pdf")
    with open(pdf_path, 'wb') as f: f.write(bg_file_data)
    
    thumb_path = os.path.join(OUT, "thumbnail.png")
    with open(thumb_path, 'wb') as f: f.write(thumb_file_data)
    
    xml_path = os.path.join(OUT, "xml_data.xml")
    with open(xml_path, 'w', encoding='utf-8') as f: f.write(xml_data_str)
    
    # Checksums
    artifacts = {
        "original.xlsx": {"sha256": sha256_of(def_file_data), "md5": md5_of(def_file_data), "size": len(def_file_data)},
        "background.pdf": {"sha256": sha256_of(bg_file_data), "md5": md5_of(bg_file_data), "size": len(bg_file_data)},
        "thumbnail.png": {"sha256": sha256_of(thumb_file_data), "md5": md5_of(thumb_file_data), "size": len(thumb_file_data)},
        "xml_data.xml": {"sha256": sha256_of(xml_data_str.encode('utf-8')), "md5": md5_of(xml_data_str.encode('utf-8')), "size": len(xml_data_str)},
    }
    
    # Fetch def_cluster data
    cur.execute("""
        SELECT dsh.def_sheet_id, dsh.def_sheet_name, dsh.def_sheet_no,
               dc.cluster_id, dc.cell_addr,
               dc.left_position, dc.right_position, dc.top_position, dc.bottom_position,
               1 as def_cluster_id
        FROM def_cluster dc
        JOIN def_sheet dsh ON dc.def_sheet_id = dsh.def_sheet_id
        WHERE dsh.def_top_id = %s
        ORDER BY dsh.def_sheet_no, dc.cluster_id
    """, (FID,))
    
    clusters = []
    for r in cur.fetchall():
        clusters.append({
            "def_sheet_id": r[0], "sheet_name": r[1], "sheet_no": r[2],
            "cluster_id": r[3], "cell_addr": r[4],
            "left_position": str(r[5]), "right_position": str(r[6]),
            "top_position": str(r[7]), "bottom_position": str(r[8]),
            "def_cluster_id": 0,
            "_left_type": type(r[5]).__name__,
            "_right_type": type(r[6]).__name__,
            "_top_type": type(r[7]).__name__,
            "_bottom_type": type(r[8]).__name__,
        })
    
    # Fetch def_sheet data
    cur.execute("SELECT * FROM def_sheet WHERE def_top_id = %s", (FID,))
    sheet_cols = [desc[0] for desc in cur.description]
    sheets = []
    for r in cur.fetchall():
        sheets.append(dict(zip(sheet_cols, r)))
    
    # Fetch def_top_size
    cur.execute("SELECT * FROM def_top_size WHERE def_top_id = %s", (FID,))
    size_cols = [desc[0] for desc in cur.description]
    sizes = []
    for r in cur.fetchall():
        sizes.append(dict(zip(size_cols, r)))
    
    conn.close()
    
    # Build metadata
    metadata = {
        "investigation": "Phase 11J.18 — Source of Truth Verification",
        "timestamp": datetime.now().isoformat(),
        "def_top_id": FID,
        "def_top_row": {k: str(v)[:200] for k, v in row_dict.items()},
        "artifacts": artifacts,
        "clusters": clusters,
        "sheets": sheets,
        "def_top_size": sizes,
        "note_ratio_types": "All def_cluster position columns are VARCHAR in DB schema"
    }
    
    meta_path = os.path.join(OUT, "metadata.json")
    with open(meta_path, 'w') as f:
        json.dump(metadata, f, indent=2, default=str)
    
    print("\nExtracted artifacts:")
    for name, info in artifacts.items():
        print(f"  {name}: {info['size']} bytes  SHA256={info['sha256'][:16]}...")
    
    print("\nClusters from database:")
    for c in clusters:
        print(f"  cluster_id={c['cluster_id']} addr={c['cell_addr']}")
        print(f"    L={c['left_position']} R={c['right_position']} T={c['top_position']} B={c['bottom_position']}")
        print(f"    types: L={c['_left_type']} R={c['_right_type']} T={c['_top_type']} B={c['_bottom_type']}")
    
    print("\nInvestigation A complete. Files written to:", OUT)
    return metadata

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION B — Workbook Fingerprint via OpenXML
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_b():
    """Read original.xlsx using zipfile + XML parsing to dump all print-related properties."""
    print("\n" + "=" * 70)
    print("INVESTIGATION B — Workbook Fingerprint (OpenXML)")
    print("=" * 70)
    
    xlsx_path = os.path.join(OUT, "original.xlsx")
    if not os.path.exists(xlsx_path):
        print("ERROR: original.xlsx not found. Run Investigation A first.")
        return
    
    zf = zipfile.ZipFile(xlsx_path, 'r')
    file_list = zf.namelist()
    print("\nXLSX internal files (%d total):" % len(file_list))
    for fn in sorted(file_list):
        print(f"  {fn}")
    
    # Read workbook.xml
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    ws_xml_raw = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
    styles_xml = zf.read("xl/styles.xml").decode("utf-8") if "xl/styles.xml" in file_list else ""
    
    # Parse workbook for sheets
    sheet_names = []
    for m in re.finditer(r'<sheet[^>]*name="([^"]*)"[^>]*/>', wb_xml):
        sheet_names.append(m.group(1))
    # try alternative pattern
    if not sheet_names:
        for m in re.finditer(r'<sheet[^>]*name="([^"]*)"', wb_xml):
            sheet_names.append(m.group(1))
    
    print("\nSheets:", sheet_names)
    
    # Parse defined names (including PrintArea, PrintTitles)
    defined_names = []
    for m in re.finditer(r'<definedName[^>]*>([^<]+)</definedName>', wb_xml):
        defined_names.append(m.group(0))
    
    print("\nDefined Names / PrintArea:")
    for dn in defined_names:
        print(f"  {dn[:200]}")
    
    # Parse worksheet for print-related properties
    # PrintArea from sheet
    pa = re.search(r'<printArea[^>]*>([^<]+)</printArea>', ws_xml_raw)
    print("\nPrintArea (from worksheet XML):", pa.group(1) if pa else "NOT SET")
    
    # PageMargins
    pm = re.search(r'<pageMargins\s+([^>]+)/>', ws_xml_raw)
    print("\nPageMargins:", pm.group(1) if pm else "NOT FOUND")
    
    # PageSetup
    ps = re.search(r'<pageSetup\s+([^>]+)/>', ws_xml_raw)
    print("\nPageSetup:", ps.group(1) if ps else "NOT FOUND")
    
    # SheetPr (page setup properties)
    sp = re.search(r'<sheetPr[^>]+/>', ws_xml_raw)
    print("\nSheetPr:", sp.group(0) if sp else "NOT FOUND")
    
    # Column widths
    cols_section = re.search(r'<cols>(.*?)</cols>', ws_xml_raw, re.DOTALL)
    print("\nColumns:")
    total_col_width = 0
    if cols_section:
        for m in re.finditer(r'<col\s+([^>]+)/?>', cols_section.group(1)):
            attrs = m.group(1)
            min_c = int(re.search(r'min="(\d+)"', attrs).group(1))
            max_c = int(re.search(r'max="(\d+)"', attrs).group(1))
            w = float(re.search(r'width="([\d.]+)"', attrs).group(1))
            hidden = 'hidden="1"' in attrs
            custom = 'customWidth="1"' in attrs
            total_col_width += w * (max_c - min_c + 1)
            print(f"  col {min_c}-{max_c}: width={w} hidden={hidden} customWidth={custom}")
    
    # Row heights
    rows = re.findall(r'<row\s+([^>]+?)/?>', ws_xml_raw)
    print("\nRows:")
    for r_text in rows:
        if 'r="' in r_text and 'ht="' in r_text:
            r_num = int(re.search(r'r="(\d+)"', r_text).group(1))
            ht = float(re.search(r'ht="([\d.]+)"', r_text).group(1))
            hidden = 'hidden="1"' in r_text
            print(f"  row {r_num}: height={ht} hidden={hidden}")
    
    # Merged cells
    mc = re.search(r'<mergeCells[^>]*>(.*?)</mergeCells>', ws_xml_raw, re.DOTALL)
    print("\nMerged Cells:")
    if mc:
        for m in re.finditer(r'<mergeCell ref="([^"]+)"', mc.group(1)):
            print(f"  {m.group(1)}")
    
    # Sheet format pr (fitToPage, etc.)
    sfp = re.search(r'<sheetFormatPr[^>]+/>', ws_xml_raw)
    print("\nSheetFormatPr:", sfp.group(0) if sfp else "NOT FOUND")
    
    # Check for hidden sheets (veryHidden)
    all_sheets_wb = re.findall(r'<sheet[^>]+>', wb_xml)
    print("\nAll sheets with state:")
    for s in all_sheets_wb:
        state = re.search(r'state="([^"]+)"', s)
        state_v = state.group(1) if state else "visible"
        name = re.search(r'name="([^"]+)"', s)
        name_v = name.group(1) if name else "unknown"
        print(f"  {name_v}: state={state_v}")
    
    zf.close()
    
    workbook_dump = {
        "sheets": sheet_names,
        "defined_names": defined_names,
        "print_area": pa.group(1) if pa else None,
        "page_margins_raw": pm.group(1) if pm else None,
        "page_setup_raw": ps.group(1) if ps else None,
        "worksheet_xml_snippet": ws_xml_raw[:3000],
    }
    
    dump_path = os.path.join(OUT, "workbook_dump.json")
    with open(dump_path, 'w') as f:
        json.dump(workbook_dump, f, indent=2)
    
    print("\nInvestigation B complete.")
    return workbook_dump

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION C — Fresh Excel Export
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_c():
    """Open original.xlsx via Excel COM and re-export to PDF. No modifications."""
    print("\n" + "=" * 70)
    print("INVESTIGATION C — Fresh Excel PDF Export")
    print("=" * 70)
    
    xlsx_path = os.path.join(OUT, "original.xlsx")
    reexport_path = os.path.join(OUT, "reexport.pdf")
    
    if not os.path.exists(xlsx_path):
        print("ERROR: original.xlsx not found.")
        return
    
    try:
        import win32com.client
        from win32com.client import constants
        import pythoncom
    except ImportError:
        print("ERROR: win32com not available. Install with: pip install pywin32")
        return
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    page_setup_before = {}
    page_setup_after = {}
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print("\nOpening workbook...")
        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets("Sheet1")
        
        # Dump ALL PageSetup properties BEFORE export
        ps = ws.PageSetup
        page_setup_props = [
            'PaperSize', 'Orientation', 'Zoom', 'FitToPagesWide', 'FitToPagesTall',
            'CenterHorizontally', 'CenterVertically', 'LeftMargin', 'RightMargin',
            'TopMargin', 'BottomMargin', 'HeaderMargin', 'FooterMargin',
            'PrintHeadings', 'PrintGridlines', 'PrintQuality', 'Order',
            'FirstPageNumber', 'BlackAndWhite', 'Draft',
            'LeftHeader', 'CenterHeader', 'RightHeader',
            'LeftFooter', 'CenterFooter', 'RightFooter',
            'PrintTitleRows', 'PrintTitleColumns',
            'PageSetup.Pages.Count'
        ]
        
        print("\nPageSetup BEFORE export:")
        for prop in page_setup_props:
            try:
                val = getattr(ps, prop.split('.')[0]) if '.' not in prop else None
                if '.' in prop:
                    obj_path, attr = prop.split('.', 1)
                    obj = ps
                    for part in obj_path.split('.'):
                        obj = getattr(obj, part)
                    val = getattr(obj, attr)
                print(f"  {prop}: {val}")
                page_setup_before[prop] = str(val)
            except Exception as e:
                print(f"  {prop}: ERROR: {e}")
                page_setup_before[prop] = f"ERROR: {e}"
        
        # Dump printer info
        try:
            printer_info = {
                'name': excel.ActivePrinter,
            }
            print(f"\n  ActivePrinter: {excel.ActivePrinter}")
        except Exception as e:
            printer_info = {'error': str(e)}
        
        # Dump worksheet properties
        print(f"\n  Worksheet.Name: {ws.Name}")
        print(f"  Worksheet.Index: {ws.Index}")
        try:
            pa = ws.PageSetup.PrintArea
            print(f"  PrintArea: {pa}")
        except:
            print(f"  PrintArea: (none/error)")
        
        try:
            ur = ws.UsedRange
            print(f"  UsedRange: {ur.Address}  Rows={ur.Rows.Count} Cols={ur.Columns.Count}")
            print(f"  UsedRange.Width: {ur.Width}  Height: {ur.Height}")
        except Exception as e:
            print(f"  UsedRange: ERROR: {e}")
        
        # Check display page breaks
        try:
            print(f"  DisplayPageBreaks: {ws.DisplayPageBreaks}")
        except:
            pass
        
        # Check printer
        print(f"  Default Printer: {excel.ActivePrinter}")
        
        # Export to PDF
        print("\nExporting to PDF...")
        pdf_format = 0  # xlTypePDF
        wb.ExportAsFixedFormat(pdf_format, reexport_path)
        print(f"  Exported to: {reexport_path}")
        
        # Dump PageSetup again AFTER export
        print("\nPageSetup AFTER export:")
        for prop in page_setup_props:
            try:
                val = getattr(ps, prop.split('.')[0]) if '.' not in prop else None
                if '.' in prop:
                    obj_path, attr = prop.split('.', 1)
                    obj = ps
                    for part in obj_path.split('.'):
                        obj = getattr(obj, part)
                    val = getattr(obj, attr)
                print(f"  {prop}: {val}")
                page_setup_after[prop] = str(val)
            except Exception as e:
                page_setup_after[prop] = f"ERROR: {e}"
        
        wb.Close(SaveChanges=False)
        excel.Quit()
        
    except Exception as e:
        print(f"ERROR during COM operation: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except: pass
        try:
            if excel: excel.Quit()
        except: pass
        pythoncom.CoUninitialize()
    
    # Save page setup dump
    ps_dump = {
        "before_export": page_setup_before,
        "after_export": page_setup_after,
        "printer_info": printer_info if 'printer_info' in dir() else {},
    }
    dump_path = os.path.join(OUT, "page_setup_dump.json")
    with open(dump_path, 'w') as f:
        json.dump(ps_dump, f, indent=2, default=str)
    
    print("\nInvestigation C complete.")
    return reexport_path

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION D — PDF Forensics
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_d():
    """Compare background.pdf vs reexport.pdf at binary, structural, and visual levels."""
    print("\n" + "=" * 70)
    print("INVESTIGATION D — PDF Forensics")
    print("=" * 70)
    
    bg_path = os.path.join(OUT, "background.pdf")
    re_path = os.path.join(OUT, "reexport.pdf")
    
    if not os.path.exists(bg_path):
        print("ERROR: background.pdf not found.")
        return
    if not os.path.exists(re_path):
        print("ERROR: reexport.pdf not found. Run Investigation C first.")
        return
    
    # Binary comparison
    with open(bg_path, 'rb') as f: bg_data = f.read()
    with open(re_path, 'rb') as f: re_data = f.read()
    
    binary_identical = bg_data == re_data
    bg_sha = sha256_of(bg_data)
    re_sha = sha256_of(re_data)
    bg_md5 = md5_of(bg_data)
    re_md5 = md5_of(re_data)
    
    print(f"\nBinary comparison:")
    print(f"  background.pdf: {len(bg_data)} bytes  SHA256={bg_sha[:16]}...  MD5={bg_md5}")
    print(f"  reexport.pdf:   {len(re_data)} bytes  SHA256={re_sha[:16]}...  MD5={re_md5}")
    print(f"  Binary identical: {binary_identical}")
        
    # Structural comparison via PyMuPDF
    bg_doc = fitz.open(bg_path)
    re_doc = fitz.open(re_path)
    
    bg_page = bg_doc[0]
    re_page = re_doc[0]
    
    print(f"\nPDF Structure:")
    print(f"  background.pdf pages: {bg_doc.page_count}")
    print(f"  reexport.pdf pages:   {re_doc.page_count}")
    print(f"  background.pdf MediaBox: {bg_page.rect}")
    print(f"  reexport.pdf MediaBox:   {re_page.rect}")
    
    # Page dimensions
    bg_w = round(bg_page.rect.width, 2)
    bg_h = round(bg_page.rect.height, 2)
    re_w = round(re_page.rect.width, 2)
    re_h = round(re_page.rect.height, 2)
    print(f"  background.pdf dims: {bg_w} x {bg_h}")
    print(f"  reexport.pdf dims:   {re_w} x {re_h}")
    
    # Content bounds from text blocks
    for label, page in [("background.pdf", bg_page), ("reexport.pdf", re_page)]:
        blocks = page.get_text('blocks')
        drawings = page.get_drawings()
        images = page.get_images(full=True)
        
        all_bounds = []
        for b in blocks:
            all_bounds.append((b[0], b[1], b[2], b[3]))
        for d in drawings:
            rr = d.get('rect')
            if rr:
                all_bounds.append((rr.x0, rr.y0, rr.x1, rr.y1))
        for img in images:
            try:
                ir = page.get_image_bbox(img)
                if ir:
                    all_bounds.append((ir.x0, ir.y0, ir.x1, ir.y1))
            except:
                pass
        
        if all_bounds:
            x0 = min(b[0] for b in all_bounds)
            y0 = min(b[1] for b in all_bounds)
            x1 = max(b[2] for b in all_bounds)
            y1 = max(b[3] for b in all_bounds)
            print(f"\n  {label} content bounds:")
            print(f"    left={x0:.2f} top={y0:.2f} right={x1:.2f} bottom={y1:.2f}")
            print(f"    width={x1-x0:.2f} height={y1-y0:.2f}")
        else:
            print(f"\n  {label}: no content found")
        
        print(f"  {label} text blocks: {len(blocks)}")
        print(f"  {label} drawings: {len(drawings)}")
        print(f"  {label} images: {len(images)}")
    
    # Check for Tj/TJ operators in content stream (text presence)
    for label, doc in [("background.pdf", bg_doc), ("reexport.pdf", re_doc)]:
        for i in range(doc.page_count):
            p = doc[i]
            xref = p.get_contents()
            if xref:
                try:
                    stream = doc.xref_stream(xref[0]) if isinstance(xref, list) else doc.xref_stream(xref)
                    if stream:
                        stream_str = stream.decode('latin-1')
                        tj_count = stream_str.count('Tj') + stream_str.count('TJ')
                        cm_count = stream_str.count('cm')
                        re_count = stream_str.count(' re ')
                        print(f"\n  {label} page {i} content stream stats:")
                        print(f"    Tj/TJ (text): {tj_count}")
                        print(f"    cm (transform): {cm_count}")
                        print(f"    re (rectangle): {re_count}")
                except:
                    pass
    
    bg_doc.close()
    re_doc.close()
    
    pdf_compare = {
        "binary_identical": binary_identical,
        "background": {
            "sha256": bg_sha, "md5": bg_md5, "size": len(bg_data),
            "pages": 1,
        },
        "reexport": {
            "sha256": re_sha, "md5": re_md5, "size": len(re_data),
            "pages": 1,
        },
        "bg_dims_pt": {"w": bg_w, "h": bg_h},
        "re_dims_pt": {"w": re_w, "h": re_h},
    }
    
    # Restore from re-read
    bg_doc2 = fitz.open(bg_path)
    re_doc2 = fitz.open(re_path)
    pdf_compare["background"]["pages"] = bg_doc2.page_count
    pdf_compare["reexport"]["pages"] = re_doc2.page_count
    bg_doc2.close()
    re_doc2.close()
    
    compare_path = os.path.join(OUT, "pdf_compare.json")
    with open(compare_path, 'w') as f:
        json.dump(pdf_compare, f, indent=2)
    
    print("\nInvestigation D complete.")
    return pdf_compare

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION E — Overlay Validation
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_e(metadata):
    """Convert database ratios to page coordinates and overlay on both PDFs."""
    print("\n" + "=" * 70)
    print("INVESTIGATION E — Database Coordinate Overlay")
    print("=" * 70)
    
    # Get page dimensions from xml_data
    xml_path = os.path.join(OUT, "xml_data.xml")
    with open(xml_path, 'r') as f:
        xml_str = f.read()
    
    pw = float(re.search(r'<width>(\d+)</width>', xml_str).group(1)) if re.search(r'<width>(\d+)</width>', xml_str) else 612
    ph = float(re.search(r'<height>(\d+)</height>', xml_str).group(1)) if re.search(r'<height>(\d+)</height>', xml_str) else 792
    
    # Also get margins from XML
    ml = float(re.search(r'<marginLeft>([\d.]+)</marginLeft>', xml_str).group(1)) if re.search(r'<marginLeft>([\d.]+)</marginLeft>', xml_str) else 0
    mr = float(re.search(r'<marginRight>([\d.]+)</marginRight>', xml_str).group(1)) if re.search(r'<marginRight>([\d.]+)</marginRight>', xml_str) else 0
    mt = float(re.search(r'<marginTop>([\d.]+)</marginTop>', xml_str).group(1)) if re.search(r'<marginTop>([\d.]+)</marginTop>', xml_str) else 0
    mb = float(re.search(r'<marginBottom>([\d.]+)</marginBottom>', xml_str).group(1)) if re.search(r'<marginBottom>[\d.]+</marginBottom>', xml_str) else 0
    
    ch = re.search(r'<centerH>([^<]+)</centerH>', xml_str)
    center_h = ch and ch.group(1).strip() in ("1", "true")
    cv = re.search(r'<centerV>([^<]+)</centerV>', xml_str)
    center_v = cv and cv.group(1).strip() in ("1", "true")
    
    print(f"\nPage dimensions from XML: {pw} x {ph}")
    print(f"Margins: L={ml} R={mr} T={mt} B={mb}")
    print(f"Center: H={center_h} V={center_v}")
    
    # Convert database ratios to page coordinates
    clusters = metadata['clusters']
    geometry = []
    
    for c in clusters:
        left_pt = float(c['left_position']) * pw
        top_pt = float(c['top_position']) * ph
        right_pt = float(c['right_position']) * pw
        bottom_pt = float(c['bottom_position']) * ph
        w = right_pt - left_pt
        h = bottom_pt - top_pt
        
        geo = {
            "cluster_id": c['cluster_id'],
            "cell_addr": c['cell_addr'],
            "ratio": {"L": c['left_position'], "R": c['right_position'], "T": c['top_position'], "B": c['bottom_position']},
            "page_pt": {"left": round(left_pt, 4), "top": round(top_pt, 4), "right": round(right_pt, 4), "bottom": round(bottom_pt, 4)},
            "size_pt": {"width": round(w, 4), "height": round(h, 4)},
        }
        geometry.append(geo)
        print(f"\n  cluster={c['cluster_id']} {c['cell_addr']}:")
        print(f"    ratios: L={c['left_position']} R={c['right_position']} T={c['top_position']} B={c['bottom_position']}")
        print(f"    page(pt): left={left_pt:.2f} top={top_pt:.2f} right={right_pt:.2f} bottom={bottom_pt:.2f}")
        print(f"    size(pt): W={w:.2f} H={h:.2f}")
    
    # Save geometry dump
    geometry_dump = {
        "page_dimensions_pt": {"width": pw, "height": ph},
        "margins_pt": {"left": ml, "right": mr, "top": mt, "bottom": mb},
        "center": {"horizontal": center_h, "vertical": center_v},
        "clusters": geometry,
    }
    geo_path = os.path.join(OUT, "geometry_dump.json")
    with open(geo_path, 'w') as f:
        json.dump(geometry_dump, f, indent=2)
    
    # Overlay on background.pdf
    _draw_overlay_pdf(
        os.path.join(OUT, "background.pdf"),
        os.path.join(OUT, "db_overlay_background.png"),
        geometry, pw, ph, "Database coordinates on background.pdf"
    )
    
    # Overlay on reexport.pdf
    re_path = os.path.join(OUT, "reexport.pdf")
    if os.path.exists(re_path):
        _draw_overlay_pdf(
            re_path,
            os.path.join(OUT, "db_overlay_reexport.png"),
            geometry, pw, ph, "Database coordinates on reexport.pdf"
        )
    
    print("\nInvestigation E complete.")
    return geometry

def _draw_overlay_pdf(pdf_path, output_png, geometry, pw, ph, title):
    """Render a PDF page and draw database coordinate rectangles on it."""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        
        # Render at 150 DPI for overlay
        zoom = 150.0 / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Create a new pixmap with overlay
        rv = 150.0 / 72.0  # render scale
        
        # We'll use fitz annotations instead
        for g in geometry:
            left = float(g['page_pt']['left']) * rv
            top = float(g['page_pt']['top']) * rv
            w = float(g['size_pt']['width']) * rv
            h = float(g['size_pt']['height']) * rv
            
            # Stay within page bounds
            pw_px = pw * rv
            ph_px = ph * rv
            
            # Add rectangle annotation
            rect = fitz.Rect(left, top, left + w, top + h)
            annot = page.add_rect_annot(rect)
            annot.set_border(width=2)
            annot.set_colors(stroke=[1, 0, 0])  # Red
            annot.update()
        
        # Add title text
        # Save with annotations
        pix2 = page.get_pixmap(matrix=mat)
        pix2.save(output_png)
        doc.close()
        print(f"  Saved overlay: {output_png}")
    except Exception as e:
        print(f"  ERROR drawing overlay on {pdf_path}: {e}")
        import traceback
        traceback.print_exc()

# ═══════════════════════════════════════════════════════════════════════════════
# Render comparison at multiple DPI (Investigation E extension)
# ═══════════════════════════════════════════════════════════════════════════════
def render_comparison():
    """Render both PDFs at multiple DPI and create comparison overlays."""
    print("\n" + "=" * 70)
    print("INVESTIGATION E (ext) — Render Comparison at Multiple DPI")
    print("=" * 70)
    
    bg_path = os.path.join(OUT, "background.pdf")
    re_path = os.path.join(OUT, "reexport.pdf")
    
    if not os.path.exists(re_path):
        print("WARNING: reexport.pdf not found. Skipping render comparison.")
        return
    
    dpis = [72, 150, 300, 600]
    
    for dpi in dpis:
        zoom = dpi / 72.0
        
        try:
            # Render background.pdf
            bg_doc = fitz.open(bg_path)
            bg_page = bg_doc[0]
            mat = fitz.Matrix(zoom, zoom)
            bg_pix = bg_page.get_pixmap(matrix=mat)
            
            # Render reexport.pdf
            re_doc = fitz.open(re_path)
            re_page = re_doc[0]
            re_pix = re_page.get_pixmap(matrix=mat)
            
            # Compare dimensions
            print(f"\n  {dpi} DPI:")
            print(f"    background.pdf: {bg_pix.width}x{bg_pix.height}px")
            print(f"    reexport.pdf:   {re_pix.width}x{re_pix.height}px")
            
            # Create difference image if same size
            if bg_pix.width == re_pix.width and bg_pix.height == re_pix.height:
                bg_samples = bg_pix.samples
                re_samples = re_pix.samples
                diff = bytearray(len(bg_samples))
                
                diff_pixels = 0
                max_diff = 0
                for i in range(0, len(bg_samples), 3):
                    dr = abs(bg_samples[i] - re_samples[i])
                    dg = abs(bg_samples[i+1] - re_samples[i+1])
                    db = abs(bg_samples[i+2] - re_samples[i+2])
                    d = max(dr, dg, db)
                    if d > 10:  # threshold
                        diff_pixels += 1
                        diff[i:i+3] = (255, 0, 0)  # Red highlight
                    else:
                        diff[i:i+3] = bg_samples[i:i+3]
                    if d > max_diff: max_diff = d
                
                total_pixels = len(bg_samples) // 3
                pct = 100.0 * diff_pixels / total_pixels if total_pixels > 0 else 0
                print(f"    pixel differences (threshold>10): {diff_pixels}/{total_pixels} = {pct:.4f}%")
                print(f"    max per-channel diff: {max_diff}")
                
                # Save diff image
                from fitz import Pixmap
                diff_pix = Pixmap(fitz.csRGB, bg_pix.width, bg_pix.height, bytes(diff), False)
                diff_path = os.path.join(OUT, f"compare_overlay_{dpi}dpi.png") if dpi > 72 else os.path.join(OUT, "compare_overlay_72.png")
                if dpi == 150:
                    diff_path = os.path.join(OUT, "compare_overlay_150.png")
                elif dpi == 300:
                    diff_path = os.path.join(OUT, "compare_overlay_300.png")
                elif dpi == 600:
                    diff_path = os.path.join(OUT, "compare_overlay_600.png")
                diff_pix.save(diff_path)
                print(f"    saved: {diff_path}")
            else:
                print(f"    SIZE MISMATCH: cannot compare directly")
            
            bg_doc.close()
            re_doc.close()
            
        except Exception as e:
            print(f"    ERROR at {dpi} DPI: {e}")
            import traceback
            traceback.print_exc()

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION F — COM Pipeline Validation
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_f(metadata):
    """Simulate COM coordinate calculation and compare with database coordinates."""
    print("\n" + "=" * 70)
    print("INVESTIGATION F — COM Pipeline Coordinate Validation")
    print("=" * 70)
    
    # Get XML data for page setup
    xml_path = os.path.join(OUT, "xml_data.xml")
    with open(xml_path, 'r') as f:
        xml_str = f.read()
    
    pw = float(re.search(r'<width>(\d+)</width>', xml_str).group(1))
    ph = float(re.search(r'<height>(\d+)</height>', xml_str).group(1))
    
    ml = float(re.search(r'<marginLeft>([\d.]+)</marginLeft>', xml_str).group(1)) if re.search(r'<marginLeft>([\d.]+)</marginLeft>', xml_str) else 51.024
    mr = float(re.search(r'<marginRight>([\d.]+)</marginRight>', xml_str).group(1)) if re.search(r'<marginRight>([\d.]+)</marginRight>', xml_str) else 51.024
    mt = float(re.search(r'<marginTop>([\d.]+)</marginTop>', xml_str).group(1)) if re.search(r'<marginTop>([\d.]+)</marginTop>', xml_str) else 53.858
    mb = float(re.search(r'<marginBottom>([\d.]+)</marginBottom>', xml_str).group(1)) if re.search(r'<marginBottom>[\d.]+</marginBottom>', xml_str) else 53.858
    
    printable_w = pw - ml - mr
    printable_h = ph - mt - mb
    
    # Read worksheet geometry from the XLSX
    xlsx_path = os.path.join(OUT, "original.xlsx")
    zf = zipfile.ZipFile(xlsx_path, 'r')
    ws_raw = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
    zf.close()
    
    # Parse column widths
    col_defs = {}
    cols_match = re.search(r'<cols>(.*?)</cols>', ws_raw, re.DOTALL)
    if cols_match:
        for m in re.finditer(r'<col\s+([^>]+)/?>', cols_match.group(1)):
            a = m.group(1)
            min_c = int(re.search(r'min="(\d+)"', a).group(1))
            max_c = int(re.search(r'max="(\d+)"', a).group(1))
            w = float(re.search(r'width="([\d.]+)"', a).group(1))
            hidden = 'hidden="1"' in a
            for c in range(min_c, max_c + 1):
                if not hidden:
                    col_defs[c] = w
    
    # Parse row heights
    row_defs = {}
    for m in re.finditer(r'<row\s+([^>]+?)/?>', ws_raw):
        a = m.group(1)
        rn = int(re.search(r'r="(\d+)"', a).group(1))
        ht_match = re.search(r'ht="([\d.]+)"', a)
        if ht_match:
            row_defs[rn] = float(ht_match.group(1))
        hidden = 'hidden="1"' in a
    
    def char_to_points(cw):
        """Convert Excel character width to points."""
        return (cw * 7.33 + 5.0) * 72.0 / 96.0
    
    def col_letter_to_num(s):
        n = 0
        for c in s.upper().strip(): n = n * 26 + (ord(c) - 64)
        return n
    
    def parse_range(rng):
        """Parse a range like $A$1:$B$2 into (col_start, row_start, col_end, row_end)."""
        m = re.match(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', rng)
        if m:
            return col_letter_to_num(m.group(1)), int(m.group(2)), col_letter_to_num(m.group(3)), int(m.group(4))
        # Single cell
        m = re.match(r'\$?([A-Z]+)\$?(\d+)', rng)
        if m:
            cn = col_letter_to_num(m.group(1))
            rn = int(m.group(2))
            return cn, rn, cn, rn
        return None
    
    # Get print area from workbook XML
    zf = zipfile.ZipFile(xlsx_path, 'r')
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    zf.close()
    
    # Try to find Print_Area defined name
    print_area_addr = None
    for m in re.finditer(r'<definedName[^>]*>([^<]+)</definedName>', wb_xml):
        content = m.group(1)
        if '_Print_Area' in content or 'Print_Area' in m.group(0):
            parts = content.split('!')
            if len(parts) > 1:
                print_area_addr = parts[-1].strip()
                break
    
    if not print_area_addr:
        # Also check if there's a printArea directly in the worksheet
        pa_match = re.search(r'<printArea[^>]*>([^<]+)</printArea>', ws_raw)
        if pa_match:
            print_area_addr = pa_match.group(1)
    
    print(f"\nPrintArea from XLSX: {print_area_addr}")
    
    # Calculate print area dimensions
    pa_col_start = 1
    pa_row_start = 1
    pa_col_end = 0
    pa_row_end = 0
    
    if print_area_addr:
        r = parse_range(print_area_addr)
        if r:
            pa_col_start, pa_row_start, pa_col_end, pa_row_end = r
    
    # If no print area, use max column/row from cluster data
    clusters = metadata['clusters']
    if pa_col_end == 0:
        for c in clusters:
            r = parse_range(c['cell_addr'])
            if r:
                pa_col_end = max(pa_col_end, r[2])
                pa_row_end = max(pa_row_end, r[3])
        print(f"  (PrintArea not set, using cluster extents: col={pa_col_end}, row={pa_row_end})")
    
    # Calculate PA dimensions
    pa_width = 0
    for c in range(pa_col_start, pa_col_end + 1):
        w = col_defs.get(c, 8.43)
        pa_width += char_to_points(w)
    
    pa_height = 0
    for r in range(pa_row_start, pa_row_end + 1):
        h = row_defs.get(r, 15.0)
        pa_height += h  # row heights are already in points in Excel
    
    print(f"  PrintArea computed: W={pa_width:.2f}pt H={pa_height:.2f}pt")
    print(f"  Printable area: W={printable_w:.2f}pt H={printable_h:.2f}pt")
    
    # For each cluster, compute the COM coordinates (Range.Left, Range.Top)
    # and compare to database coordinates
    comparisons = []
    for c in clusters:
        r = parse_range(c['cell_addr'])
        if not r:
            continue
        cs, rs, ce, re_val = r
        
        # Compute Range.Left (position of left edge of first column)
        range_left = 0
        for col in range(1, cs):
            w = col_defs.get(col, 8.43)
            range_left += char_to_points(w)
        
        # Compute Range.Top (position of top edge of first row)
        range_top = 0
        for row in range(1, rs):
            h = row_defs.get(row, 15.0)
            range_top += h
        
        # Compute Range.Width (width of all columns in range)
        range_width = 0
        for col in range(cs, ce + 1):
            w = col_defs.get(col, 8.43)
            range_width += char_to_points(w)
        
        # Compute Range.Height (height of all rows in range)
        range_height = 0
        for row in range(rs, re_val + 1):
            h = row_defs.get(row, 15.0)
            range_height += h
        
        # Database page coordinates (from ratios)
        db_left = float(c['left_position']) * pw
        db_top = float(c['top_position']) * ph
        db_right = float(c['right_position']) * pw
        db_bottom = float(c['bottom_position']) * ph
        db_w = db_right - db_left
        db_h = db_bottom - db_top
        
        # COM simple formula (margins + range position)
        com_left_simple = ml + range_left
        com_top_simple = mt + range_top
        com_w_simple = range_width
        com_h_simple = range_height
        
        legacy_left = db_left
        legacy_top = db_top
        
        # Compute effective dimensions from the database coords of first field
        first_field = clusters[0]
        ff_left = float(first_field['left_position']) * pw
        ff_top = float(first_field['top_position']) * ph
        
        eff_w = printable_w - 2 * (ff_left - ml)
        eff_h = printable_h - 2 * (ff_top - mt)
        
        scale_w = eff_w / pa_width if pa_width > 0 else 1
        scale_h = eff_h / pa_height if pa_height > 0 else 1
        
        # ConMas formula: page = margin + (printable - eff)/2 + range_pos * scale
        conmas_left = ml + (printable_w - eff_w) / 2 + range_left * scale_w
        conmas_top = mt + (printable_h - eff_h) / 2 + range_top * scale_h
        conmas_w = range_width * scale_w
        conmas_h = range_height * scale_h
        
        comp = {
            "cluster_id": c['cluster_id'],
            "cell_addr": c['cell_addr'],
            "worksheet_pt": {
                "Range.Left": round(range_left, 4),
                "Range.Top": round(range_top, 4),
                "Range.Width": round(range_width, 4),
                "Range.Height": round(range_height, 4),
            },
            "database_pt": {
                "left": round(db_left, 4),
                "top": round(db_top, 4),
                "width": round(db_w, 4),
                "height": round(db_h, 4),
            },
            "com_simple_pt": {
                "left": round(com_left_simple, 4),
                "top": round(com_top_simple, 4),
                "width": round(com_w_simple, 4),
                "height": round(com_h_simple, 4),
            },
            "conmas_formula_pt": {
                "left": round(conmas_left, 4),
                "top": round(conmas_top, 4),
                "width": round(conmas_w, 4),
                "height": round(conmas_h, 4),
            },
            "error_conmas_vs_db": {
                "left": round(conmas_left - db_left, 4),
                "top": round(conmas_top - db_top, 4),
                "width": round(conmas_w - db_w, 4),
                "height": round(conmas_h - db_h, 4),
            },
            "scaling": {
                "effW": round(eff_w, 4),
                "effH": round(eff_h, 4),
                "scaleW": round(scale_w, 6),
                "scaleH": round(scale_h, 6),
            }
        }
        comparisons.append(comp)
        
        print(f"\n  cluster={c['cluster_id']} {c['cell_addr']}:")
        print(f"    Range: L={range_left:.2f} T={range_top:.2f} W={range_width:.2f} H={range_height:.2f}")
        print(f"    DB:   L={db_left:.2f} T={db_top:.2f} W={db_w:.2f} H={db_h:.2f}")
        print(f"    COM:  L={com_left_simple:.2f} T={com_top_simple:.2f} W={com_w_simple:.2f} H={com_h_simple:.2f}")
        print(f"    ConMas: L={conmas_left:.2f} T={conmas_top:.2f} W={conmas_w:.2f} H={conmas_h:.2f}")
        print(f"    DeltaConMas-DB: L={conmas_left-db_left:.4f} T={conmas_top-db_top:.4f} W={conmas_w-db_w:.4f} H={conmas_h-db_h:.4f}")
        print(f"    Scaling: effW={eff_w:.2f} effH={eff_h:.2f} scaleW={scale_w:.6f} scaleH={scale_h:.6f}")
    
    # Save COM pipeline dump
    com_dump = {
        "page_dimensions_pt": {"width": pw, "height": ph, "printable_w": printable_w, "printable_h": printable_h},
        "margins_pt": {"left": ml, "right": mr, "top": mt, "bottom": mb},
        "print_area_addr": print_area_addr,
        "print_area_pt": {"width": round(pa_width, 4), "height": round(pa_height, 4)},
        "effective_dimensions_pt": {"width": round(eff_w, 4), "height": round(eff_h, 4)},
        "comparisons": comparisons,
    }
    dump_path = os.path.join(OUT, "com_pipeline_dump.json")
    with open(dump_path, 'w') as f:
        json.dump(com_dump, f, indent=2, default=str)
    
    # Draw COM coordinates on reexport.pdf
    re_path = os.path.join(OUT, "reexport.pdf")
    if os.path.exists(re_path):
        _draw_overlay_com(re_path, os.path.join(OUT, "com_overlay_reexport.png"),
                          comparisons, pw, ph, "COM coordinates on reexport.pdf")
    
    print("\nInvestigation F complete.")
    return comparisons

def _draw_overlay_com(pdf_path, output_png, comparisons, pw, ph, title):
    """Overlay COM coordinates on a PDF."""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        rv = 150.0 / 72.0
        
        for c in comparisons:
            con = c['conmas_formula_pt']
            left = float(con['left']) * rv
            top = float(con['top']) * rv
            w = float(con['width']) * rv
            h = float(con['height']) * rv
            
            rect = fitz.Rect(left, top, left + w, top + h)
            annot = page.add_rect_annot(rect)
            annot.set_border(width=2)
            annot.set_colors(stroke=[0, 0, 1])  # Blue for COM
            annot.update()
        
        # Also draw database coords in red for comparison
        for c in comparisons:
            db = c['database_pt']
            left = float(db['left']) * rv
            top = float(db['top']) * rv
            w = float(db['width']) * rv
            h = float(db['height']) * rv
            
            rect = fitz.Rect(left, top, left + w, top + h)
            annot = page.add_rect_annot(rect)
            annot.set_border(width=1)
            annot.set_colors(stroke=[1, 0, 0])  # Red for DB
            annot.update()
        
        mat = fitz.Matrix(rv, rv)
        pix = page.get_pixmap(matrix=mat)
        pix.save(output_png)
        doc.close()
        print(f"  Saved: {output_png}")
    except Exception as e:
        print(f"  ERROR: {e}")

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION G — XML Correlation
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_g(metadata):
    """Compare xml_data.xml ratios vs def_cluster database ratios."""
    print("\n" + "=" * 70)
    print("INVESTIGATION G — XML vs Database Coordinate Correlation")
    print("=" * 70)
    
    xml_path = os.path.join(OUT, "xml_data.xml")
    with open(xml_path, 'r') as f:
        xml_str = f.read()
    
    # Parse XML clusters
    xml_clusters = []
    for m in re.finditer(r'<cluster>(.*?)</cluster>', xml_str, re.DOTALL):
        cx = m.group(1)
        cluster = {}
        for tag in ['clusterId', 'cluster_id', 'cellAddress', 'cell_address',
                     'left', 'right', 'top', 'bottom',
                     'name', 'type']:
            tm = re.search(r'<' + tag + r'>(.*?)</' + tag + r'>', cx)
            if tm:
                cluster[tag] = tm.group(1)
        if cluster:
            xml_clusters.append(cluster)
    
    print(f"\nXML clusters found: {len(xml_clusters)}")
    
    # Normalize XML cluster IDs
    for xc in xml_clusters:
        cid = xc.get('clusterId', xc.get('cluster_id', ''))
        addr = xc.get('cellAddress', xc.get('cell_address', ''))
        ratio_left = xc.get('left', '')
        ratio_right = xc.get('right', '')
        ratio_top = xc.get('top', '')
        ratio_bottom = xc.get('bottom', '')
        print(f"\n  XML cluster id={cid} addr={addr}")
        print(f"    ratios: L={ratio_left} R={ratio_right} T={ratio_top} B={ratio_bottom}")
    
    # Compare with database clusters
    db_clusters = metadata['clusters']
    print(f"\n{'='*50}")
    print("XML vs DATABASE ratio comparison:")
    print(f"{'='*50}")
    
    field_matches = 0
    field_mismatches = 0
    field_details = []
    
    for xc in xml_clusters:
        x_addr = xc.get('cellAddress', xc.get('cell_address', '')).upper().replace('$', '')
        x_left = xc.get('left', '').strip()
        x_right = xc.get('right', '').strip()
        x_top = xc.get('top', '').strip()
        x_bottom = xc.get('bottom', '').strip()
        
        # Find matching DB cluster
        db_match = None
        for dc in db_clusters:
            db_addr = dc['cell_addr'].upper().replace('$', '')
            if db_addr == x_addr:
                db_match = dc
                break
        
        if db_match:
            d_left = db_match['left_position'].strip()
            d_right = db_match['right_position'].strip()
            d_top = db_match['top_position'].strip()
            d_bottom = db_match['bottom_position'].strip()
            
            match_l = abs(float(x_left) - float(d_left)) < 0.00001 if x_left and d_left else False
            match_r = abs(float(x_right) - float(d_right)) < 0.00001 if x_right and d_right else False
            match_t = abs(float(x_top) - float(d_top)) < 0.00001 if x_top and d_top else False
            match_b = abs(float(x_bottom) - float(d_bottom)) < 0.00001 if x_bottom and d_bottom else False
            
            all_match = all([match_l, match_r, match_t, match_b])
            
            detail = {
                "cell_addr": x_addr,
                "xml": {"L": x_left, "R": x_right, "T": x_top, "B": x_bottom},
                "db": {"L": d_left, "R": d_right, "T": d_top, "B": d_bottom},
                "match": {"L": match_l, "R": match_r, "T": match_t, "B": match_b},
                "all_match": all_match,
            }
            field_details.append(detail)
            
            if all_match:
                field_matches += 1
                print(f"  {x_addr}: MATCH OK")
            else:
                field_mismatches += 1
                print(f"  {x_addr}: MISMATCH ✗")
                print(f"    XML: L={x_left} R={x_right} T={x_top} B={x_bottom}")
                print(f"    DB:  L={d_left} R={d_right} T={d_top} B={d_bottom}")
        else:
            print(f"  {x_addr}: NO DB MATCH (address not found in database)")
    
    print(f"\n  Total: {field_matches} matches, {field_mismatches} mismatches")
    
    xml_correlation = {
        "xml_cluster_count": len(xml_clusters),
        "db_cluster_count": len(db_clusters),
        "field_matches": field_matches,
        "field_mismatches": field_mismatches,
        "details": field_details,
    }
    
    # Also check for additional XML fields (left_pt, top_pt, width_pt, height_pt)
    for suffix in ['_pt', 'Pt']:
        for tag in ['left', 'top', 'width', 'height']:
            tm = re.search(r'<' + tag + suffix + r'>(.*?)</' + tag + suffix + r'>', xml_str)
            if tm:
                print(f"\n  XML also has <{tag}{suffix}>: {tm.group(1)}")
                xml_correlation[f"xml_{tag}{suffix}"] = tm.group(1)
    
    corr_path = os.path.join(OUT, "xml_correlation.json")
    with open(corr_path, 'w') as f:
        json.dump(xml_correlation, f, indent=2, default=str)
    
    print("\nInvestigation G complete.")
    return xml_correlation

# ═══════════════════════════════════════════════════════════════════════════════
# INVESTIGATION H — Generation Order Reconstruction
# ═══════════════════════════════════════════════════════════════════════════════
def investigation_h(metadata, pdf_compare, comparisons, xml_correlation):
    """Analyze all evidence to determine the original generation pipeline."""
    print("\n" + "=" * 70)
    print("INVESTIGATION H — Generation Order Reconstruction")
    print("=" * 70)
    
    results = {}
    
    # H1: Is def_file the exact workbook used to generate the stored PDF?
    print("\nH1: Does def_file produce background.pdf?")
    
    if pdf_compare:
        if pdf_compare.get('binary_identical'):
            print("  YES — reexport.pdf is binary identical to background.pdf")
            results['h1_answer'] = "YES — binary identical"
            results['h1_evidence'] = "reexport.pdf (fresh ExportAsFixedFormat) matches background.pdf byte-for-byte"
        else:
            print("  NO — binary difference detected")
            bg_sz = pdf_compare['background']['size']
            re_sz = pdf_compare['reexport']['size']
            print(f"  background.pdf: {bg_sz} bytes")
            print(f"  reexport.pdf:   {re_sz} bytes")
            results['h1_answer'] = f"NO — sizes differ: bg={bg_sz}, re={re_sz}"
        
        # Check structural identity
        results['h1_background_pages'] = pdf_compare.get('background', {}).get('pages', '?')
        results['h1_reexport_pages'] = pdf_compare.get('reexport', {}).get('pages', '?')
    
    # H2: Were coordinates generated from worksheet, page, PDF, or other?
    print("\nH2: Source of def_cluster coordinates")
    
    if comparisons:
        errors = comparisons[0]['error_conmas_vs_db']
        max_err = max(abs(errors['left']), abs(errors['top']), abs(errors['width']), abs(errors['height']))
        
        if max_err < 0.5:
            print(f"  ConMas formula error: Delta={max_err:.4f}pt")
            print("  Coordinates CAN be derived from worksheet geometry + page setup")
            results['h2_answer'] = "worksheet geometry + page setup (ConMas formula)"
            results['h2_evidence'] = f"ConMas formula matches database within {max_err:.4f}pt"
            results['h2_max_error_pt'] = max_err
        else:
            print(f"  ConMas formula error: Delta={max_err:.4f}pt")
            print("  Coordinates CANNOT be derived from worksheet geometry alone")
            results['h2_answer'] = "unknown (worksheet geometry insufficient)"
            results['h2_evidence'] = f"ConMas formula error: {max_err:.4f}pt"
    
    # H3: Are XML and database generated from same source?
    print("\nH3: XML vs Database coordinate source")
    if xml_correlation:
        xml_correlation.get('field_matches', 0)
        xml_correlation.get('field_mismatches', 0)
        total = xml_correlation.get('field_matches', 0) + xml_correlation.get('field_mismatches', 0)
        match_pct = 100.0 * xml_correlation.get('field_matches', 0) / total if total > 0 else 0
        
        if xml_correlation.get('field_matches', 0) == total and total > 0:
            print(f"  IDENTICAL — {total}/{total} fields match ({match_pct:.0f}%)")
            results['h3_answer'] = "Same source (identical ratios)"
        elif match_pct > 50:
            print(f"  SIMILAR — {xml_correlation['field_matches']}/{total} match ({match_pct:.1f}%)")
            results['h3_answer'] = f"Likely same source ({match_pct:.1f}% match)"
        else:
            print(f"  DIFFERENT — only {xml_correlation['field_matches']}/{total} match")
            results['h3_answer'] = "Independently generated"
        results['h3_match_pct'] = round(match_pct, 1)
    
    # H4: Can the process be reproduced deterministically?
    print("\nH4: Deterministic reproduction")
    
    if comparisons:
        all_errors = []
        for c in comparisons:
            e = c['error_conmas_vs_db']
            all_errors.append(e['left'])
            all_errors.append(e['top'])
            all_errors.append(e['width'])
            all_errors.append(e['height'])
        
        max_abs_err = max(abs(x) for x in all_errors)
        avg_abs_err = sum(abs(x) for x in all_errors) / len(all_errors)
        
        print(f"  Max error: {max_abs_err:.4f}pt")
        print(f"  Avg error: {avg_abs_err:.4f}pt")
        
        threshold = 1.0  # 1 point threshold
        if max_abs_err < threshold:
            print(f"  YES — within {threshold}pt threshold")
            results['h4_answer'] = "YES"
            results['h4_max_error_pt'] = round(max_abs_err, 4)
            results['h4_avg_error_pt'] = round(avg_abs_err, 4)
            results['h4_pipeline'] = "Original def_file → COM Open → ExportAsFixedFormat → PDF coordinates measured from rendered output"
        else:
            print(f"  NO — exceeds {threshold}pt threshold (max={max_abs_err:.2f}pt)")
            results['h4_answer'] = "NO"
            results['h4_remaining_unknown'] = f"Coordinate formula error of {max_abs_err:.2f}pt remains unexplained"
    
    # Candidate pipeline
    print("\n\nCandidate Pipeline Assessment:")
    
    results['candidates'] = {
        "A (Worksheet→Coords→XML→PDF)": "Possible if export was final step",
        "B (Worksheet→PDF→Coords→XML)": "Likely — coords measured from rendered PDF",
        "C (Worksheet→Print Layout→Coords→PDF)": "Possible — coords from layout object",
        "D (Intermediate Layout→Coords→PDF→XML)": "Unlikely — extra complexity"
    }
    
    if comparisons:
        best_candidate = "B"
        if results.get('h2_answer', '').startswith('worksheet'):
            best_candidate = "C"
        results['recommended_candidate'] = best_candidate
        print(f"  Recommended: Candidate {best_candidate}")
    
    # Save pipeline reconstruction
    pipeline_path = os.path.join(OUT, "pipeline_reconstruction.json")
    with open(pipeline_path, 'w') as f:
        json.dump(results, f, indent=2, default=str)
    
    print("\nInvestigation H complete.")
    return results

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN — Run all investigations
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    import sys
    
    print("Phase 11J.18 — Reconstruct the Original ConMas Generation Pipeline")
    print("=" * 70)
    
    # Phase A: Fresh extraction
    metadata = investigation_a()
    if not metadata:
        print("FATAL: Investigation A failed")
        return
    
    # Phase B: Workbook fingerprint
    investigation_b()
    
    # Phase C: Fresh Excel PDF export
    investigation_c()
    
    # Phase D: PDF Forensics
    pdf_compare = investigation_d()
    
    # Phase E: Overlay validation + render comparison
    geometry = investigation_e(metadata)
    render_comparison()
    
    # Phase F: COM pipeline validation
    comparisons = investigation_f(metadata)
    
    # Phase G: XML correlation
    xml_correlation = investigation_g(metadata)
    
    # Phase H: Pipeline reconstruction
    pipeline_results = investigation_h(metadata, pdf_compare, comparisons, xml_correlation)
    
    # ─── Final Summary ───────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("FINAL SUMMARY")
    print("=" * 70)
    
    print("\n1. Is def_file the exact workbook used to generate the stored PDF?")
    print(f"   Answer: {pipeline_results.get('h1_answer', 'INCONCLUSIVE')}")
    
    print("\n2. Were def_cluster coordinates generated from...")
    print(f"   Answer: {pipeline_results.get('h2_answer', 'INCONCLUSIVE')}")
    
    print("\n3. Are XML and database from the same coordinate source?")
    print(f"   Answer: {pipeline_results.get('h3_answer', 'INCONCLUSIVE')}")
    
    print("\n4. Can legacy generation be reproduced deterministically?")
    print(f"   Answer: {pipeline_results.get('h4_answer', 'INCONCLUSIVE')}")
    
    if pipeline_results.get('h4_answer') == "YES":
        print(f"   Max error: {pipeline_results.get('h4_max_error_pt', '?')}pt")
        print(f"   Avg error: {pipeline_results.get('h4_avg_error_pt', '?')}pt")
    else:
        unknown = pipeline_results.get('h4_remaining_unknown', 'Unknown')
        print(f"   Remaining unknown: {unknown}")
    
    print(f"\nAll artifacts in: {OUT}")
    print("Phase 11J.18 complete.")

if __name__ == '__main__':
    main()
