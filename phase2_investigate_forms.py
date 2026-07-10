# -*- coding: utf-8 -*-
"""Investigate regressed forms for Phase 2 worksheet resolution fix."""
import psycopg2, zipfile, re
from io import BytesIO
from xml.etree import ElementTree as ET

conn = psycopg2.connect(host='127.0.0.1', port=5432,
                        dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

def deep_investigate(fid):
    print(f"\n{'='*70}")
    print(f"Form {fid}")
    print(f"{'='*70}")

    # Get XLSX and xml_data
    cur.execute("SELECT def_file, xml_data FROM def_top WHERE def_top_id = %s", (fid,))
    row = cur.fetchone()
    if not row or not row[0]:
        print("  No XLSX file")
        return
    xlsx = bytes(row[0])
    xml_data = str(row[1]) if row[1] else ""

    # Get sheets from DB
    cur.execute("SELECT def_sheet_id, def_sheet_no, def_sheet_name "
                "FROM def_sheet WHERE def_top_id = %s ORDER BY def_sheet_no",
                (fid,))
    sheets = cur.fetchall()
    print(f"  Sheets in DB: {len(sheets)}")
    for sid, snum, sname in sheets[:5]:
        print(f"    Sheet#{snum}: id={sid} name=\"{sname}\"")

    # Get clusters        cur.execute("SELECT cell_address FROM def_cluster WHERE def_top_id = %s AND rownum <= 5 ORDER BY def_cluster_id", (fid,))
    clusters = cur.fetchall()
    print(f"  Got {len(clusters)}+ clusters (showing first 5):")
    for c in clusters[:5]:
        print(f"    Cell: {c[0]}")

    # Parse XLSX
    try:
        with zipfile.ZipFile(BytesIO(xlsx)) as z:
            # List all files
            all_files = z.namelist()
            sheet_files = sorted([n for n in all_files if 'worksheets/sheet' in n.lower()])
            print(f"  XLSX sheet files: {sheet_files}")

            # Read workbook.xml
            wb = ET.fromstring(z.read('xl/workbook.xml'))
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

            # List sheets in workbook.xml
            sheet_elems = wb.findall('.//s:sheet', ns)
            print(f"  Sheets in workbook.xml:")
            for sh in sheet_elems:
                name = sh.get('name', '?')
                rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '?')
                sid = sh.get('sheetId', '?')
                state = sh.get('state', 'visible')
                print(f"    sheetId={sid} name=\"{name}\" r:id={rid} state={state}")

            # Read workbook.xml.rels
            rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            rns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            print(f"  Relationships:")
            sheet_rels = {}
            for rel in rels:
                rid = rel.get('Id', '?')
                target = rel.get('Target', '?')
                if 'worksheet' in target.lower() or 'sheet' in target.lower():
                    sheet_rels[rid] = target
                    print(f"    {rid} -> {target}")

            # Build name-to-path mapping
            name_to_path = {}
            for sh in sheet_elems:
                name = sh.get('name', '?')
                rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '?')
                target = sheet_rels.get(rid, '?')
                path = 'xl/' + target.replace('\\', '/') if target != '?' else '?'
                name_to_path[name] = path
                print(f"    Name \"{name}\" -> {path}")

            # Read cols from each sheet
            print(f"  Column data per sheet:")
            for sf in sheet_files:
                try:
                    sheet_xml = ET.fromstring(z.read(sf))
                    cols = sheet_xml.find('.//s:cols', ns)
                    pa = sheet_xml.find('.//s:printArea', ns)
                    if cols is not None:
                        col_elems = cols.findall('s:col', ns)
                        widths = [(int(c.get('min', '0')), int(c.get('max', '0')),
                                   float(c.get('width', '0')),
                                   c.get('customWidth', '0') == '1',
                                   c.get('hidden', '0') == '1')
                                  for c in col_elems]
                        total_chars = sum((max-min+1)*w for min,max,w,ch,hd in widths if not hd)
                        total_pts = sum((max-min+1)*(w*7.33+5)*72/96 for min,max,w,ch,hd in widths if not hd)
                        print(f"    {sf}: {len(col_elems)} col defs, total_chars={total_chars:.1f}, total_pt={total_pts:.1f}")
                        for min_c, max_c, w, ch, hd in widths:
                            pt = (w * 7.33 + 5) * 72 / 96
                            print(f"      col {min_c}-{max_c}: width={w:.2f}chars -> {pt:.2f}pt custom={ch} hidden={hd}")
                    else:
                        print(f"    {sf}: NO <cols> element")
                    if pa is not None:
                        pa_str = ET.tostring(pa, encoding='unicode')
                        print(f"    {sf}: has printArea: {pa_str[:150]}")
                except Exception as e:
                    print(f"    {sf}: ERROR={e}")

    except Exception as e:
        print(f"  XLSX parse failed: {e}")

    # Check xml_data for ConMas page settings
    if xml_data:
        ch_match = re.search(r'horizontalCentered\s*=\s*"([^"]*)"', xml_data[:5000])
        if ch_match:
            print(f"  ConMas XML: centerHorizontally={ch_match.group(1)}")
        margin_l = re.search(r'marginLeft\s*=\s*"([^"]*)"', xml_data[:5000])
        if margin_l:
            print(f"  ConMas XML: marginLeft={margin_l.group(1)}")

# Investigate all regressed forms
for fid in [186, 187, 193, 142, 155]:
    deep_investigate(fid)

conn.close()
print("\nDone.")
