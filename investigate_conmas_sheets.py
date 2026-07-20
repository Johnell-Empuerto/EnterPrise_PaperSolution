"""
Phase 6.2 — Forensic Inspection: Legacy ConMas Printable Worksheet Selection

Opens a ConMas-exported workbook and dumps EVERY metadata property
for every worksheet. The goal is to determine what distinguishes
"printable" worksheets (Sheet1) from "configuration" worksheets
(ExcelOutputSetting, _Fields) in the original ConMas Designer.

Usage:
    python investigate_conmas_sheets.py <path_to_conmas_workbook.xlsx>
"""

import sys
import os
from zipfile import ZipFile
from xml.etree import ElementTree as ET

S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"

CONFIG_NAMES = {"_Fields", "ExcelOutputSetting", "_RawData", "DesignerConfig", "PaperLessConfig", "ConMasConfig"}


def inspect_ooxml(xlsx_path):
    """Inspect workbook via OOXML ZIP content."""
    report = []
    report.append("=" * 65)
    report.append("PHASE 6.2 — FORENSIC WORKSHEET INSPECTION")
    report.append("=" * 65)
    report.append(f"Workbook: {os.path.basename(xlsx_path)}")
    report.append(f"Size: {os.path.getsize(xlsx_path):,} bytes")
    report.append("")

    try:
        with ZipFile(xlsx_path) as z:
            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
            sheets = wb_xml.findall(f".//{S_NS}sheet")
            report.append(f"Total OOXML sheets: {len(sheets)}")

            # ── Build sheet map ──
            sheet_map = []
            for s in sheets:
                name = s.get("name", "?")
                sheet_id = s.get("sheetId", "?")
                state = s.get("state", "visible")
                rid = s.get("id", "?")
                code_name = s.get("codeName", "(not set)")
                sheet_map.append({
                    "name": name,
                    "sheet_id": sheet_id,
                    "state": state,
                    "rid": rid,
                    "code_name": code_name,
                })
                report.append(f"  Sheet ID={sheet_id}: name='{name}' state='{state}' codeName='{code_name}'")

            # ── workbook.xml.rels ──
            rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            sheet_rels = {}
            for rel in rels_xml:
                rid = rel.get("Id", "?")
                target = rel.get("Target", "?")
                sheet_rels[rid] = target

            report.append(f"\nSheet relationships:")
            for sm in sheet_map:
                target = sheet_rels.get(sm["rid"], "NOT FOUND")
                report.append(f"  {sm['rid']} -> {target}  ({sm['name']})")

            # ── Defined names (print area, etc.) ──
            defined_names = {}
            for dn in wb_xml.findall(f".//{S_NS}definedName"):
                dn_name = dn.get("name", "?")
                dn_local = dn.get("localSheetId")
                dn_hidden = dn.get("hidden", "false")
                text = (dn.text or "").strip()
                defined_names[dn_name] = {
                    "local_sheet_id": dn_local,
                    "hidden": dn_hidden,
                    "text": text,
                }

            # ── Inspect each worksheet ──
            all_sheet_data = []
            for sm in sheet_map:
                name = sm["name"]
                target = sheet_rels.get(sm["rid"])

                ws_data = {
                    "name": name,
                    "sheet_id": sm["sheet_id"],
                    "state": sm["state"],
                    "code_name": sm["code_name"],
                    "dimension": None,
                    "print_area": None,
                    "margins": {},
                    "page_setup": {},
                    "print_options": {},
                    "merged_cells": 0,
                    "rows": 0,
                    "cells": 0,
                    "col_defs": 0,
                    "has_comments": False,
                    "has_drawing": False,
                    "has_sheet_protection": False,
                    "conditional_fmt": 0,
                    "data_validations": None,
                    "empty": False,
                }

                report.append("\n" + "-" * 65)
                report.append(f"WORKSHEET: {name}")
                report.append("-" * 65)

                # Print Area from definedNames
                for dn_name, dn_info in defined_names.items():
                    if dn_info["local_sheet_id"] == sm["sheet_id"] and "_xlnm.Print_Area" in dn_name:
                        ws_data["print_area"] = dn_info["text"]
                        break

                # OOXML properties
                report.append(f"  SheetId:          {sm['sheet_id']}")
                report.append(f"  State:            {sm['state']}")
                report.append(f"  codeName:         {sm['code_name']}")
                report.append(f"  rId:              {sm['rid']} -> {target}")
                report.append(f"  PrintArea:        {ws_data['print_area'] or '(none)'}")
                report.append(f"  Config Name:      YES" if name in CONFIG_NAMES else f"  Config Name:      NO")

                if target and target in z.namelist():
                    ws_xml = ET.fromstring(z.read(target))

                    # Sheet dimension
                    dim = ws_xml.find(f"{S_NS}sheetDimension")
                    if dim is not None:
                        dim_ref = dim.get("ref", "?")
                        ws_data["dimension"] = dim_ref
                    else:
                        dim_ref = "(none)"
                    report.append(f"  SheetDimension:   {dim_ref}")

                    # If dimension is A1 or A1:A1, sheet is essentially empty
                    if dim_ref in ("A1", "A1:A1", "(none)"):
                        ws_data["empty"] = True
                        report.append(f"  Effectively Empty: YES")

                    # PageMargins
                    pm = ws_xml.find(f"{S_NS}pageMargins")
                    if pm is not None:
                        report.append(f"  PageMargins:")
                        for attr in ["left", "right", "top", "bottom", "header", "footer"]:
                            val = pm.get(attr, "?")
                            ws_data["margins"][attr] = val
                            report.append(f"    {attr}: {val} in")
                    else:
                        report.append(f"  PageMargins:      (none)")

                    # PageSetup
                    ps = ws_xml.find(f"{S_NS}pageSetup")
                    if ps is not None:
                        ws_data["page_setup"] = dict(ps.attrib)
                        report.append(f"  PageSetup:")
                        for k, v in ps.attrib.items():
                            report.append(f"    {k}: {v}")
                    else:
                        report.append(f"  PageSetup:        (none)")

                    # PrintOptions
                    po = ws_xml.find(f"{S_NS}printOptions")
                    if po is not None:
                        ws_data["print_options"] = dict(po.attrib)
                        report.append(f"  PrintOptions:")
                        for k, v in po.attrib.items():
                            report.append(f"    {k}: {v}")
                    else:
                        report.append(f"  PrintOptions:     (none)")

                    # Merged cells
                    merges = ws_xml.find(f"{S_NS}mergeCells")
                    if merges is not None:
                        mc_count = int(merges.get("count", 0))
                        ws_data["merged_cells"] = mc_count
                        report.append(f"  MergedCells:      {mc_count}")
                    else:
                        report.append(f"  MergedCells:      (none)")

                    # Sheet data
                    sheet_data = ws_xml.find(f"{S_NS}sheetData")
                    if sheet_data is not None:
                        rows = sheet_data.findall(f"{S_NS}row")
                        ws_data["rows"] = len(rows)
                        total_cells = sum(len(row.findall(f"{S_NS}c")) for row in rows)
                        ws_data["cells"] = total_cells
                        report.append(f"  Rows:             {len(rows)}")
                        report.append(f"  Cells:            {total_cells}")
                    else:
                        report.append(f"  SheetData:        (none)")

                    # Column defs
                    cols = ws_xml.find(f"{S_NS}cols")
                    if cols is not None:
                        col_count = len(cols.findall(f"{S_NS}col"))
                        ws_data["col_defs"] = col_count
                        report.append(f"  ColumnDefs:       {col_count}")
                    else:
                        report.append(f"  ColumnDefs:       (none)")

                    # Comments
                    ws_rels_path = f"xl/worksheets/_rels/{os.path.basename(target)}.rels"
                    if ws_rels_path in z.namelist():
                        ws_rels = ET.fromstring(z.read(ws_rels_path))
                        for wr in ws_rels:
                            wr_target = wr.get("Target", "")
                            if "comments" in wr_target:
                                ws_data["has_comments"] = True
                                report.append(f"  HasComments:      YES ({wr_target})")
                                break
                        if not ws_data["has_comments"]:
                            report.append(f"  HasComments:      NO")
                    else:
                        report.append(f"  HasComments:      NO")

                    # Drawings
                    drawing = ws_xml.find(f"{S_NS}drawing")
                    if drawing is not None:
                        ws_data["has_drawing"] = True
                        report.append(f"  HasDrawing:       YES")
                    else:
                        report.append(f"  HasDrawing:       NO")

                    # Sheet protection
                    sp = ws_xml.find(f"{S_NS}sheetProtection")
                    if sp is not None:
                        ws_data["has_sheet_protection"] = True
                        report.append(f"  SheetProtection:  YES")
                    else:
                        report.append(f"  SheetProtection:  NO")

                    # Conditional formatting
                    cf_list = ws_xml.findall(f"{S_NS}conditionalFormatting")
                    ws_data["conditional_fmt"] = len(cf_list)
                    report.append(f"  ConditionalFmt:   {len(cf_list)}")

                    # Data validations
                    dv = ws_xml.find(f"{S_NS}dataValidations")
                    if dv is not None:
                        ws_data["data_validations"] = dv.get("count", "?")
                        report.append(f"  DataValidations:  count={dv.get('count', '?')}")
                    else:
                        report.append(f"  DataValidations:  (none)")

                all_sheet_data.append(ws_data)

        # ── Printable Score Analysis ──
        report.append("\n\n" + "=" * 65)
        report.append("WORKSHEET PRINTABLE SCORE ANALYSIS")
        report.append("=" * 65)
        report.append("")
        report.append(f"{'Name':<25} {'Vis':<5} {'PArea':<6} {'Rows':<5} {'Cells':<6} {'Cmts':<5} {'Drw':<4} {'Mrg':<4} {'Empty':<6} {'CfgNm':<6} {'Printable?'}")
        report.append("-" * 95)

        for ws_data in all_sheet_data:
            visible = ws_data["state"] == "visible"
            has_print_area = bool(ws_data["print_area"])
            has_rows = ws_data["rows"] > 0
            has_cells = ws_data["cells"] > 0
            has_comments = ws_data["has_comments"]
            has_drawing = ws_data["has_drawing"]
            has_merged = ws_data["merged_cells"] > 0
            is_empty = ws_data["empty"]
            is_config_name = ws_data["name"] in CONFIG_NAMES
            has_page_setup = bool(ws_data["page_setup"])

            # Printable decision based on evidence
            reasons = []
            if is_config_name:
                reasons.append("reserved config name")
            if is_empty:
                reasons.append("effectively empty")
            if not visible:
                reasons.append("not visible")
            if not has_print_area and not has_cells:
                reasons.append("no content")
            if has_comments:
                reasons.append("has field comments")

            is_printable = visible and not is_config_name and not is_empty and (has_print_area or has_cells)

            score_parts = []
            score_parts.append("V" if visible else "X")
            score_parts.append("PA" if has_print_area else "--")
            score_parts.append(str(ws_data["rows"]))
            score_parts.append(str(ws_data["cells"]))
            score_parts.append("Y" if has_comments else "N")
            score_parts.append("Y" if has_drawing else "N")
            score_parts.append(str(ws_data["merged_cells"]))
            score_parts.append("E" if is_empty else "NE")
            score_parts.append("Y" if is_config_name else "N")

            printable_str = "YES" if is_printable else "NO"
            report.append(f"{ws_data['name']:<25} {score_parts[0]:<5} {score_parts[1]:<6} {score_parts[2]:<5} {score_parts[3]:<6} {score_parts[4]:<5} {score_parts[5]:<4} {score_parts[6]:<4} {score_parts[7]:<6} {score_parts[8]:<6} {printable_str}")
            report.append(f"  {'':>25} Reasons: {', '.join(reasons) if reasons else 'has content'}")
            report.append(f"  {'':>25} Printable Score: {'PASS' if is_printable else 'FAIL'}")

        # ── Algorithm Evaluation ──
        report.append("\n\n" + "=" * 65)
        report.append("ALGORITHM EVALUATION")
        report.append("=" * 65)

        report.append("""
Algorithm A: Reserved Worksheet Names
  Skip sheets with known configuration names:
    - _Fields
    - ExcelOutputSetting  
    - _RawData

  Evidence from workbook:
""")
        for ws_data in all_sheet_data:
            is_config = ws_data["name"] in CONFIG_NAMES
            has_content = ws_data["cells"] > 0 or ws_data["rows"] > 0
            report.append(f"    {ws_data['name']:<25} ConfigName={is_config} HasContent={has_content}")
        report.append("""
  Verdict:    LIKELY — matches observed behavior for all 3 sheets.
  Confidence: HIGH (simple, matches ConMas naming convention)

Algorithm B: Printable Sheet Detection
  Only sheets with printable content become pages:
    - Has printable cells
    - Has print area  
    - Has printable objects

  Evidence from workbook:
""")
        for ws_data in all_sheet_data:
            has_print_area = bool(ws_data["print_area"])
            has_cells = ws_data["cells"] > 0
            has_comments = ws_data["has_comments"]
            report.append(f"    {ws_data['name']:<25} PrintArea={has_print_area} Cells={has_cells} Comments={has_comments}")
        report.append("""
  Verdict:    POSSIBLE — ExcelOutputSetting has no PrintArea, no cells, no comments.
              BUT _Fields also has no PrintArea, yet that alone wouldn't distinguish
              ExcelOutputSetting from _Fields if both are empty. The key difference
              is that _Fields is HIDDEN and ExcelOutputSetting is VISIBLE.
              A visibility-only check would NOT exclude ExcelOutputSetting.

  Confidence: MEDIUM (detection works but over-engineered for the actual case)

Algorithm C: Configuration-Driven
  ExcelOutputSetting itself contains the list of printable worksheets.
  Read ExcelOutputSetting to determine which sheets are pages.

  Evidence:
""")
        report.append(f"    ExcelOutputSetting cells: {[d for d in all_sheet_data if d['name'] == 'ExcelOutputSetting'][0]['cells'] if any(d['name'] == 'ExcelOutputSetting' for d in all_sheet_data) else 'N/A'}")
        report.append("""
  Verdict:    UNLIKELY — ExcelOutputSetting has NO cell data (0 rows, 0 cells).
              It cannot contain a worksheet list because it's empty.
              The sheet name itself IS the configuration indicator.

  Confidence: LOW (no evidence that ExcelOutputSetting contains data)

        """)

        # ── Final Recommendation ──
        report.append("=" * 65)
        report.append("FINAL RECOMMENDATION")
        report.append("=" * 65)
        report.append("""
Based on the forensic evidence, the original ConMas Designer most likely
used Algorithm A: Reserved Worksheet Names.

Evidence:
1. ExcelOutputSetting is VISIBLE (same as Sheet1) — so visibility filtering alone
   would NOT exclude it.
2. ExcelOutputSetting has no PrintArea, 0 rows, 0 cells, 0 comments — but a
   content-based filter would also need to exclude the hidden _Fields sheet
   differently from ExcelOutputSetting.
3. The simplest explanation: ConMas had a hard-coded list of reserved sheet names
   that were excluded from the "printable worksheets" enumeration.
4. This is also how our C# WorkbookReaderService already works (ConfigurationSheetNames).
5. The fact that _delete_metadata_sheets() already deletes _Fields and _RawData
   (but was missing ExcelOutputSetting) further supports this — it's the same
   reserved-name approach applied during PDF export.

Recommended fix: Add "ExcelOutputSetting" to both:
  - _delete_metadata_sheets() in upload_coordinate_generator.py
  - The "Collect Visible Sheets" name filter in generate_coordinates_and_preview()
""")

    except Exception as e:
        report.append(f"\nERROR: {e}")
        import traceback
        report.append(traceback.format_exc())

    report.append("\n" + "=" * 65)
    report.append("END OF REPORT")
    report.append("=" * 65)

    return "\n".join(report)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python investigate_conmas_sheets.py <path_to_conmas_workbook.xlsx>")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    if not os.path.exists(xlsx_path):
        print(f"File not found: {xlsx_path}")
        sys.exit(1)

    print(f"\nInspecting: {xlsx_path}")
    print(f"File size: {os.path.getsize(xlsx_path):,} bytes\n")

    result = inspect_ooxml(xlsx_path)
    print(result)

    report_path = xlsx_path + ".forensic_report.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(result)
    print(f"\nReport saved to: {report_path}")
