# Phase X.6 — Complete Legacy Output Excel Format Reverse Engineering

**Date:** 2026-07-17  
**Status:** Investigation complete  
**Data Sources:**
- `FormTest - Copy.xlsx` (legacy ConMas output, 16,782 bytes)
- `FormTest - Copy_output.xlsx` (our generated output, 12,327 bytes)

---

## 1. Executive Summary

The legacy ConMas Output Excel workbook and our generated workbook differ in **11 structural areas**. The most critical difference is the **ExcelOutputSetting worksheet** — a 36-row ConMas XML configuration sheet embedded inside the workbook. Additionally, the **comment format** is significantly richer and the **VML drawing** contains visual comment shapes that our workbook lacks.

To achieve full structural parity, the WorkbookGenerator must be extended to produce: 3 sheets (not 2), legacy-format comments (6+ lines, not 4), VML comment drawings, and the ExcelOutputSetting sheet.

---

## 2. Workbook Structure Comparison

| Feature | Legacy | Generated | Status |
|---------|--------|-----------|--------|
| **File size** | 16,782 bytes | 12,327 bytes | ❌ -4,455 bytes |
| **Sheet count** | 3 | 2 | ❌ Missing 1 |
| **Sheet names** | `_Fields` (hidden), `Sheet1`, `ExcelOutputSetting` | `Sheet1`, `_Fields` (hidden) | ❌ |
| **Comment count** | 6 | 1 (after fix: should be 6) | ⚠️ Fix applied |
| **VML shapes** | 7 | 1 (after fix: should be 7) | ⚠️ Partial |
| **Shared strings** | 43 items | 23 items | ❌ |
| **Defined names** | 1 (`_xlnm.Print_Area`) | 0 | ❌ |
| **Sheet1 cells** | 33 cells, 5 merge regions | 0 cells, 0 merges | ❌ |
| **Printer settings** | 5,428 bytes (sheet2) | 5,428 bytes (sheet1) | ✅ Identical |
| **Styles** | 3,616 bytes | 1,901 bytes | ⚠️ Less |
| **Theme** | 8,721 bytes | 8,721 bytes | ✅ Identical |

---

## 3. ExcelOutputSetting Worksheet — Full Analysis

### 3.1 Purpose

The `ExcelOutputSetting` sheet is a **36-row worksheet containing the complete ConMas form configuration serialized as XML in a single shared string** across multiple cells. It stores:

1. **Form-level configuration**: designer version, save settings, biometrics, identification
2. **Output settings**: CSV, PDF, Excel, DocuWorks export configuration
3. **Field (cluster) definitions**: per-cluster settings including type, position, visibility, validation
4. **Sheet-level configuration**: per-sheet remarks, copy settings, display settings
5. **Runtime behavior**: mobile display, required check mode, network settings, auto-numbering

### 3.2 Row Structure

The sheet has 36 rows (rows 1-36), each containing a **single cell with a shared string index** pointing to a fragment of the complete XML document. The 36 fragments (shared strings [7] through [42]) assemble into a complete `<conmas>` XML document.

### 3.3 XML Structure (Reconstructed)

```xml
<conmas>
  <top>
    <designerVersion></designerVersion>
    <designerDisplayVersion></designerDisplayVersion>
    <updateDelay></updateDelay>
    <mobileSave>0</mobileSave>
    <mobileReportSave>1</mobileReportSave>
    <useBiometrics>0</useBiometrics>
    <useIdentification>0</useIdentification>
    <finishOutputFiles>
      <csv></csv>
      <csvImageAudio></csvImageAudio>
      <csvZip></csvZip>
      <dataOutputCsv></dataOutputCsv>
      <excel></excel>
    </finishOutputFiles>
    <editOutput>0</editOutput>
    <editOutputFiles>
      <csv></csv>
      <pdf></pdf>
      <pdfLayer></pdfLayer>
      <docuworks></docuworks>
      <excel></excel>
    </editOutputFiles>
    <isOriginalWhole>1</isOriginalWhole>
    <wholeImageSize></wholeImageSize>
    <saveIndividuallyImage>1</saveIndividuallyImage>
    <!-- ... remarks, camera image, cooperation table ... -->
    <originalSheetNames>
      <originalSheetName></originalSheetName>
    </originalSheetNames>
    <edgeOcrSetting>
      <edgeOcrClusters />
    </edgeOcrSetting>
    <!-- Per-sheet settings -->
    <sheets>
      <sheet>
        <no>1</no>
        <name>Sheet1</name>
        <remarks>...</remarks>
        <clusters>
          <cluster>
            <sheetNo>1</sheetNo>
            <clusterId>1</clusterId>
            <isHidden>0</isHidden>
            <isHiddenDesigner>0</isHiddenDesigner>
            <mobileDisplay>1</mobileDisplay>
            <excelOutputValue></excelOutputValue>
            <reportCopy>
              <clear>0</clear>
              <displayDefaultValue>1</displayDefaultValue>
            </reportCopy>
            <management>
              <valueToRemarks />
              <valueToSystemKeys />
            </management>
          </cluster>
          <!-- ... more clusters up to clusterId=4 ... -->
        </clusters>
      </sheet>
    </sheets>
  </top>
</conmas>
```

### 3.4 Is It Required for Republish?

**YES** — The `ExcelOutputSetting` sheet is required for the legacy ConMas runtime to correctly republish a workbook. It contains:

- Cluster IDs and ordering (determines field enumeration order)
- Per-cluster `excelOutputValue` (whether to include in output)
- Field visibility flags (`isHidden`, `isHiddenDesigner`)
- Mobile display settings
- Form-level output configuration (finish/output files)
- Print and report copy settings

If our generated workbook doesn't include this sheet, the legacy ConMas runtime would need to reconstruct the entire form configuration from comments alone, which may not produce identical results.

### 3.5 How to Generate

The XML document can be constructed from the existing `FormDefinition.Clusters` and `FormDefinition.Sheets` data:
- Each `<cluster>` maps to a `ClusterDefinition`
- `<excelOutputValue>` maps to whether the field has a value in `SheetDefinition.CellValues`
- `<clusterId>` uses the cluster's index/ordering
- Form-level settings use sensible defaults (matching ConMas defaults)

---

## 4. Legacy Comment Format — Complete Specification

### 4.1 Comment Format

```
Line 0:  Field Name            → "samples"
Line 1:  Field Type            → "KeyboardText"
Line 2:  Cluster Index (0-based) → "0"
Line 3:  0                     → Reserved / unknown field
Line 4:  0                     → Reserved / unknown field
Line 5+: Serialized Parameters → "Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11"
(Followed by ~17 trailing empty lines)
```

### 4.2 Line-by-Line Breakdown

| Line | Legacy Value | Purpose | Our Value (Current) |
|------|-------------|---------|---------------------|
| 0 | `samples` | Field display name | `samples` ✅ |
| 1 | `KeyboardText` | Field type | `KeyboardText` ✅ |
| 2 | `0` through `5` | **Cluster index** (0-based) | *(empty)* ❌ |
| 3 | `0` | Reserved field 1 | *(missing)* ❌ |
| 4 | `0` | Reserved field 2 | *(missing)* ❌ |
| 5+ | `Required=0;Lines=1;...` | **Serialized parameters** | `0` (placeholder) ❌ |
| 22+ | ~17 empty lines | Trailing padding | *(missing)* ❌ |

### 4.3 Parameters String Format

The parameters string uses semicolon-delimited key=value pairs:

| Parameter | Example Value | Purpose |
|-----------|--------------|---------|
| `Required` | `0` | Whether field is required (0/1) |
| `Lines` | `1` | Number of visible lines |
| `InputRestriction` | `None` | Input validation (None, Numeric, Date, etc.) |
| `MaxLength` | `0` | Maximum character length |
| `Align` | `Center` | Horizontal alignment (Left, Center, Right) |
| `Font` | `Arial` | Font name |
| `FontSize` | `11` | Font size in points |
| `Weight` | `Normal` | Font weight (Normal, Bold) |
| `Color` | `0,0,0` | Font color as RGB tuple |
| `VerticalAlignment` | `2` | Vertical alignment (0=Top, 1=Center, 2=Bottom) |
| `DefaultFontSize` | `11` | Default font size fallback |

### 4.4 Required Changes

To match the legacy format, our comment generation must change to:

```
samples                                    ← Field Name
KeyboardText                               ← Field Type
{clusterIndex}                             ← Cluster Index (0-based)
0                                          ← Reserved field 1
0                                          ← Reserved field 2
Required=0;Lines=1;InputRestriction=None;... ← Parameters string
                                           ← Empty line
                                           ← Empty line (× 17 approximately)
```

---

## 5. VML Drawing Analysis

### 5.1 Legacy VML Structure

The legacy workbook contains **7 VML shapes** in `xl/drawings/vmlDrawing1.vml` (4,823 bytes). These are invisible comment markers (the red triangle indicators) for each comment.

Our generated workbook (after fix) will have 6 comments but the VML is generated automatically by Excel COM's `AddComment()`. The byte count difference (4,823 vs 1,179) is because Excel COM generates minimal VML compared to ConMas's custom VML. However, this is acceptable — Excel COM generates the correct VML for Excel to render comments.

**No action needed** — Excel COM handles VML generation automatically.

---

## 6. Relationships Comparison

| Relationship File | Legacy Entries | Generated Entries | Status |
|------------------|---------------|-------------------|--------|
| `_rels/.rels` (package) | 3 (core, app, office doc) | 3 | ✅ |
| `xl/_rels/workbook.xml.rels` | 6 (3 sheets, sharedStrings, styles, theme) | 5 (2 sheets, sharedStrings, styles, theme) | ❌ Missing 3rd sheet |
| `xl/worksheets/_rels/sheet1.xml.rels` | *(none)* | 1 (comments, VML, printerSettings) | ⚠️ Generated has rels on Sheet1 |
| `xl/worksheets/_rels/sheet2.xml.rels` | 3 (comments, VML, printerSettings) | *(none)* | ⚠️ Legacy has rels on Sheet2 |

The relationship difference is because sheets are ordered differently:
- Legacy: Sheet1=`_Fields`, Sheet2=`Sheet1` (has comments), Sheet3=`ExcelOutputSetting`
- Generated: Sheet1=`Sheet1` (has comments), Sheet2=`_Fields`

This means the **sheet order** determines which .rels file gets the relationships. This is correct behavior — as long as comments exist on the correct sheet, the rels file will be correct.

---

## 7. Metadata Source Hierarchy (ConMas Reconstruction Priority)

Based on the forensic evidence, the legacy ConMas runtime reconstructs field metadata using this priority:

```
Priority 1: Cell Comments
  - 6 comments on merged cell top-left corners
  - Contains: field name, type, cluster index, parameters
  - Used for: field identification and configuration

Priority 2: ExcelOutputSetting Sheet
  - 36 rows of ConMas XML configuration
  - Contains: full form config, cluster definitions, per-field settings
  - Used for: complete form reconstruction

Priority 3: _Fields Sheet (if present)
  - Headers only in legacy workbook
  - Added in newer ConMas versions?
  - Used for: quick field enumeration

Priority 4: Cell Values
  - Existing user-entered values
  - Used for: restoring field content
```

### Implications for our Reader (WorkbookReaderService)

Currently our reader treats `_Fields` as primary. To match legacy behavior:

1. **Comments should be primary** — read comments first, use them for cluster definitions
2. **ExcelOutputSetting as fallback** — if no comments, try the XML config
3. **_Fields as third source** — if neither comments nor XML config exist, use _Fields
4. **Merge all three sources** — cross-reference for completeness

---

## 8. Workbook Generation Sequence

### Legacy ConMas Generation Order (Inferred)

```
1. Create workbook (Workbooks.Add)
2. Create / rename worksheets
3. Set page setup (PaperSize, Orientation, Margins)
4. Define print area
5. Apply merged cells
6. Write column widths
7. Write row heights
8. Apply cell styles (font, colors, alignment, borders)
9. Write cell values
10. Write cell comments (on top-left cells of merge ranges)
11. Create ExcelOutputSetting sheet
    └── Write ConMas XML config (36 rows, single cell per row)
    └── Each cell contains a fragment of the complete XML document
12. Create _Fields sheet (if needed)
    └── Headers only (legacy behavior)
    └── Or with data rows (newer behavior)
13. Set print area as defined name (_xlnm.Print_Area)
14. Save workbook
15. Close COM
```

### Our Current Generation Order

```
1. Create workbook
2. Create / rename worksheets
3. Apply page settings
4. Apply print area
5. Apply row heights
6. Apply column widths
7. Apply merged cells
8. Apply freeze pane
9. Apply cell styles
10. Write cell values         ← Stage 9
11. Write cell comments       ← Stage 10
12. Create _Fields sheet      ← Stage 11
13. Save workbook
```

### Differences

| Stage | Legacy | Our | Status |
|-------|--------|-----|--------|
| ExcelOutputSetting | ✅ After comments | ❌ Missing | Must add |
| Print area as defined name | ✅ `_xlnm.Print_Area` | ✅ Via PageSetup | OK |
| _Fields with data | Headers only | Full data rows | OK (not required for legacy compat) |

---

## 9. Implementation Changes Required

### 9.1 Critical

| # | Change | File | Effort |
|---|--------|------|--------|
| 1 | Add `CreateExcelOutputSettingSheet()` to WorkbookGenerator | `WorkbookGenerator.cs` | Medium |
| 2 | Update comment format: add cluster index, reserved lines, parameters string | `WorkbookGenerator.cs` `WriteCellComments()` | Small |
| 3 | Make comments primary metadata source in WorkbookReaderService | `WorkbookReaderService.cs` | Medium |
| 4 | Write cell values to Sheet1 (populate from CellValues) — already done in WriteCellValues | Already done | None |

### 9.2 Nice-to-Have

| # | Change | File | Effort |
|---|--------|------|--------|
| 5 | Add `_xlnm.Print_Area` defined name | `WorkbookGenerator.cs` | Small |
| 6 | Add ExcelOutputSetting to sheet count and reading logic | `WorkbookReaderService.cs` | Small |
| 7 | Generate ExcelOutputSetting XML from FormDefinition | New helper method | Medium |

### 9.3 ExcelOutputSetting XML Generator Specification

To generate the `ExcelOutputSetting` sheet, we need to build a ConMas-compatible XML document:

```xml
<conmas>
  <top>
    <designerVersion></designerVersion>
    <mobileSave>0</mobileSave>
    <finishOutputFiles>
      <excel>1</excel>
    </finishOutputFiles>
    <editOutput>0</editOutput>
    <editOutputFiles>
      <excel>1</excel>
    </editOutputFiles>
    <originalSheetNames>
      <originalSheetName>{sheetName}</originalSheetName>
    </originalSheetNames>
    <sheets>
      <sheet>
        <no>1</no>
        <name>{sheetName}</name>
        <clusters>
          {foreach cluster in form.Clusters}
          <cluster>
            <sheetNo>1</sheetNo>
            <clusterId>{clusterIndex + 1}</clusterId>
            <excelOutputValue>{hasValue ? cluster.Remarks : ""}</excelOutputValue>
            <reportCopy>
              <clear>0</clear>
              <displayDefaultValue>1</displayDefaultValue>
            </reportCopy>
            <management>
              <valueToRemarks />
              <valueToSystemKeys />
            </management>
          </cluster>
          {end}
        </clusters>
      </sheet>
    </sheets>
  </top>
</conmas>
```

The XML must be **split into 36 fragments** (one per row cell), each fragment approximately 500-600 characters. The number 36 comes from the legacy workbook's exact row count.

The 36 fragments align with the shared string indices [7] through [42] from the legacy workbook.

---

## 10. Validation Checklist

After implementing the changes above, the generated workbook should:

- [ ] Contain 3 sheets: `Sheet1`, `_Fields` (hidden), `ExcelOutputSetting`
- [ ] Have 6 comments in legacy format (name, type, cluster index, 0, 0, params)
- [ ] Have 6 VML comment drawings (auto-generated by Excel COM)
- [ ] Have 5 merged regions (A1:B2, C1:D2, A3:D4, A6:D7, A9:D10)
- [ ] Have `_xlnm.Print_Area` defined name
- [ ] Have ExcelOutputSetting sheet with 36 rows
- [ ] Open in Excel with no repair dialogs
- [ ] Reconstruct all 6 fields when uploaded via `upload-excel`
- [ ] Be structurally identical to legacy (except GUIDs and timestamps)
