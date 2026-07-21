# Phase 22.3 — Style Persistence Trace Report

**Date:** 2026-07-21
**Scope:** Verify that browser style persistence (Phase 22) remains connected after replacing `WorkbookValueWriter` with `ConMasCompatibleWorkbookWriter`.

---

## 1. Pipeline Flow Diagram

```
Browser Style Edit
    ↓  FieldConfig (appearance + layout)
PaperlessDesigner → handleExportExcel(fieldConfigs)
    ↓
handleSaveEdited(fieldConfigs)
    ↓
runtimeFormToWorkbookDefinition(runtimeForm, values, sessionId, fieldConfigs)
    │  calls buildFieldStyle(field, fieldConfigs[f.id]) per field
    ↓
WbDef payload (JSON) — includes style: { font, fill, alignment }
    ↓
POST /api/form/save-edited
    ↓
FormController.SaveEdited()
    ↓  model binding → FieldDefinition.Style = CellStyle
FormSaveService.SaveEditedValuesAsync()
    ├── [1] ConMasCompatibleWorkbookWriter.WriteValues()    ← COM: values ONLY
    ├── [2] WorkbookStyleWriter.ApplyStyles()               ← OpenXML SDK: styles ONLY
    └── [3] CompatibilityValidator.Validate()
    ↓
Download XLSX
    ↓  (user re-uploads)
DesignerModelReader.Read()
    └── ReadCellFormatting()   ← Stage 23: reads styles from workbook
    └── BuildAllCellStylesFromWorkbook() + compare  ← Stage 25: round-trip PASS/FAIL
```

---

## 2. Critical Question: Is Phase 22 Still Connected?

**Answer: YES.** `WorkbookStyleWriter.ApplyStyles()` IS still called.

**Evidence** — `FormSaveService.cs:406-427`:

```
Line 409:  _valueWriter.WriteValues(...)            ← ConMasCompatibleWorkbookWriter
Line 417:  if (definition.Sheets.Any(s => ...))     ← checks for style overrides
Line 419:      _styleWriter.ApplyStyles(...)         ← WorkbookStyleWriter
```

Guard condition at line 417 checks `HasStyleProperties(f.Style)` per field. If ANY field has a non-default style, `ApplyStyles()` executes. If no styles are present, it logs `"[PHASE22] No field style overrides found — skipping style writer"` and moves on.

---

## 3. ConMasCompatibleWorkbookWriter — Actual Writes

**File:** `ConMasCompatibleWorkbookWriter.cs:78-85`

```csharp
range = ws.Range[cellRef];
if (IsNumericField(field) && double.TryParse(field.Value, out double numVal))
    range.Value = numVal;
else
    range.Value = field.Value;
```

| Property | Written? |
|---|---|
| `range.Value` | ✅ Yes |
| `range.Font.Name` | ❌ No |
| `range.Font.Size` | ❌ No |
| `range.Font.Bold` | ❌ No |
| `range.Font.Italic` | ❌ No |
| `range.Font.Color` | ❌ No |
| `range.Interior.Color` | ❌ No |
| `range.HorizontalAlignment` | ❌ No |
| `range.VerticalAlignment` | ❌ No |
| `range.WrapText` | ❌ No |

This is correct by design — the COM writer is value-only. Style persistence is delegated entirely to `WorkbookStyleWriter`.

---

## 4. WorkbookStyleWriter — Actual Writes

**File:** `WorkbookStyleWriter.cs:104-166`

For each field with `field.Style != null && HasStyleProperties(field.Style)`:

| Property | Written? | Mechanism |
|---|---|---|
| Font Name | ✅ | `FindOrCreateFont()` → Font record → `cell.StyleIndex` |
| Font Size | ✅ | `FindOrCreateFont()` → FontSize record |
| Bold | ✅ | `FindOrCreateFont()` → Bold record |
| Italic | ✅ | `FindOrCreateFont()` → Italic record |
| Underline | ✅ | `FindOrCreateFont()` → Underline record |
| Font Color | ✅ | `FindOrCreateFont()` → Color record (Rgb hex) |
| Fill Color | ✅ | `FindOrCreateFill()` → Fill record → PatternFill foreground |
| H-Align | ✅ | CellFormat.Alignment.Horizontal |
| V-Align | ✅ | CellFormat.Alignment.Vertical |
| Wrap Text | ✅ | CellFormat.Alignment.WrapText |

Creates new `CellFormat` (xf) entries in the stylesheet for each unique style combination, deduplicating identical styles. Sets `cell.StyleIndex` to reference the new format.

---

## 5. Frontend — Style Payload

**File:** `paperless/app/page.tsx`

### WbDefStyle Interface (line 11)
```typescript
interface WbDefStyle {
  font?: {
    name?: string; sizePt?: number; bold?: boolean;
    italic?: boolean; underline?: boolean; colorArgb?: string;
  };
  alignment?: { horizontal?: string; vertical?: string; };
  fill?: { patternType?: string; colorArgb?: string; };
  wrapText?: boolean;
}
```

### buildFieldStyle() Mapping (line 366)
| FieldConfig Source | WbDefStyle Property | Status |
|---|---|---|
| `appearance.fontFamily` | `font.name` | ✅ |
| `appearance.fontSize` | `font.sizePt` | ✅ |
| `appearance.fontWeight === "bold"` | `font.bold = true` | ✅ |
| `appearance.textColor` | `font.colorArgb` | ✅ |
| `appearance.backgroundColor` | `fill.colorArgb` (+ patternType="solid") | ✅ |
| `layout.horizontalAlign` | `alignment.horizontal` | ✅ |
| `layout.verticalAlign` | `alignment.vertical` | ❌ **NOT MAPPED** |
| *(italic/underline/wrap from FieldConfig)* | *(not set)* | ❌ Not sent |

### Stage 8 Console Logging (line 511)
Already logs the full style payload immediately before `fetch()`:
```
STAGE 8 — FINAL JSON PAYLOAD (before fetch)
  Field: name='...' cell='A1' value='...'
    Style: font='Meiryo/14' bold=true fill='#FFFFFF00' hAlign='center'
```

---

## 6. Backend Model — Style Definition Path

```
JSON payload:  style.font.name         →  CellStyle.Font.Name         →  Font.FontName.Val
               style.font.sizePt       →  CellStyle.Font.SizePt       →  Font.FontSize.Val
               style.font.bold         →  CellStyle.Font.Bold         →  Font.Bold
               style.font.colorArgb    →  CellStyle.Font.ColorArgb    →  Color.Rgb
               style.fill.colorArgb    →  CellStyle.Fill.ColorArgb    →  PatternFill.ForegroundColor.Rgb
               style.alignment.horizontal → CellStyle.Alignment.Horizontal → Alignment.Horizontal
               style.alignment.vertical →  CellStyle.Alignment.Vertical   →  Alignment.Vertical
```

**Model chain:** `FieldDefinition.Style` → `CellStyle` → `FontDefinition` / `FillDefinition` / `AlignmentDefinition` (all in `StyleDefinition.cs`)

---

## 7. Stage 25 Re-read — Style Round-Trip

**File:** `DesignerModelReader.cs`

### Stage 23 — Read (line 1031)
```csharp
field.Style = new DesignerFieldStyle {
    FontFamily = style.FontName,
    FontSize = style.FontSize,
    Bold = style.Bold,
    Italic = style.Italic,
    Underline = style.Underline,
    FontColor = style.FontColor,
    FillColor = style.FillColor,
    HorizontalAlignment = style.HorizontalAlignment,
    VerticalAlignment = style.VerticalAlignment,
    WrapText = style.WrapText,
    ...
};
```

### Stage 25 — Compare (line 390)
Re-reads workbook via `BuildAllCellStylesFromWorkbook()` and compares with `DesignerField.Style`:

| Property | Workbook Source | Comparison |
|---|---|---|
| Font Name | Font.FontName.Val | PASS/FAIL |
| Font Size | Font.FontSize.Val | PASS/FAIL |
| Bold | Font.Bold | PASS/FAIL |
| Italic | Font.Italic | PASS/FAIL |
| Font Color | Color.Rgb hex | PASS/FAIL |
| Fill Color | PatternFill.ForegroundColor.Rgb | PASS/FAIL |
| H-Align | Alignment.Horizontal | PASS/FAIL |
| V-Align | Alignment.Vertical | PASS/FAIL |
| Wrap Text | Alignment.WrapText | PASS/FAIL |

---

## 8. Pre-Existing Console Logging Inventory

| Stage | Location | Logs |
|---|---|---|
| Stage 6 | `page.tsx:428` | runtimeFormToWorkbookDefinition INPUT |
| Stage 7 | `page.tsx:475` | runtimeFormToWorkbookDefinition OUTPUT (style per field) |
| Stage 8 | `page.tsx:512` | Final JSON payload before fetch (style per field) |
| Stage 1-4 | `FormController.cs:621` | Model binding result, Style.Font logged at line 665 |
| Stage 7 | `FormSaveService.cs:289` | Field style font properties logged at line 377-396 |
| Stage 24 | `WorkbookStyleWriter.cs:148` | Style Write — font/size/bold/color/fill/alignment per cell |
| Stage 23 | `DesignerModelReader.cs:1031` | Style Read — font/size/bold/color/fill/alignment per cell |
| Stage 25 | `DesignerModelReader.cs:390` | Round-trip comparison — PASS/FAIL per property |

---

## 9. Risks and Gaps

| Risk | Severity | Status |
|---|---|---|
| Style writer not called (registration removed) | **Critical** | ✅ Called at `FormSaveService.cs:419` |
| Style writer before value writer (styles overwritten) | **Critical** | ✅ Values first (COM), styles second (SDK) at lines 409→419 |
| Frontend drops all style properties | **High** | ✅ All critical properties sent. Only vertical alignment dropped |
| Style not read back on re-upload | **High** | ✅ `ReadCellFormatting()` reads all properties |
| OpenXML SDK mutations in style writer | **Medium** | Same as old system — no regression |
| `layout.verticalAlign` not mapped in frontend | **Low** | `buildFieldStyle()` line 386-388 only maps `horizontalAlign` |

---

## 10. Acceptance Test Readiness

The pipeline is correctly wired end-to-end:

```
Original Excel (Calibri 11) 
  → Browser edits Font=Meiryo, Size=14, Bold=true, Fill=Yellow, H-Align=Center
  → buildFieldStyle() produces { font: {name:"Meiryo",sizePt:14,bold:true}, fill:{colorArgb:"#FFFFFF00",patternType:"solid"}, alignment:{horizontal:"center"} }
  → JSON sent to POST /api/form/save-edited
  → FieldDefinition.Style populated with CellStyle { Font, Fill, Alignment }
  → ConMasCompatibleWorkbookWriter writes value="John" via COM
  → WorkbookStyleWriter creates Font(Meiryo,14,Bold), Fill(Yellow), CellFormat(Center) via OpenXML SDK
  → cell.StyleIndex set to new xf
  → Downloaded XLSX has: A1="John", A1.Font=Meiryo/14/Bold, A1.Interior=Yellow, A1.HorizontalAlignment=Center
  → Re-upload → DesignerModelReader reads all properties → Stage 25 compares → should PASS
```

**Only gap:** `FieldConfig.layout.verticalAlign` is ignored by `buildFieldStyle()` in `page.tsx:386-388`. Vertical alignment edits from the browser will not persist.
