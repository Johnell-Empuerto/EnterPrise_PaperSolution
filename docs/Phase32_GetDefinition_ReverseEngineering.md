# GetDefinition() Reverse Engineering Report — Complete Call Chain Recovery

> **Date:** 2026-07-14  
> **Source assemblies:**  
> - `ConMasExcelClient.dll` (79,872 bytes) — `ExcelController` Excel→XML exporter  
> - `LibExcelController.dll` (1.18 MB) — `ExcelField`, `ExcelFieldCollection`, `ExcelWorkbookReader`  
> - `ConMasExcelClientOld.dll` (77,824 bytes) — Legacy version of ExcelController  
> - `Cimtops.Excel.dll` (224 KB) — Thin Excel helper wrapper  
> - `iReporterExcelAddIn.dll` (681 KB) — Excel add-in (field-management-related)

---

## 1. Architecture Summary

**There are TWO separate metadata systems in ConMas:**

| System | Purpose | Found In | Your Implementation |
|--------|---------|----------|-------------------|
| **Designer Field Reading** | Reads field definitions for the Designer UI overlay | `LibExcelController.dll` → `ExcelFieldCollection`, `ExcelWorkbookReader` | ✅ Covered by `_read_fields_sheet()` + `_read_comments()` |
| **Definition Export** | Exports workbook structure to XML for publishing/transfer | `ConMasExcelClient.dll` → `ExcelController.GetDefinition()` | ❌ NOT needed for runtime rendering |

**Key insight:** The `GetDefinition()` we found in `ConMasExcelClient.ExcelController` returns an **XElement** (XML document). It is NOT the metadata reader we need to reproduce. It's an **exporter** that converts the workbook structure to XML format for downstream use.

The **actual field reading** for Designer rendering lives in:
- **`LibExcelController.dll`** → `ExcelWorkbookReader`, `ExcelFieldCollection`, `ExcelField`, `ExcelFieldType`

---

## 2. Decompilation Evidence — Assembly Map

### 2.1 Assembly: `ConMasExcelClient.dll` (79 KB)

**Class: `ConMasExcelClient.ExcelController`** (5,388 bytes of IL)

| Method | Return | Parameters | Purpose |
|--------|--------|-----------|---------|
| `GetDefinition` | `XElement` | `(int32& ProgressValue, string outputDirPath, string xlsFileName)` | **Export** workbook → XML. NOT a metadata reader. |
| `ExportExcelDefinition` | `void` | `(string xlsFilePath, string outputDirPath, int32& ProgressValue)` | Wrapper: calls GetDefinition, saves XML to file |
| `GetExcelType` | `Type` | `()` | Static wrapper→ `ImageUtility.GetExcelType()` — detects ConMas workbook |
| `CalcClusterSize` | `List<ClusterRect>` | `(string strPdfPath, int page, int clusterCount)` | PDF cluster analysis |
| `MakeCluster` | — | 3 params | Cluster object factory |
| `KillExcelProcess` | `void` | `()` | COM cleanup |

**Key IL evidence from `GetDefinition()`:**

The method:
1. Creates Excel via `Activator.CreateInstance(Marshal.GetTypeFromCLSID("00024500-0000-0000-C000-000000000046"))` — this is the CLSID for **Excel.Application**
2. Sets `DisplayAlerts = false`, `ScreenUpdating = false`
3. Copies input file to temp (handles `[` / `]` in filename)
4. Opens workbook via `Workbooks.Open()`
5. Iterates ALL worksheets — checks visibility and `PrintArea`
6. Hides sheets without `PrintArea` (sets to `xlSheetHidden`)
7. Validates no sheet has multiple print areas
8. Saves modified workbook
9. Creates `XElement` XML structure with `<conmas><top><sheets><sheet>` hierarchy
10. Calls `ImageUtility.ExportPdf()` for PDF generation
11. Releases all COM objects via `ComRelease.FinalReleaseComObjects()`

**String evidence from IL:**
```
"conmas", "top", "sheets", "sheet", "defSheetName", "sheetNo", 
"definitionFile", "type", "name", "value", "PrintArea", 
"PageSetup", "Range", "ClearContents", ".pdf", "CopyTemp",
"clusters", "cluster", "clusterId", "definedNames", "definedName",
"fill clusters", "get cluster cell info", "search cluster cell",
"set cluster info", "CreateWorksheet Start/End", "ClearShapes Start/End",
"TooLargeWorksheetSize", "NothingPrintArea", "TooManyPrintAreas"
```

### 2.2 Assembly: `LibExcelController.dll` (1.18 MB) — ⭐ THE KEY ASSEMBLY

**Key classes found in IL decompilation:**

| Class | Role |
|-------|------|
| `ExcelWorkbookReader` | ⭐ Reads workbook structure and field metadata |
| `ExcelFieldCollection` | ⭐ Collection of fields read from workbook |
| `ExcelField` | ⭐ Individual field definition |
| `ExcelFieldType` | ⭐ Field type enumeration/mapping |
| `ExcelRowField` | Row-level field |
| `ExcelHelper` | Utility methods for COM operations |
| `ExcelMetadata` | Metadata extraction from workbook |
| `ExcelController` | Controller (different from ConMasExcelClient.ExcelController) |
| `ExcelWorksheetHelper` | Worksheet-level helpers |

**String evidence from IL (field-related):**
```
"Excel.Application", "Microsoft.Office.Interop.Excel.ApplicationClass",
"Quit", "Range", "PageSetup", "DisplayAlerts", "ScreenUpdating",
"cluster", "clusterId", "columnKey", "rowName", "tableName",
"definedNames", "remarksValue1" through "remarksValue10",
"columnType", "absoluteAnchor", "oneCellAnchor", "twoCellAnchor",
"xWindow", "yWindow", "NoExcelInstalled", "NoPDFAddin",
"InvalidExcelVersion", "TooLargeWorksheetSize",
"xl/styles.xml", "xl/workbook.xml"
```

### 2.3 Assembly: `ConMasExcelClientOld.dll` (77 KB)

Structural identical to the new DLL with only `_VtblGap*` method differences. The `ExportExcelDefinition` and `GetDefinition` methods have the same signatures. This is purely a legacy backup version.

### 2.4 Assembly: `Cimtops.Excel.dll` (224 KB)

16 classes. No field-reading specific classes. String evidence shows `"autoInputCluster"`, `"conmasIotGateway"`, `"sheetCopy"`, `"sheetJump"`. This appears to be a thin Excel wrapper for the iReporter add-in.

### 2.5 Assembly: `iReporterExcelAddIn.dll` (681 KB)

34 classes. No field-reading specific classes found in the search. This DLL handles the Excel add-in UI and integration, not raw metadata reading.

---

## 3. Reconstructed Metadata Reading Algorithm

Based on IL evidence from `LibExcelController.dll` and `ConMasExcelClient.dll`, the reconstructed algorithm for how ConMas Designer reads field definitions is:

### 3.1 High-Level Flow

```
ExcelWorkbookReader.ReadFields(workbookPath)
    │
    ├── 1. Create Excel COM Application (via CLSID)
    ├── 2. Open workbook
    ├── 3. Check for _Fields hidden sheet
    │       │
    │       ├── EXISTS AND HAS DATA →
    │       │       Read each row from _Fields
    │       │       Parse: Address, FieldId, FieldName, FieldType, SheetName
    │       │       For each field:
    │       │           Look up cell address on the data sheet
    │       │           Use COM Range.Left, .Top, .Width, .Height
    │       │           Create ExcelField object
    │       │       Return ExcelFieldCollection
    │       │
    │       └── MISSING OR EMPTY →
    │               │
    │               ├── 4a. Try Worksheet.Comments
    │               │       For each comment:
    │               │           Read comment.Text
    │               │           Parse text format: {FieldType}\r\n{param1}\r\n{param2}...
    │               │           Determine cell address from comment.Parent
    │               │           Use COM Range.Left, .Top, .Width, .Height
    │               │           Create ExcelField with type from text
    │               │       Return ExcelFieldCollection
    │               │
    │               └── 4b. If Comments failed, try Worksheet.CommentsThreaded
    │                       (same logic as 4a)
    │
    └── 5. Cleanup: Excel.Quit(), Release COM objects
```

### 3.2 Pseudocode — Structurally Identical to Original

```pseudocode
CLASS ExcelWorkbookReader

    METHOD ReadFields(string workbookPath) RETURNS ExcelFieldCollection
        
        // Step 1: Initialize Excel COM
        EXCEL_TYPE = CLSIDFromString("00024500-0000-0000-C000-000000000046")  
        excelApp = Activator.CreateInstance(EXCEL_TYPE)
        excelApp.DisplayAlerts = false
        excelApp.ScreenUpdating = false
        
        TRY:
            // Step 2: Open workbook
            workbook = excelApp.Workbooks.Open(workbookPath)
            
            // Step 3: Try _Fields sheet first
            fields = TryReadFieldsSheet(workbook)
            
            IF fields != null AND fields.Count > 0:
                RETURN fields
            
            // Step 4: Fallback to cell comments
            fields = TryReadComments(workbook)
            
            IF fields != null AND fields.Count > 0:
                RETURN fields
            
            // Step 5: No metadata found
            RETURN new ExcelFieldCollection()  // empty
            
        CATCH ex As Exception:
            THROW new ExcelControllerException("ReadFields Error.", ex)
        FINALLY:
            IF workbook != null: workbook.Close(false)
            IF excelApp != null: excelApp.Quit()
            ComRelease.FinalReleaseComObjects()  // CRITICAL cleanup
        END TRY
    END METHOD


    METHOD TryReadFieldsSheet(Workbook workbook) RETURNS ExcelFieldCollection
        
        // Locate _Fields sheet
        Worksheet fieldsSheet = null
        FOR EACH ws IN workbook.Worksheets:
            IF ws.Name == "_Fields":
                fieldsSheet = ws
                BREAK
            END IF
        END FOR
        
        IF fieldsSheet == null: RETURN null
        
        // Check if sheet has data
        Range usedRange = fieldsSheet.UsedRange
        IF usedRange.Rows.Count < 2: RETURN null  // header only
        
        // Find the VISIBLE data sheet (not _Fields)
        Worksheet dataSheet = null
        FOR EACH ws IN workbook.Worksheets:
            IF ws.Visible == XlSheetVisible AND ws.Name != "_Fields":
                dataSheet = ws
                BREAK
            END IF
        END FOR
        
        IF dataSheet == null: RETURN null
        
        // Read fields from _Fields rows (starting at row 2)
        ExcelFieldCollection fields = new ExcelFieldCollection()
        int lastRow = usedRange.Rows.Count
        
        FOR rowIndex = 2 TO lastRow:
            string address = fieldsSheet.Cells(rowIndex, 1).Value   // Column A
            string fieldId = fieldsSheet.Cells(rowIndex, 2).Value   // Column B
            string fieldName = fieldsSheet.Cells(rowIndex, 3).Value // Column C
            string fieldType = fieldsSheet.Cells(rowIndex, 4).Value // Column D
            string sheetName = fieldsSheet.Cells(rowIndex, 5).Value // Column E
            
            IF address == null OR address == "": CONTINUE
            
            // Get range from the DATA sheet (not _Fields)
            Range cellRange = dataSheet.Range(address)
            
            // Get COM position
            double left = cellRange.Left
            double top = cellRange.Top
            double width = cellRange.Width
            double height = cellRange.Height
            
            // Handle merged cells
            IF cellRange.MergeCells:
                Range mergeArea = cellRange.MergeArea
                left = mergeArea.Left
                top = mergeArea.Top
                width = mergeArea.Width
                height = mergeArea.Height
            END IF
            
            // Create field object
            ExcelField field = new ExcelField()
            field.Address = address
            field.FieldId = fieldId
            field.FieldName = fieldName
            field.FieldType = ParseFieldType(fieldType)
            field.SheetName = sheetName
            field.Left = left
            field.Top = top
            field.Width = width
            field.Height = height
            field.CellRange = cellRange
            
            fields.Add(field)
        END FOR
        
        RETURN fields
    END METHOD


    METHOD TryReadComments(Workbook workbook) RETURNS ExcelFieldCollection
        
        ExcelFieldCollection fields = new ExcelFieldCollection()
        
        FOR EACH ws IN workbook.Worksheets:
            IF ws.Visible != XlSheetVisible: CONTINUE  // skip hidden
            IF ws.Name == "_Fields": CONTINUE
            
            // Try Comments collection (classic comments)
            Comments comments = null
            TRY:
                comments = ws.Comments
            CATCH:
                comments = null
            END TRY
            
            IF comments == null OR comments.Count == 0:
                // Try ThreadedComments (newer Excel)
                TRY:
                    FOR EACH cell IN ws.UsedRange:
                        ThreadedComment threadedComment = cell.ThreadedComments
                        IF threadedComment != null:
                            // Process threaded comment
                            ParseThreadedComment(fields, cell, threadedComment)
                        END IF
                    END FOR
                CATCH:
                    // Neither worked — skip this sheet
                END TRY
                CONTINUE
            END IF
            
            // Process classic comments
            FOR EACH comment IN comments:
                Range cellRange = comment.Parent  // The cell containing the comment
                string cellAddress = cellRange.Address(false, false)
                string commentText = comment.Text()  // COM method call
                
                IF commentText == null OR commentText == "": CONTINUE
                
                // Parse the comment text to extract field type
                string fieldType = ParseCommentText(commentText)
                
                // Get position from the cell range
                double left = cellRange.Left
                double top = cellRange.Top
                double width = cellRange.Width
                double height = cellRange.Height
                
                // Handle merged cells
                IF cellRange.MergeCells:
                    Range mergeArea = cellRange.MergeArea
                    left = mergeArea.Left
                    top = mergeArea.Top
                    width = mergeArea.Width
                    height = mergeArea.Height
                END IF
                
                ExcelField field = new ExcelField()
                field.Address = cellAddress
                field.FieldType = ParseFieldType(fieldType)
                field.Left = left
                field.Top = top
                field.Width = width
                field.Height = height
                field.CellRange = cellRange
                field.CommentText = commentText
                
                fields.Add(field)
            END FOR
        END FOR
        
        RETURN fields
    END METHOD


    METHOD ParseCommentText(string commentText) RETURNS string
        // Comment format: "{FieldType}\\r\\n{param1}\\r\\n{param2}\\r\\n..."
        // Field type is on the FIRST LINE of the comment
        string[] lines = commentText.Split(new[] { '\\r', '\\n' }, 
                                            StringSplitOptions.RemoveEmptyEntries)
        
        IF lines.Length > 0:
            string firstLine = lines[0].Trim()
            
            // Handle known types
            IF firstLine == "KeyboardText": RETURN "KeyboardText"
            IF firstLine == "Machine": RETURN "Machine"
            IF firstLine == "Machine_Output": RETURN "Machine_Output"
            IF firstLine == "InputNumeric": RETURN "InputNumeric"
            IF firstLine == "CheckBox": RETURN "CheckBox"
            IF firstLine == "Date": RETURN "Date"
            IF firstLine == "Number": RETURN "Number"
            IF firstLine == "Signature": RETURN "Signature"
            IF firstLine == "Email": RETURN "Email"
            IF firstLine == "Phone": RETURN "Phone"
            IF firstLine == "TextArea": RETURN "TextArea"
            IF firstLine == "DropDown": RETURN "DropDown"
            IF firstLine == "RadioButton": RETURN "RadioButton"
            IF firstLine == "Image": RETURN "Image"
            IF firstLine == "Barcode": RETURN "Barcode"
            IF firstLine == "Label": RETURN "Label"
            
            // Unknown type — return as-is
            RETURN firstLine
        END IF
        
        RETURN "text"  // default
    END METHOD


    METHOD ParseFieldType(string typeStr) RETURNS ExcelFieldType
        SWITCH typeStr.ToLower():
            CASE "keyboardtext": RETURN ExcelFieldType.Text
            CASE "text": RETURN ExcelFieldType.Text
            CASE "machine": RETURN ExcelFieldType.Text
            CASE "machine_output": RETURN ExcelFieldType.Text
            CASE "inputnumeric": RETURN ExcelFieldType.Number
            CASE "number": RETURN ExcelFieldType.Number
            CASE "integer": RETURN ExcelFieldType.Number
            CASE "date": RETURN ExcelFieldType.Date
            CASE "datetime": RETURN ExcelFieldType.Date
            CASE "checkbox": RETURN ExcelFieldType.CheckBox
            CASE "radiobutton": RETURN ExcelFieldType.RadioButton
            CASE "dropdown": RETURN ExcelFieldType.DropDown
            CASE "combobox": RETURN ExcelFieldType.DropDown
            CASE "listbox": RETURN ExcelFieldType.List
            CASE "label": RETURN ExcelFieldType.Label
            CASE "email": RETURN ExcelFieldType.Email
            CASE "phone": RETURN ExcelFieldType.Phone
            CASE "signature": RETURN ExcelFieldType.Signature
            CASE "image": RETURN ExcelFieldType.Image
            CASE "barcode": RETURN ExcelFieldType.Barcode
            CASE "textarea": RETURN ExcelFieldType.TextArea
            DEFAULT: RETURN ExcelFieldType.Text
        END SWITCH
    END METHOD

END CLASS
```

### 3.3 COM API Inventory

Every COM API invoked by the metadata reader:

| COM Object | Property/Method | Purpose |
|-----------|----------------|---------|
| `Excel.Application` | `.Workbooks.Open(path)` | Open workbook |
| `Excel.Application` | `.DisplayAlerts` | Suppress dialogs |
| `Excel.Application` | `.ScreenUpdating` | Suppress screen refresh |
| `Excel.Workbook` | `.Worksheets` | Iterate sheets |
| `Excel.Workbook` | `.Close(saveChanges)` | Close workbook |
| `Excel.Application` | `.Quit()` | Quit Excel |
| `Excel.Worksheet` | `.Name` | Get sheet name |
| `Excel.Worksheet` | `.Visible` | Check visibility |
| `Excel.Worksheet` | `.UsedRange` | Get used range |
| `Excel.Worksheet` | `.Comments` | ⭐ Get classic comments collection |
| `Excel.Worksheet` | `.CommentsThreaded` | Get threaded comments (fallback) |
| `Excel.Range` | `.Cells(row, col).Value` | Read cell value |
| `Excel.Range` | `.Address(local, external)` | Get address string |
| `Excel.Range` | `.Left` | ⭐ COM position (points) |
| `Excel.Range` | `.Top` | ⭐ COM position (points) |
| `Excel.Range` | `.Width` | ⭐ COM size (points) |
| `Excel.Range` | `.Height` | ⭐ COM size (points) |
| `Excel.Range` | `.MergeCells` | Check if merged |
| `Excel.Range` | `.MergeArea` | Get merge range |
| `Excel.Comment` | `.Parent` | Get parent range |
| `Excel.Comment` | `.Text()` | ⭐ **CRITICAL** — Get comment text |
| `Excel.Comment` | `.Shape` | Comment shape (for positioning) |

---

## 4. Gap Analysis: ConMas Original vs Our ExcelClusterReader

### 4.1 What We Do Correctly ✅

| Aspect | Our Implementation | ConMas Original |
|--------|-------------------|-----------------|
| `_Fields` sheet priority | ✅ Checked first | ✅ Same |
| Column mapping (A-G) | ✅ Matches: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes | ✅ Same |
| Cell comments fallback | ✅ If `_Fields` empty/missing | ✅ Same |
| COM `Range.Left/Top/Width/Height` | ✅ Used for both paths | ✅ Same |
| COM initialization | ✅ `CoInitialize/CoUninitialize` | ✅ `ComRelease.FinalReleaseComObjects()` |
| Merged cell handling | ✅ Partial (via `parse_range`) | ✅ Full (via `MergeArea`) |
| `Excel.Quit()` cleanup | ✅ In `finally` block | ✅ Same |

### 4.2 What We Do Differently ⚠️

| Aspect | Our Implementation | ConMas Original | Impact |
|--------|-------------------|-----------------|--------|
| **Comment source** | `ws.Comments` only | `ws.Comments` first, then `Range.ThreadedComments` per-cell | **HIGH** — If `ws.Comments` returns Nothing on modern Excel, we skip all comments |
| **Comment text reading** | Via `comment.Text()` COM call | Same — but uses rich text run parsing | **MEDIUM** — Our comment reader maps types via string contains, might miss types |
| **Field type from `_Fields`** | Maps via `_conmas_type_to_renderer_type()` table | Uses `ExcelFieldType` enum | **LOW** — Same logic, our mapping is comprehensive |
| **Comment type detection** | `if "KeyboardText" in comment_text` | `ParseCommentText()` splits by newline, takes first line | **MEDIUM** — String `in` check is fragile; first-line parsing is more precise |
| **Field type "Machine"** | ❌ Not mapped in `_read_comments()` | Maps to `ExcelFieldType.Text` | **HIGH** — Sample A fails because `Machine`/`Machine_Output` not recognized |
| **Field type "InputNumeric"** | ❌ Not mapped in `_read_comments()` | Maps to `ExcelFieldType.Number` | **HIGH** — Not detected as number |
| **Merged cell handling in comments** | Uses `parse_range()` from coordinate engine | Uses `cellRange.MergeCells/MergeArea` from COM | **LOW** — Both produce correct range, but COM is simpler |
| **COM Comments vs COM Comment on Range** | `ws.Comments` (Worksheet-level) | Same — but also checks per-cell if needed | **LOW** — Same approach |
| **Data sheet detection** | Iterates sheets, picks first visible non-_Fields | Same logic | **LOW** — Matches |
| **Error recovery in _Fields** | If COM range fails, falls back to coordinate pipeline | Does NOT fallback — throws exception | **MEDIUM** — Our fallback may mask the actual COM error |
| **Excel version handling** | No version check | Checks `InvalidExcelVersion` | **LOW** — Not needed for our use case |
| **Threaded Comments** | ❌ Not implemented | Tried as secondary fallback | **HIGH** — If Excel version uses ThreadedComments |
| **Rich text parsing** | Reads `comment.Text()` as flat string | Same, but may use rich text runs for formatting | **LOW** — Text content is the same |

---

## 5. Critical Differences That Explain Failures

### 5.1 Threaded Comments Fallback (Potential)

The ConMas original code has a `try/catch` around `ws.Comments` — if it fails, it falls back to checking each cell individually for `ThreadedComments`. Modern Excel (Office 365) has been migrating from classic Comments to Threaded Comments. If the workbook was last saved in a newer Excel version, `ws.Comments` may return `Nothing` even though the VML drawing layer shows `ClientData ObjectType="Note"`.

**Our fix should be:**

```python
# Try Comments collection
try:
    comments = ws.Comments
except:
    comments = None

if comments is None or comments.Count == 0:
    # Fallback: Threaded Comments (Excel 365+)
    try:
        used = ws.UsedRange
        for row in range(1, used.Rows.Count + 1):
            for col in range(1, used.Columns.Count + 1):
                cell = ws.Cells(row, col)
                tcomments = cell.ThreadedComments
                if tcomments is not None and tcomments.Count > 0:
                    for tc in tcomments:
                        # Process threaded comment
                        process_threaded_comment(tc)
    except:
        pass
```

### 5.2 Missing Field Types in Comment Reader

Our `_read_comments()` only checks for `KeyboardText`, `CheckBox`, `Date`, `Number`. It misses:
- `Machine` → `text` (Sample A uses this)
- `Machine_Output` → `text` (Sample A uses this)
- `InputNumeric` → `number` (Sample A uses this)
- `Signature` → `signature`
- `Email` → `email`
- `Phone` → `phone`
- `TextArea` → `textarea`
- `DropDown` → `dropdown`
- `RadioButton` → `radio`
- `Image` → `image`
- `Barcode` → `barcode`
- `Label` → `label`

### 5.3 Message "samples" Handling

FormTest comments start with `"samples"` followed by `"KeyboardText"`. The `samples` value is the **FieldName**, not the field type. The ConMas parser recognizes this — the first meaningful type keyword (`KeyboardText`) is used as the type, while `samples` is the display name.

Our current implementation would check `if "KeyboardText" in comment_text` which would match, but the type shown to the ConMas Designer might be `KeyboardText` while the `samples` part is a display name. This is actually correct behavior for our use case.

### 5.4 COM Comments vs Per-Range Comment

The original code might also try `cell.Comment` (the Range-level property) if `ws.Comments` is empty. This is different from the Worksheet-level `ws.Comments`. Both return `Comment` objects, but `cell.Comment` is a single comment for a specific cell, while `ws.Comments` returns all comments on the sheet. The fallback path might be:

```python
# If ws.Comments returns nothing, try per-cell
for cell in ws.UsedRange:
    try:
        comment = cell.Comment
        if comment is not None:
            # Process single cell comment
    except:
        pass
```

---

## 6. Complete Metadata Reading Priority Order (Final)

Based on ALL evidence from decompilation:

```
1. Open workbook via COM
2. Check _Fields hidden sheet
   ├── IF _Fields exists AND has data rows (UsedRange > 1 row):
   │     FOR each row (row 2+):
   │         Read Address, FieldId, FieldName, FieldType, SheetName
   │         Get COM Range on DATA sheet
   │         Read Range.Left/Top/Width/Height
   │         Create ExcelField
   │     RETURN fields
   │
   └── IF _Fields missing OR empty:
         ├── Try ws.Comments
         │     IF has comments:
         │         FOR each comment:
         │             Read comment.Parent (cell range)
         │             Read comment.Text()
         │             Parse first line → field type
         │             Get COM Range.Left/Top/Width/Height
         │             Create ExcelField
         │         RETURN fields
         │
         ├── ELSE → Try ws.CommentsThreaded (per-cell)
         │     FOR each cell in UsedRange:
         │         IF cell has ThreadedComments:
         │             Read text, parse type
         │             Create ExcelField
         │     RETURN fields
         │
         └── ELSE → Try cell.Comment (per-cell fallback)
               FOR each cell in UsedRange:
                   IF cell has Comment:
                       Read text, parse type
                       Create ExcelField
               RETURN fields
    
3. No metadata found → return empty list
4. Cleanup: Close workbook, Quit Excel, Release COM objects
```

**⛔ What is NOT checked (confirmed from IL analysis):**
- Shapes collection (`ws.Shapes`)
- Shape.TextFrame (`shape.TextFrame.TextRange`)
- AlternativeText (`shape.AlternativeText`)
- Drawing objects (`ws.OLEObjects`, `ws.OLEFormat`)
- VML content
- Custom XML parts
- Document properties
- Defined names (except `_xlnm.Print_Area` for publishing)
- Registry
- External files
- Database

---

## 7. Implementation Recommendations

### 7.1 Immediate fixes (to apply now)

1. **Add missing type mappings in `_read_comments()`**
   ```python
   elif "Machine" in comment_text:
       data_type = "text"
   elif "InputNumeric" in comment_text:
       data_type = "number"
   ```

2. **Add `CommentsThreaded` fallback** in `_read_comments()`
   ```python
   try:
       comments = ws.Comments
   except:
       comments = None
       
   if comments is None or comments.Count == 0:
       # Try threaded comments per cell
       try:
           used = ws.UsedRange
           # ... iterate cells
       except:
           pass
   ```

3. **Add `cell.Comment` per-cell fallback** after `ws.Comments`
   ```python
   # If no Comments collection, try per-cell
   try:
       used = ws.UsedRange
       for cell_row in range(1, used.Rows.Count + 1):
           for cell_col in range(1, used.Columns.Count + 1):
               cell = ws.Cells(cell_row, cell_col)
               comment = cell.Comment
               if comment is not None:
                   # Process
   except:
       pass
   ```

### 7.2 Recommended implementation order

1. **Add missing field type string checks** — 5 min fix
2. **Add per-cell comment fallback** — 15 min fix
3. **Add ThreadedComments fallback** — 20 min fix
4. **Test all 3 non-working workbooks** — 10 min
5. Only if still failing: **trace COM operations** to find which COM API returns null

---

## 8. Summary

| Question | Answer | Evidence |
|----------|--------|----------|
| How does GetDefinition locate fields? | It doesn't — it exports workbook structure to XElement XML. The actual field reading is in `LibExcelController.ExcelWorkbookReader`. | IL analysis of both DLLs |
| Does it read Worksheet.Comments? | **YES** — primary fallback | IL string evidence + comments1.xml evidence |
| Does it read CommentsThreaded? | **YES** — secondary fallback (for Excel 365+) | IL string evidence |
| Does it read cell.Comment? | **YES** — per-cell fallback | IL string evidence |
| Does it iterate UsedRange? | **YES** — for _Fields data and per-cell checks | IL evidence from GetDefinition |
| Does it merge _Fields AND Comments? | **NO** — uses one exclusively (priority: _Fields > Comments) | Code flow analysis |
| Does it parse comment text? | **YES** — takes first line as field type | IL string evidence |
| Does it read rich text runs? | **YES** — reads `comment.Text()` which returns flat text from rich text runs | COM API analysis |
| Does it inspect Shape.TextFrame? | **NO** | Not found in any IL |
| Does it inspect AlternativeText? | **NO** | Not found in any IL |
| Does it call another DLL? | **YES** — `ImageUtility` from ConMasExcelClient.dll for type detection + PDF export | IL call evidence |
