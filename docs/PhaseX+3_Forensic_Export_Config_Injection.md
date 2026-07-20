# Phase X+3 — Forensic Report: How Legacy ConMas Injects Designer Configuration into the Output Excel

## Investigation Goal

**Primary question:** Where exactly does menuOutputExcel_Click write designer configuration (cluster metadata, cell values, formulas, formatting) into the workbook before SaveAs/SaveDocument?

**Answer:** The designer configuration is injected into the workbook via **three distinct mechanisms**, all inside the async state machine of menuOutputExcel_Click (ConMasClient.exe, IL lines 183774–188959):

| Mechanism | What is Written | Where in IL |
|-----------|----------------|-------------|
| **Cell Comments** (PRIMARY metadata store) | 20-field newline-delimited cluster metadata string | Lines 185565–186317 |
| **Cell Formulas** | Original function formula or ChangeCellCalculate() result | Lines 186340–186425 |
| **Cell Formatting** | Font size, name, bold, italic, underline, strikethrough, color, background, alignment | Lines 186452–186638 |

No _Fields sheet, no ExcelOutputSetting object, no hidden sheets, no custom XML parts, no document properties are created.

---

## 1. The Cluster Iteration Structure

`
foreach (int sheetIdx in reportIndexList)               // V_52 loop over V_41
{
    BlankReport sheet = topReport.Sheets[V_52];
    foreach (BlankCluster cluster in sheet.Clusters)     // V_53 = current cluster
    {
        // Skip if cluster.CellAddress is null or empty
        // ... process cluster ...
    }
}
`

- IL entry at line 185442 (IL_0e8e), loop back at IL_222d
- Cluster enumerator at IL lines 185452–185465

---

## 2. Comment Text Construction (V_55)

### Phase 1 — Base Cluster Metadata (IL lines 185565–185715)

`
V_55 = cluster.Name                                    // line 185569
V_55 += "\n" + ClusterTypeToStringType(cluster.Type)   // line 185579
V_55 += "\n" + cluster.Index.ToString()                // line 185590
V_55 += "\n" + cluster.ReadOnly.ToString()             // line 185599
V_55 += "\n" + cluster.External.ToString()             // line 185608
V_55 += "\n" + V_54                                    // line 185615 (InputParameter)
V_55 += "\n" + cluster.RemarksValue[0]                 // line 185625
V_55 += "\n" + cluster.RemarksValue[1]                 // line 185635
V_55 += "\n" + cluster.RemarksValue[2]                 // line 185645
V_55 += "\n" + cluster.RemarksValue[3]                 // line 185655
V_55 += "\n" + cluster.RemarksValue[4]                 // line 185665
V_55 += "\n" + cluster.RemarksValue[5]                 // line 185675
V_55 += "\n" + cluster.RemarksValue[6]                 // line 185685
V_55 += "\n" + cluster.RemarksValue[7]                 // line 185695
V_55 += "\n" + cluster.RemarksValue[8]                 // line 185705
V_55 += "\n" + cluster.RemarksValue[9]                 // line 185715
`

Where **V_54 = cluster.InputParameter**, possibly with FunctionVersion=4.2.0000 appended if the input parameters contain a "Function" key but no "FunctionVersion" key (determined by parsing InputParameter via ParseParameterText()).

### Phase 2 — Table/Column/Row Context (IL lines 185716–186104)

Before Phase 2, the code searches 	opReport.Tables for a matching TableCluster:

`
foreach (TableData table in topReport.Tables)
    foreach (TableRow row in table.Rows)
        foreach (TableCluster tc in row.Clusters)
            if (tc.SheetNo == cluster.SheetNo && tc.ClusterId == cluster.Index)
            {
                V_56 = table.No.ToString()
                V_57 = table.Name
                V_58 = table.CooperationTable.ToString()

                // Inside: search table.Columns for column matching tc.ColumnKey
                // Skip columns where Column.Key == "NONE" (up to V_77 times)
                V_59 = column.No.ToString()
                V_60 = column.Name
                V_61 = column.Key
                V_62 = column.Type

                V_63 = row.No.ToString()
                V_64 = row.Name
                V_74 = true  // found
            }
`

Then the table info is appended:

`
V_55 += "\n" + V_56  // Table No
V_55 += "\n" + V_57  // Table Name
V_55 += "\n" + V_58  // CooperationTable
V_55 += "\n" + V_59  // Column No
V_55 += "\n" + V_60  // Column Name
V_55 += "\n" + V_61  // Column Key
V_55 += "\n" + V_62  // Column Type
V_55 += "\n" + V_63  // Row No
V_55 += "\n" + V_64  // Row Name
`

### Complete Comment Format (20 fields, delimited by \n)

`
line 0:  ClusterName
line 1:  ClusterTypeString
line 2:  ClusterIndex
line 3:  ReadOnly
line 4:  External
line 5:  InputParameter
line 6:  RemarksValue[0]
line 7:  RemarksValue[1]
line 8:  RemarksValue[2]
line 9:  RemarksValue[3]
line 10: RemarksValue[4]
line 11: RemarksValue[5]
line 12: RemarksValue[6]
line 13: RemarksValue[7]
line 14: RemarksValue[8]
line 15: RemarksValue[9]
line 16: TableNo
line 17: TableName
line 18: CooperationTable
line 19: ColumnNo
line 20: ColumnName
line 21: ColumnKey
line 22: ColumnType
line 23: RowNo
line 24: RowName
`

**Total: 25 lines = 25 newline-separated fields** (counting from 0).

---

## 3. Cell Address Resolution

The target cell address is resolved from BlankCluster.CellAddress:

`
V_65 = cluster.CellAddress.Split(':')[0]    // Take first part before ':'
V_65 = V_65.Trim('$', ' ')                  // Remove $ signs
`

If V_65 is in R1C1 format (contains R and C but not RC or CR):
- Convert to A1 format via AnalyzeExcelFunction.ChangeRCtoAlphabet(V_65)
- If conversion returns empty: show error (ExcelMessage25) and exit

The result V_65 is a standard Excel cell reference like "A1" or "".

---

## 4. Writing the Comment — Two Code Paths

Based on ExcelProcessConfig.ExportMethod:

### DevExpress Path (ExportMethod == 1)
`csharp
string author = workbook.CurrentAuthor;
Worksheet sheet = workbook.Worksheets[sheetIndex - 1];  // V_39 - 1
CellRange cell = sheet.Cells[V_65];                      // e.g., "A1"
RangeExtensions.ClearComments(cell);                      // Clear existing comment
sheet.Comments.Add(cell, author, V_55);                  // Add new comment
`

### COM Interop Path (ExportMethod != 1) — DEFAULT
`csharp
dynamic sheet = workbook.Sheets[V_39];
dynamic range = sheet.Range[V_65];
range.AddComment(V_55);                                  // Add comment
range.Comment.Text(V_55, Type.Missing, Type.Missing);    // Set text
`

IL lines: DevExpress at 186189–186218, COM at 186220–186317.

---

## 5. Post-Comment Processing

### 5a. Formula Writing (only for ClusterType 67 or 55)

If cluster.InputParameter contains "FunctionVersion=4.3.0000" AND cluster.OriginalFunction is not null/empty:
- **Set cell.Formula = cluster.OriginalFunction** (wrapped in try-catch)

Otherwise:
- Parse InputParameter as dictionary
- Look for "Function" key with non-empty value
- Call ChangeCellCalculate(functionValue, cluster.SheetNo, cluster.ClusterType)
- If result does NOT contain "#REF":
  - **Set cell.Formula = result** (wrapped in try-catch)

IL lines: 186340–186425.

### 5b. Font Formatting (if ExcelFontCheck is true)

Skips cluster types: 20, 10, 15, 90, 100, 116, 117, 118.

For type 126 with "ButtonFontVerticalAlignment=2" or any type with "VerticalAlignment=2":
- Set cell vertical alignment = Center

If cluster.EditParameter is true — parse InputParameter dictionary and apply:

| Key | Formatting Applied |
|-----|-------------------|
| ButtonFontSize | ont.Size = Double.Parse(value) |
| ButtonFontAlign = "Left" | lignment.Horizontal = Left |
| ButtonFontAlign = "Right" | lignment.Horizontal = Right |
| ButtonFontAlign = "Center" | lignment.Horizontal = Center |
| ButtonFontColor | ont.Color = Color.FromArgb(...) |
| ButtonFontBackground | interior.Color = Color.FromArgb(...) |
| ButtonFontBold = "True" | ont.Bold = true |
| ButtonFontItalic = "True" | ont.Italic = true |
| ButtonFontUnderline | ont.Underline = ... |
| ButtonFontName | ont.Name = value |
| ButtonFontStrikout = "True" | ont.Strikout = true |

All formatting is applied via DevExpress API or COM API depending on ExportMethod.

---

## 6. Save Flow

### DevExpress Path (ExportMethod == 1)
`csharp
workbook.SaveDocument(tempPath);       // IL 188653
workbook.Dispose();
ExcelProcessorBase.EnsurePrintAreaQuoted_OpenXml(tempPath);  // Post-processing fix
`

### COM Interop Path (ExportMethod != 1) — DEFAULT
`csharp
// Activate first sheet, select, sleep 300ms
workbook.CheckCompatibility = false;
workbook.Saved = true;
app.DisplayAlerts = false;
app.AlertBeforeOverwriting = false;
// Set TopMost window for Excel dialog
workbook.SaveAs(tempPath, ..., XlSaveAsAccessMode.xlNoChange, ...);
// Restore display settings
`

### Post-save File Copy (COM path only, when filename contains "[")
`csharp
File.Copy(tempPath, saveDialog.FileName, overwrite: true);
File.Delete(tempPath);
`

---

## 7. Cleanup

### DevExpress
`csharp
workbook.Dispose();
workbook = null;
`

### COM Interop
`csharp
workbook.Close(Type.Missing, Type.Missing, Type.Missing);
workbooks.Close();
app.Quit();
Marshal.ReleaseComObject(workbook);
Marshal.ReleaseComObject(workbooks);
Marshal.ReleaseComObject(app);
workbook = null;
workbooks = null;
app = null;
`

---

## 8. Key Findings

| # | Finding | Evidence |
|---|---------|----------|
| 1 | Cell comments are the PRIMARY mechanism for persisting cluster metadata in the output Excel | Every cluster iteration writes a comment (V_55) via AddComment or Comments.Add |
| 2 | The comment format is **25 newline-separated fields** spanning cluster identity, parameters, remarks, and table/column/row provenance | Exact string concatenation traced in IL (lines 185565-186104) |
| 3 | Cell formulas are written for function clusters (type 67/55) using OriginalFunction or ChangeCellCalculate() | IL 186340-186425 with try-catch |
| 4 | Cell formatting (font, alignment, colors) is applied from InputParameter dictionary keys (ButtonFontSize, etc.) | IL 186452-186638 with extensive InputParameter parsing |
| 5 | The InputParameter may be auto-patched with FunctionVersion=4.2.0000 if missing | IL 185544-185564 |
| 6 | No _Fields sheet, ExcelOutputSetting, hidden sheets, custom XML parts, or document properties are created | Grep search confirms zero references in export block |
| 7 | The workbook is opened from a COPY of the original file (not Workbooks.Add) | Lines 185104-185106: LoadDocument(path), lines 185135-185144: Workbooks.Open(path) |
| 8 | R1C1 cell addresses are converted to A1 format before comment writing | IL 186129-186178: ChangeRCtoAlphabet() |
| 9 | Existing cell comments are cleared before writing (DevExpress path) | IL 186203: ClearComments(cellRange) |
| 10 | The COM Interop path writes comment text twice: via AddComment(text) and .Comment.Text(text, ...) | IL 186307 and IL 186314 |

## 9. Conclusion

**The exported Output Excel acquires designer configuration entirely through cell comments, cell formulas, and cell formatting — all applied inline in menuOutputExcel_Click without delegating to any GetDefinition() or ExportExcelDefinition() helper.**

The cell comment (V_55) carries the FULL cluster metadata, making the workbook self-describing. The cell formula carries the FUNCTION logic for executable clusters. The cell formatting carries the VISUAL configuration. Together, these three mechanisms make the output Excel fully self-reopenable.

During reopen, the rendering pipeline reads:
- **Cell comments** → RegisterCommentCells() maps cell addresses to cluster IDs (for rendering)
- **XML form definition** (from if files) → InputFromIfFiles creates BlankCluster objects from XML elements (for cluster definitions)
- **Cell formulas** → evaluated by Excel at runtime (for function behavior)
- **Cell formatting** → visible in Excel (for visual appearance)
