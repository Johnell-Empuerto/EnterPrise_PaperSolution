# Phase X.43 — Determine How the Designer Produces Server Rendering Assets

**Date**: 2026-07-19
**Objective**: Reconstruct the complete pipeline by which the Designer WPF client processes an Excel template and produces the assets consumed by the server renderer (background PDF + cluster coordinates + images).

---

## 1. High-Level Pipeline

```
Excel Template (.xlsx)
    │
    ▼
LibExcelController.dll ─────────────────────────────────────┐
    │                                                       │
    ├─ ExportPdf() → ExportPdf2010()                        │
    │   → COM: Workbook.ExportAsFixedFormat(PDF)            │
    │   → produces background PDF per sheet                 │
    │                                                       │
    ├─ CreateImageFromPdf()                                 │
    │   → renders PDF page as internal Image (pixel buffer) │
    │                                                       │
    ├─ GetClusterSize() → GetAddress()                      │
    │   → pixel scan to detect black-bordered rectangles    │
    │   → returns List<ClusterRect> (Top/Bottom/Left/Right) │
    │                                                       │
    └─ CalcClusterSize()                                    │
        → serializes ClusterRect[] into XML                  │
        → <cluster><top/><bottom/><left/><right/></cluster> │
                                                             │
                                                             ▼
LibConMas.dll ──────────────────────────────────────────────┐
    │ ConMasWebClient.UploadDefinition(filePath)             │
    │   → HTTP POST via WebClient.UploadFile()               │
    │   → URL: {ServerUrl}\?command=RegistDefinition         │
    │   → response: XML <error><code>VALUE</code></error>    │
    │                                                         │
    └─ Cluster images uploaded separately                     │
        → RegistClusterImage.aspx handler (on server)        │
                                                              │
                                                              ▼
Server (ConMasAPI.dll / Cimtops.ConMasBiz.dll) ─────────────┐
    │ DefinitionBiz.Regist()  → def_top table                 │
    │ ClusterImageBiz.Regist() → cluster image storage        │
    │                                                         │
    └─ Later: ClusterWriter subclasses                       │
        → draw on PdfGraphics from stored metadata            │
        → produce final rendered PDF                         │
```

---

## 2. PDF Creation from Excel

**File**: `LibExcelController.il`
**Class**: `LibExcelController.Lib.ImageUtility`

### Entry Point

```
ExportPdf(string sourceXlsPath, string destPdfPath)
```

- Logs start/end
- Calls `GetExcelVersion()` to determine Excel version
- Routes to appropriate implementation:
  - **Version `"12.0"`** (Excel 2007) → `ExportPdf2007()`
  - **Version `"14.0"`** (Excel 2010) → `ExportPdf2010()`
  - Any other → throws `ExcelControllerException("InvalidExcelVersion")`

### ExportPdf2010 (Lines 7433–7604)

Uses **Microsoft.Office.Interop.Excel** (not dynamic/C# `dynamic` like other parts):

```
Workbooks workbooks = application.Workbooks;
Workbook workbook = workbooks.Open(sourceXlsPath, ...);
workbook.ExportAsFixedFormat(
    XlFixedFormatType.xlTypePDF,   // = 0
    destPdfPath,                   // filename
    Quality: 0,                    // XlFixedFormatQuality.xlQualityStandard
    IncludeDocProperties: true,
    IgnorePrintAreas: false,
    OpenAfterPublish: false
);
```

Cleanup:
- `Thread.Sleep(3000)` before close (race condition mitigation)
- `workbook.Close(false, Missing, Missing)`
- `Marshal.FinalReleaseComObject(workbook)`
- `Marshal.FinalReleaseComObject(workbooks)`

### ExportPdf2007 (Lines 7606–7645)

Simply delegates to `ExportPdf2010()`. If **any** exception is thrown:
- Logs error
- Throws `ExcelControllerException("NoPDFAddin", exceptionMessage)`

This tells us Excel 2007 requires the "Microsoft Save as PDF" add-in; Excel 2010+ has it built-in.

### CanExportPdf

- Checks if Excel application object is non-null
- Checks if `GetExcelVersion()` returns `>= "12.0"` (Excel 2007+)

---

## 3. Coordinate Extraction (Cluster Detection)

**File**: `LibExcelController.il`
**Class**: `LibExcelController.Lib.ImageUtility`

### GetClusterSize (static, Lines 7647–7699)

```
static List<ClusterRect> GetClusterSize(string pdfPath, int page, int clusterCount)
```

1. Creates a rendered `Image` from the PDF page via `ExcelControllerBase.CreateImageFromPdf(pdfPath, page, useCache: true)`
2. Calls `GetAddress(image, clusterCount)` to detect clusters
3. Returns `List<ClusterRect>`

### GetAddress (private static, Lines 7701–8455+)

This is the **core coordinate extraction algorithm**. It performs a **pixel-level scan** of the rendered PDF image to locate black-bordered rectangles (the cells/ranges that users draw in Excel templates).

Algorithm:
1. **Initialize**: Create `List<ClusterRect>(clusterCount)`, loop `y = 1..Height`, `x = 1..Width`
2. **Find cluster start** (top-left corner of a black border):
   - Current pixel is black (`GetPixelColor(x, y) == Color.Black`)
   - Pixel to left (`x-1, y`) is NOT black → this is a left border start
   - Pixel above (`x, y-1`) is NOT black → this is a top border start
3. **Validate border thickness**: The next 6 pixels to the right (`x+1..x+6`) must all be black at `y-1` (above), confirming a horizontal border line. If any non-black pixel found in this check, the candidate is rejected.
4. **Extend horizontally**: Scan right from `(x, y)` while pixels are black, to find the right edge
5. **Extend vertically**: Scan down from `(x, y)` while pixels are black, to find the bottom edge
6. **Construct ClusterRect**: `new ClusterRect { Left=x, Top=y, Right=rightX, Bottom=bottomY }`
7. **Deduplicate**: Compare with existing rectangles; skip if the `Left` matches any existing rect (within tolerance?). The `CompareTo` method sorts by `Left`.
8. **Loop until** all `clusterCount` rectangles are found

This algorithm is why **Excel cells must have black borders** — the system identifies cluster boundaries by finding black rectangular outlines in the PDF rendering.

### ClusterRect Class (Lines 6053–6244)

```
class ClusterRect : IComparable, IComparable<ClusterRect>
{
    float Top { get; set; }
    float Bottom { get; set; }
    float Left { get; set; }
    float Right { get; set; }

    // Comparison by Left only
    int CompareTo(ClusterRect other) => Left.CompareTo(other.Left);
}
```

Coordinates are in **pixel units** of the rendered PDF image (depends on rendering DPI).

### Image Class (Custom, pixel-based)

The PDF page is rendered into an internal `LibExcelController.Lib.Image` class that provides:
- `int Width` / `int Height`
- `Color GetPixelColor(int x, int y)` — returns enum { White, Black, Other }
- RGB pixel buffer (likely GDI+ `Bitmap` wrapped or custom managed buffer)

---

## 4. Serialization & Upload

### CalcClusterSize (LibExcelController.il, ~Line 3440)

This method:
1. Exports each sheet to PDF
2. Calls `GetClusterSize()` to get rectangles
3. Builds an **XML document** (`xConmas` element):

```xml
<xConmas>
  ...
  <sheet index="0">
    <cluster>
      <top>123.45</top>
      <bottom>678.90</bottom>
      <left>55.0</left>
      <right>444.0</right>
    </cluster>
    <cluster>
      ...
    </cluster>
  </sheet>
</xConmas>
```

Coordinates are serialized as `System.Single` (float32) values.

### UploadDefinition (LibConMas.il, Line 306334)

```
int UploadDefinition(string filePath, string orgTopId,
                     bool withApprove, string approveComment)
```

**Class**: `LibConMas.Utility.ConMasWebClient`
**Transport**: `WebClientEx` (custom subclass of `System.Net.WebClient` with timeout + cookie support)

Payload construction:
1. URL: `{ServerUrl}\?command=RegistDefinition`
2. If `withApprove`: appends `&approveType=1&comment={approveComment}`
3. Calls `webClient.UploadFile(url, filePath)` — HTTP POST, file body as `multipart/form-data`
4. Response: XML `<result><error><code>VALUE</code></error></result>`
5. Returns `int.Parse(code)` (0 = success?)

**Note**: The `filePath` contains the complete definition XML including cluster coordinates. The actual `.xlsx` file itself is **NOT** uploaded — only the processed metadata (XML definition + background PDF + cluster images).

### Cluster Image Upload

A separate upload path exists for cluster images:
- Server-side handler: `ConMasAPI.Rests.RegistClusterImage` → `ClusterImageBiz.Regist(httpContext)`
- This handles the cropped images of individual clusters (not the XML definition)
- Upload mechanism details are in the runtime binary (`ConMasGeneratorUtility.exe` / `ConMasGeneratorLib.dll`), not in the Designer itself

---

## 5. Alternative Path: ConMasExcelConverter.exe

**File**: `ConMasExcelConverter\ConMasExcelConverter.exe`
**Technology**: WinForms, NetOffice (ExcelApi.dll)

This is a **batch converter** for bulk Excel processing. It does NOT produce PDF or rendering assets — it converts Excel workbook **data structure**:

1. Opens Excel via NetOffice COM wrappers
2. Opens `ExcelConverter.xlsm` (password-protected VBA project, password: `"cimtops"`)
3. Runs VBA macro: `Application.Run("CopyWorkbookContents", ...)`
4. The VBA macro copies worksheet structure (named ranges, formulas) from source to template
5. Falls back to `ConvertImage()` using Infragistics if VBA fails

This tool is used for **data import/conversion**, not for the Designer-to-server rendering pipeline.

---

## 6. Full Upload Payload Summary

| Asset | Producer | Upload Method | Server Handler | Destination |
|---|---|---|---|---|
| **Background PDF** | `ExportPdf2010()` via COM | Part of definition file? Or separate? | Needs investigation | `rep_top.background_image_file` |
| **Cluster XML** (coordinates) | `GetAddress()` → `CalcClusterSize()` → XML | `UploadDefinition()` via `WebClient.UploadFile()` | `RegistDefinition.aspx` → `DefinitionBiz.Regist()` | `def_top` / `def_cluster` |
| **Cluster images** (cropped) | (Likely `GetAddress()` crop + encode) | (Separate HTTP upload) | `RegistClusterImage.aspx` → `ClusterImageBiz.Regist()` | Cluster image storage |

---

## 7. Key Design Decisions

1. **Excel dependency is client-side only**: The Designer has Excel installed; the server does not. All Excel→PDF conversion happens on the Designer workstation via COM Interop.

2. **Pixel-based coordinate extraction**: Rather than using Excel's object model (Range.Left, Range.Top, etc.), the Designer renders the PDF to a bitmap and scans pixels to find black borders. This makes the cluster detection dependent on the visual appearance of the PDF, not the internal Excel structure.

3. **Redundant Excel controller DLLs**: `LibExcelController.dll` (80 KB, COM + Infragistics) and `ConMasExcelClient.dll` (78 KB, similar) both exist. The latter may be an older or alternative version. `LibExcelController` is the primary one used by the WPF Designer.

4. **Single-file upload**: The entire definition (XML + embedded assets?) is uploaded as a single file via `WebClient.UploadFile()`. The server parses the definition XML and stores it in relational tables.

---

## 8. Open Questions

1. **Are cluster images uploaded within the same file as the definition XML, or separately?** The `RegistClusterImage.aspx` handler exists on the server side, suggesting a separate upload for cluster images. The exact multipart structure of `UploadDefinition` needs further investigation.

2. **How is the background PDF included in the upload?** Does `UploadDefinition` send the `.xlsx`, the `.pdf`, or the XML definition? The parameter `filePath` could be any of these.

3. **What is the exact MIME format of the upload?** `WebClient.UploadFile()` sends `multipart/form-data`, but the field names and structure are unknown without seeing the server's parse logic.

4. **How does `CreateImageFromPdf()` work?** This is in `ExcelControllerBase` (base class in the same assembly) and renders a PDF page to an internal `Image` object. The rendering technology (GDI+, PDFium, custom PDF parser?) is not immediately visible without decompiling the base class methods.
