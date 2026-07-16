using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Services.Interfaces;

namespace ExcelAPI.Services
{
    /// <summary>
    /// Service that:
    ///   1. Opens an Excel workbook via COM
    ///   2. Exports the print area to PDF
    ///   3. Converts PDF to PNG
    ///   4. Reads comments via COM and computes field coordinates from Range geometry
    ///   5. Returns CaptureResult with field metadata
    ///
    /// All coordinate geometry is measured directly from Excel COM:
    ///   - Range.Left, Range.Top, Range.Width, Range.Height
    ///   - Comment text determines field type
    ///   - No external dependencies (no Python, no OpenXML recalculation)
    /// </summary>
    public class ExcelCaptureService : IExcelCaptureService
    {
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<ExcelCaptureService> _logger;
        private readonly CoordinateTransformer _coordTransformer;

        // 300 DPI / 72 points per inch = 4.1667 pixels per point
        private const double PointsToPixels = 300.0 / 72.0;

        public ExcelCaptureService(
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<ExcelCaptureService> logger,
            CoordinateTransformer coordTransformer)
        {
            _env = env;
            _options = options;
            _logger = logger;
            _coordTransformer = coordTransformer;
        }

        /// <inheritdoc />
        public async Task<CaptureResult> CapturePrintAreaAsync(
            string excelFilePath,
            string? fileId = null,
            CancellationToken cancellationToken = default)
        {
            if (!System.IO.File.Exists(excelFilePath))
            {
                throw new FileNotFoundException("Excel file not found.", excelFilePath);
            }

            var sw = Stopwatch.StartNew();

            Application? excelApp = null;
            Workbook? workbook = null;
            Excel.Worksheet? worksheet = null;
            string? pdfPath = null;
            string? previewPath = null;
            string localFileId = fileId ?? Guid.NewGuid().ToString("N");

            try
            {
                _logger.LogInformation("[Stage 1/8] Capture started for file: {FilePath}", excelFilePath);

                // ── Stage 1: Launch Excel COM ──────────────────────────
                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogDebug("Launching Microsoft Excel application...");
                excelApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                _logger.LogInformation("[Stage 1/8] Excel started in {Duration}ms", sw.ElapsedMilliseconds);

                string previewFolder = Path.Combine(_env.ContentRootPath, _options.Value.PreviewDirectory);
                Directory.CreateDirectory(previewFolder);

                // ── Stage 2: Open workbook ─────────────────────────────
                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogDebug("Opening workbook: {FilePath}", excelFilePath);
                workbook = excelApp.Workbooks.Open(excelFilePath);
                _logger.LogInformation("[Stage 2/8] Workbook opened in {Duration}ms", sw.ElapsedMilliseconds);

                // ── Stage 3: Select first visible worksheet ────────────
                _logger.LogDebug("Selecting first visible worksheet...");
                Excel.Worksheet? selectedSheet = null;
                for (int sheetIdx = 1; sheetIdx <= workbook.Worksheets.Count; sheetIdx++)
                {
                    var ws = (Excel.Worksheet)workbook.Worksheets[sheetIdx];
                    bool isVisible = false;
                    try
                    {
                        isVisible = ws.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                        if (isVisible)
                        {
                            selectedSheet = ws;
                            _logger.LogDebug("Selected visible worksheet: \"{Name}\" (Index: {Index})",
                                ws.Name, ws.Index);
                        }
                    }
                    finally
                    {
                        if (!isVisible)
                            Marshal.ReleaseComObject(ws);
                    }

                    if (isVisible)
                        break;
                }

                if (selectedSheet == null)
                {
                    throw new PrintAreaNotConfiguredException(
                        "No visible worksheets found in the workbook.");
                }
                worksheet = selectedSheet;
                _logger.LogInformation("[Stage 3/8] Worksheet \"{Name}\" selected in {Duration}ms",
                    worksheet.Name, sw.ElapsedMilliseconds);

                // ── Stage 4: Detect print area ─────────────────────────
                string? printArea = DetectPrintArea(workbook, worksheet);
                if (string.IsNullOrEmpty(printArea))
                {
                    throw new PrintAreaNotConfiguredException(
                        "No print area is configured in this worksheet. " +
                        "Please set a print area (Page Layout > Print Area > Set Print Area) " +
                        "in the Excel file before uploading.");
                }

                _logger.LogInformation("[Stage 4/8] Print area detected: \"{PrintArea}\" on \"{Sheet}\" in {Duration}ms",
                    printArea, worksheet.Name, sw.ElapsedMilliseconds);

                // ── Stage 5: Export PDF ────────────────────────────────
                string pdfFileName = $"page_{localFileId}.pdf";
                string previewFileName = $"page_{localFileId}.png";
                pdfPath = Path.Combine(previewFolder, pdfFileName);
                previewPath = Path.Combine(previewFolder, previewFileName);

                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogDebug("Exporting print area to PDF: {PdfPath}", pdfPath);

                worksheet.ExportAsFixedFormat(
                    Type: XlFixedFormatType.xlTypePDF,
                    Filename: pdfPath,
                    Quality: XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: false,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);

                if (!System.IO.File.Exists(pdfPath))
                {
                    throw new InvalidOperationException(
                        "Excel did not generate a PDF file. The print area may be empty or invalid.");
                }

                var pdfFileInfo = new FileInfo(pdfPath);
                if (pdfFileInfo.Length == 0)
                {
                    throw new InvalidOperationException(
                        "Excel generated an empty PDF file. The print area may be empty or invalid.");
                }

                _logger.LogInformation("[Stage 5/8] PDF exported in {Duration}ms: {PdfPath} ({Size} bytes)",
                    sw.ElapsedMilliseconds, pdfPath, pdfFileInfo.Length);

                // ── Stage 6: Convert PDF to PNG ────────────────────────
                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogDebug("Converting PDF to PNG at 300 DPI: {PngPath}", previewPath);

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath, cancellationToken);

                using var bitmap = PDFtoImage.Conversion.ToImage(
                    pdfBytes,
                    page: 0,
                    options: new PDFtoImage.RenderOptions
                    {
                        Dpi = 300,
                        WithAnnotations = false
                    });

                using var pngImage = SkiaSharp.SKImage.FromBitmap(bitmap);
                using var pngData = pngImage.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                await System.IO.File.WriteAllBytesAsync(previewPath, pngData.ToArray(), cancellationToken);

                var pngFileInfo = new FileInfo(previewPath);
                _logger.LogInformation("[Stage 6/8] PNG saved in {Duration}ms: {PngPath} ({Size} bytes)",
                    sw.ElapsedMilliseconds, previewPath, pngFileInfo.Length);

                // ── Stage 7: Read PNG dimensions ───────────────────────
                int pngWidth = 0, pngHeight = 0;
                try
                {
                    using var pngStream = System.IO.File.OpenRead(previewPath);
                    using var pngBitmap = SkiaSharp.SKBitmap.Decode(pngStream);
                    pngWidth = pngBitmap.Width;
                    pngHeight = pngBitmap.Height;
                    _logger.LogInformation("[Stage 7/8] PNG dimensions: {Width}x{Height}", pngWidth, pngHeight);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read PNG dimensions from: {PngPath}", previewPath);
                }

                // ── Stage 8: Read comments via COM and compute field coordinates ──
                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogInformation("[Stage 8/8] Reading comments and measuring field geometry via COM...");

                var fields = MeasureFieldsFromCom(worksheet, printArea, pngWidth, pngHeight);

                _logger.LogInformation(
                    "[Stage 8/8] COM field measurement complete: {FieldCount} fields, page={W}x{H}px in {Duration}ms",
                    fields.Count, pngWidth, pngHeight, sw.ElapsedMilliseconds);

                return new CaptureResult
                {
                    ImageUrl = $"/preview/{previewFileName}",
                    Page = new PageInfo
                    {
                        Width = pngWidth,
                        Height = pngHeight
                    },
                    Fields = fields
                };
            }
            catch (COMException ex)
            {
                _logger.LogError(ex, "COM error while processing Excel file: {FilePath}", excelFilePath);
                throw new InvalidOperationException(
                    $"Failed to process the Excel file. " +
                    $"Ensure Microsoft Excel is installed on the server. Error: {ex.Message}", ex);
            }
            catch (PrintAreaNotConfiguredException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("Capture cancelled for file: {FilePath} ({Elapsed}ms)",
                    excelFilePath, sw.ElapsedMilliseconds);
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error while processing Excel file: {FilePath}", excelFilePath);
                throw new InvalidOperationException(
                    $"An unexpected error occurred while processing the file: {ex.Message}", ex);
            }
            finally
            {
                // Cleanup PDF (keep PNG for frontend)
                try
                {
                    if (pdfPath != null && System.IO.File.Exists(pdfPath))
                    {
                        System.IO.File.Delete(pdfPath);
                        _logger.LogDebug("Deleted intermediate PDF: {PdfPath}", pdfPath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to delete intermediate PDF: {PdfPath}", pdfPath);
                }

                ReleaseComObject(worksheet);
                CleanupComObjects(workbook, excelApp);

                sw.Stop();
                _logger.LogInformation(
                    "Capture completed for {FilePath} - Total: {Total}ms",
                    excelFilePath, sw.ElapsedMilliseconds);
            }
        }

        /// <summary>
        /// Reads all comments from the worksheet and measures each field's position
        /// using COM Range.Left/Top/Width/Height properties.
        /// Coordinates are converted from Excel points to PNG pixels at 300 DPI.
        /// Coordinates are offset by the printed page origin (margins + centering)
        /// so they align with the full-page PNG background (Phase 36).
        /// </summary>
        private List<ExcelField> MeasureFieldsFromCom(
            Excel.Worksheet worksheet,
            string printAreaAddress,
            int pngWidth,
            int pngHeight)
        {
            var fields = new List<ExcelField>();
            int tabIndex = 0;

            // Get the print area range to compute origin offset
            Excel.Range? printAreaRange = null;
            double printAreaOriginLeftPt = 0;
            double printAreaOriginTopPt = 0;
            double printedOriginXPt = 0;
            double printedOriginYPt = 0;

            try
            {
                printAreaRange = worksheet.Range[printAreaAddress];
                printAreaOriginLeftPt = GetDouble(printAreaRange.Left);
                printAreaOriginTopPt = GetDouble(printAreaRange.Top);

                _logger.LogDebug(
                    "Print area origin: Left={LeftPt:F2}pt, Top={TopPt:F2}pt",
                    printAreaOriginLeftPt, printAreaOriginTopPt);

                // ── Compute printed page origin from PageSetup ──────────
                // The PNG is a full printed page (612x792pt Letter). The print area
                // content is positioned within the page by margins and centering.
                // We must offset field coordinates by this printed origin so they
                // align with the cells in the full-page PNG.
                //
                // NOTE: COM interop properties require dynamic dispatch.
                // Cast through object first to bypass static type checking.
                try
                {
                    dynamic pageSetup = worksheet.PageSetup;

                    double pageWidthPt = 0;
                    double pageHeightPt = 0;
                    double leftMarginPt = 0;
                    double rightMarginPt = 0;
                    double topMarginPt = 0;
                    double bottomMarginPt = 0;
                    bool centerHorizontally = false;
                    bool centerVertically = false;

                    // PageWidth/PageHeight may need double PageSetup access in COM interop
                    // Fallback: use PaperWidth/PaperHeight which are on the same interface as margins
                    try { pageWidthPt = Convert.ToDouble(pageSetup.PageSetup.PageWidth); } catch { }
                    try { pageHeightPt = Convert.ToDouble(pageSetup.PageSetup.PageHeight); } catch { }
                    // Fallback: use PaperWidth/PaperHeight from PageSetup (available on primary interface)
                    if (pageWidthPt <= 0) { try { pageWidthPt = Convert.ToDouble(pageSetup.PaperWidth); } catch { } }
                    if (pageHeightPt <= 0) { try { pageHeightPt = Convert.ToDouble(pageSetup.PaperHeight); } catch { } }
                    // Last resort: assume Letter (612x792 pt)
                    if (pageWidthPt <= 0) { _logger.LogWarning("Page width not available from COM, assuming Letter (612pt)"); pageWidthPt = 612; }
                    if (pageHeightPt <= 0) { _logger.LogWarning("Page height not available from COM, assuming Letter (792pt)"); pageHeightPt = 792; }
                    try { leftMarginPt = Convert.ToDouble(pageSetup.LeftMargin); } catch { }
                    try { rightMarginPt = Convert.ToDouble(pageSetup.RightMargin); } catch { }
                    try { topMarginPt = Convert.ToDouble(pageSetup.TopMargin); } catch { }
                    try { bottomMarginPt = Convert.ToDouble(pageSetup.BottomMargin); } catch { }
                    try { centerHorizontally = Convert.ToBoolean(pageSetup.CenterHorizontally); } catch { }
                    try { centerVertically = Convert.ToBoolean(pageSetup.CenterVertically); } catch { }

                    // Content dimensions = print area range in points
                    double contentWidthPt = GetDouble(printAreaRange.Width);
                    double contentHeightPt = GetDouble(printAreaRange.Height);

                    (printedOriginXPt, printedOriginYPt) = _coordTransformer.ComputePrintedOrigin(
                        pageWidthPt, pageHeightPt,
                        leftMarginPt, rightMarginPt,
                        topMarginPt, bottomMarginPt,
                        contentWidthPt, contentHeightPt,
                        centerHorizontally, centerVertically);

                    _logger.LogInformation(
                        "[PHASE36] Printed page origin: ({OX:F2},{OY:F2})pt " +
                        "| page={PW:F1}x{PH:F1} content={CW:F1}x{CH:F1}",
                        printedOriginXPt, printedOriginYPt,
                        pageWidthPt, pageHeightPt,
                        contentWidthPt, contentHeightPt);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not compute printed page origin, using (0,0)");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not read print area Range geometry, using origin (0,0)");
            }

            // Read all comments from the worksheet
            try
            {
                var comments = worksheet.Comments;
                if (comments == null)
                {
                    _logger.LogInformation("No comments found on worksheet");
                    return fields;
                }

                int commentCount = comments.Count;
                _logger.LogDebug("Found {Count} comment(s) on worksheet", commentCount);

                for (int i = 1; i <= commentCount; i++)
                {
                    Excel.Comment? comment = null;
                    Excel.Range? cellRange = null;
                    Excel.Range? mergeArea = null;

                    try
                    {
                        comment = comments[i];
                        if (comment == null) continue;

                        // Get the cell that owns this comment
                        cellRange = comment.Parent as Excel.Range;
                        if (cellRange == null) continue;

                        // Get comment text
                        string commentText = GetCommentText(comment);
                        if (string.IsNullOrWhiteSpace(commentText)) continue;

                        // Get cell reference
                        string cellRef = GetCellRef(cellRange);
                        if (string.IsNullOrEmpty(cellRef)) continue;

                        // Check for merged cell — use merge area if applicable
                        double cellLeftPt, cellTopPt, cellWidthPt, cellHeightPt;
                        bool isMerged = false;
                        string? mergeAddress = null;

                        try
                        {
                            mergeArea = cellRange.MergeArea;
                            bool isInMerge = (bool)cellRange.MergeCells;
                            if (isInMerge && mergeArea != null)
                            {
                                isMerged = true;
                                mergeAddress = GetCellRef(mergeArea);
                                cellLeftPt = GetDouble(mergeArea.Left);
                                cellTopPt = GetDouble(mergeArea.Top);
                                cellWidthPt = GetDouble(mergeArea.Width);
                                cellHeightPt = GetDouble(mergeArea.Height);
                            }
                            else
                            {
                                cellLeftPt = GetDouble(cellRange.Left);
                                cellTopPt = GetDouble(cellRange.Top);
                                cellWidthPt = GetDouble(cellRange.Width);
                                cellHeightPt = GetDouble(cellRange.Height);
                            }
                        }
                        catch
                        {
                            cellLeftPt = GetDouble(cellRange.Left);
                            cellTopPt = GetDouble(cellRange.Top);
                            cellWidthPt = GetDouble(cellRange.Width);
                            cellHeightPt = GetDouble(cellRange.Height);
                        }

                        // Convert points to pixels RELATIVE TO PRINTED PAGE ORIGIN
                        // Formula: pixel = (printedOriginPt + cellPt - printAreaPt) * scale
                        // The printed origin accounts for margins + centering baked into the full-page PNG.
                        double leftPx = (printedOriginXPt + cellLeftPt - printAreaOriginLeftPt) * PointsToPixels;
                        double topPx = (printedOriginYPt + cellTopPt - printAreaOriginTopPt) * PointsToPixels;
                        double widthPx = cellWidthPt * PointsToPixels;
                        double heightPx = cellHeightPt * PointsToPixels;

                        // Infer field type from comment text
                        string fieldType = InferFieldType(commentText);

                        // Extract display name from comment first line
                        string fieldName = ExtractFieldName(commentText, cellRef, tabIndex);

                        // Generate unique internal ID: page{page}field{index}
                        string fieldId = $"page1field{fields.Count + 1}";

                        fields.Add(new ExcelField
                        {
                            Id = fieldId,
                            Name = fieldName,
                            Cell = cellRef.Replace("$", ""),
                            Type = fieldType,
                            Comment = commentText,
                            Left = Math.Round(leftPx, 1),
                            Top = Math.Round(topPx, 1),
                            Width = Math.Round(widthPx, 1),
                            Height = Math.Round(heightPx, 1),
                            ExcelLeft = Math.Round(cellLeftPt, 2),
                            ExcelTop = Math.Round(cellTopPt, 2),
                            PrintAreaLeft = Math.Round(printAreaOriginLeftPt, 2),
                            PrintAreaTop = Math.Round(printAreaOriginTopPt, 2),
                            ExcelWidthPt = Math.Round(cellWidthPt, 2),
                            ExcelHeightPt = Math.Round(cellHeightPt, 2),
                            IsMerged = isMerged,
                            MergeAddress = mergeAddress?.Replace("$", "")
                        });

                        tabIndex++;
                    }
                    finally
                    {
                        if (mergeArea != null) Marshal.ReleaseComObject(mergeArea);
                        if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                        if (comment != null) Marshal.ReleaseComObject(comment);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error reading comments via COM");
            }
            finally
            {
                if (printAreaRange != null) Marshal.ReleaseComObject(printAreaRange);
            }

            return fields;
        }

        // ═══════════════════════════════════════════════════════════════════
        // Helpers
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Safely get a double value from a COM object, handling DBNull and type issues.
        /// </summary>
        private static double GetDouble(object value)
        {
            if (value == null || value == DBNull.Value) return 0;
            try { return Convert.ToDouble(value); }
            catch { return 0; }
        }

        /// <summary>
        /// Extract the cell reference from a COM Range.
        /// </summary>
        private static string GetCellRef(Excel.Range range)
        {
            try
            {
                string? addr = range.Address[false, false] as string;
                return addr ?? "";
            }
            catch { return ""; }
        }

        /// <summary>
        /// Get the text of a COM Comment, handling multi-line text.
        /// </summary>
        private static string GetCommentText(Excel.Comment comment)
        {
            try
            {
                // Comment.Text() can return a string with newlines
                string? text = comment.Text() as string;
                return (text ?? "").Trim();
            }
            catch { return ""; }
        }

        /// <summary>
        /// Infer the field type from comment text content.
        /// Matches the frontend's inferOverlayType logic for consistency.
        /// </summary>
        private static string InferFieldType(string commentText)
        {
            string lower = commentText.ToLowerInvariant().Trim();

            // Exact and keyword-based matching
            if (lower == "textbox" || lower == "text" || lower == "text field" || lower == "input")
                return "Text";
            if (lower == "checkbox" || lower == "check box" || lower == "tick")
                return "Checkbox";
            if (lower == "signature" || lower == "sign here")
                return "Signature";
            if (lower == "date" || lower == "date field")
                return "Date";
            if (lower == "number" || lower == "number field" || lower == "amount" || lower == "price")
                return "Number";
            if (lower.Contains("checkbox") || lower.Contains("check box") || lower.Contains("tick"))
                return "Checkbox";
            if (lower.Contains("sign"))
                return "Signature";
            if (lower.Contains("date"))
                return "Date";
            if (lower.Contains("number") || lower.Contains("amount") || lower.Contains("price"))
                return "Number";

            // Default: treat as text field
            return "Text";
        }

        /// <summary>
        /// Extract a user-visible field name from comment text.
        /// Uses the first non-empty line of the comment. Falls back to a default name.
        /// </summary>
        private static string ExtractFieldName(string commentText, string cellRef, int index)
        {
            if (!string.IsNullOrWhiteSpace(commentText))
            {
                string firstLine = commentText
                    .Replace("\r\n", "\n")
                    .Split('\n')
                    .Select(l => l.Trim())
                    .FirstOrDefault(l => l.Length > 0) ?? "";

                if (!string.IsNullOrWhiteSpace(firstLine))
                    return firstLine;
            }
            return $"p1f{index + 1}";
        }

        // ═══════════════════════════════════════════════════════════════════
        // Print Area Detection
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Detect the configured print area using multiple fallback methods.
        /// </summary>
        private string? DetectPrintArea(Workbook workbook, Excel.Worksheet worksheet)
        {
            // Method 1: Direct PageSetup.PrintArea
            string? printArea = worksheet.PageSetup.PrintArea;
            if (!string.IsNullOrEmpty(printArea))
                return printArea;

            // Method 2: Workbook-level names for Print_Area
            foreach (Excel.Name name in workbook.Names)
            {
                try
                {
                    string nameValue = name.Name ?? "";
                    string? refersTo = name.RefersTo as string;

                    if (nameValue.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0
                        && !string.IsNullOrEmpty(refersTo))
                    {
                        string rangePart = refersTo;
                        int equalsIdx = rangePart.IndexOf('=');
                        if (equalsIdx >= 0)
                            rangePart = rangePart.Substring(equalsIdx + 1);

                        int bangIdx = rangePart.IndexOf('!');
                        if (bangIdx >= 0 && bangIdx < rangePart.Length - 1)
                            rangePart = rangePart.Substring(bangIdx + 1);

                        if (!string.IsNullOrEmpty(rangePart))
                            return rangePart;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(name);
                }
            }

            // Method 3: Worksheet-level names for Print_Area
            foreach (Excel.Name name in worksheet.Names)
            {
                try
                {
                    string nameValue = name.Name ?? "";
                    string? refersTo = name.RefersTo as string;

                    if (nameValue.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0
                        && !string.IsNullOrEmpty(refersTo))
                    {
                        string rangePart = refersTo;
                        int equalsIdx = rangePart.IndexOf('=');
                        if (equalsIdx >= 0)
                            rangePart = rangePart.Substring(equalsIdx + 1);

                        int bangIdx = rangePart.IndexOf('!');
                        if (bangIdx >= 0 && bangIdx < rangePart.Length - 1)
                            rangePart = rangePart.Substring(bangIdx + 1);

                        if (!string.IsNullOrEmpty(rangePart))
                            return rangePart;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(name);
                }
            }

            // Method 4: Force page refresh
            try
            {
                worksheet.Activate();
                workbook.Application.Calculate();
                workbook.Application.CalculateFull();

                printArea = worksheet.PageSetup.PrintArea;
                if (!string.IsNullOrEmpty(printArea))
                    return printArea;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[PrintArea:Method4] Page refresh failed.");
            }

            return null;
        }

        // ═══════════════════════════════════════════════════════════════════
        // COM Object Management
        // ═══════════════════════════════════════════════════════════════════

        private static void ReleaseComObject(object? comObject)
        {
            if (comObject == null) return;
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not release COM object: {ex.Message}");
            }
        }

        private void CleanupComObjects(Workbook? workbook, Application? excelApp)
        {
            if (workbook != null)
            {
                try
                {
                    workbook.Close(SaveChanges: false);
                }
                catch (COMException ex)
                {
                    _logger.LogWarning(ex, "Warning: Could not close workbook cleanly.");
                }
                finally
                {
                    Marshal.ReleaseComObject(workbook);
                }
            }

            if (excelApp != null)
            {
                try
                {
                    excelApp.Quit();
                }
                catch (COMException ex)
                {
                    _logger.LogWarning(ex, "Warning: Could not quit Excel application cleanly.");
                }
                finally
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
