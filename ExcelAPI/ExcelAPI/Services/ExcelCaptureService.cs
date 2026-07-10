using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Services.Interfaces;

namespace ExcelAPI.Services
{
    /// <summary>
    /// Service that uses Microsoft Excel COM Interop to open an Excel workbook,
    /// read the configured print area from the first worksheet,
    /// and capture it as a PNG image.
    ///
    /// Rendering Pipeline:
    ///   1. Export the worksheet's Print Area to PDF via Excel's ExportAsFixedFormat
    ///      (uses Excel's own print engine — same output as Print Preview)
    ///   2. Convert the PDF to a high-resolution PNG (300 DPI) via PDFtoImage (PDFium)
    ///   3. Extract field metadata from cell comments/notes
    ///
    /// Coordinate System:
    ///   - Field positions are relative to the Print Area origin (not worksheet origin)
    ///   - Scale is constant: pointsToPixels = DPI / 72 = 300 / 72 ≈ 4.1667
    ///   - pixel = (cellPoint - printAreaOriginPoint) * (DPI / 72)
    ///   - NOTE: pngWidth/printAreaWidth is NOT used as scale because the PNG
    ///     includes page margins while the print area range does not. Using that
    ///     ratio would inflate coordinates by the margin factor (typically ~1.2x).
    /// </summary>
    public class ExcelCaptureService : IExcelCaptureService
    {
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<ExcelCaptureService> _logger;

        // ── Phase 10: Runtime State Capture ────────────────────────────────
        private readonly List<RuntimeStateSnapshot> _runtimeSnapshots = new();
        private const string RuntimeCaptureDir = "RuntimeCapture";

        public ExcelCaptureService(
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<ExcelCaptureService> logger)
        {
            _env = env;
            _options = options;
            _logger = logger;
        }

        /// <inheritdoc />
        public async Task<CaptureResult> CapturePrintAreaAsync(string excelFilePath, string? fileId = null, CancellationToken cancellationToken = default)
        {
            // Ensure the Excel file exists
            if (!System.IO.File.Exists(excelFilePath))
            {
                throw new FileNotFoundException("Excel file not found.", excelFilePath);
            }

            // Performance tracking
            var sw = Stopwatch.StartNew();
            long excelStartTime = 0;
            long workbookOpenTime = 0;
            long pdfExportTime = 0;
            long conversionTime = 0;
            long cleanupTime = 0;

            Application? excelApp = null;
            Workbook? workbook = null;
            Excel.Worksheet? worksheet = null;
            string? pdfPath = null;

            try
            {
                _logger.LogInformation("Capture started for file: {FilePath}", excelFilePath);

                // --- Step 3: Launch Microsoft Excel (hidden, no alerts) ---
                cancellationToken.ThrowIfCancellationRequested();

                _logger.LogDebug("Launching Microsoft Excel application...");
                excelApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                excelStartTime = sw.ElapsedMilliseconds;
                _logger.LogInformation("Excel started in {Duration}ms", excelStartTime);

                // Ensure the Preview directory exists
                string previewFolder = Path.Combine(_env.ContentRootPath, _options.Value.PreviewDirectory);
                Directory.CreateDirectory(previewFolder);

                cancellationToken.ThrowIfCancellationRequested();

                // --- Step 4: Open the uploaded workbook ---
                _logger.LogDebug("Opening workbook: {FilePath}", excelFilePath);
                workbook = excelApp.Workbooks.Open(excelFilePath);
                workbookOpenTime = sw.ElapsedMilliseconds;
                _logger.LogInformation("Workbook opened in {Duration}ms", workbookOpenTime - excelStartTime);

                // --- Step 5: Select the first VISIBLE worksheet ---
                // Skip hidden metadata sheets (e.g., "_Fields" created by VSTO Add-in)
                // that may appear before the user's actual worksheet.
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
                        // Release non-visible worksheets; keep the visible one for use
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

                // --- Step 6: Read the configured Print Area (robust detection with fallbacks) ---
                int detectionMethod = 0;

                // Method 1: Direct PageSetup.PrintArea (standard approach, works for most workbooks)
                string? printArea = worksheet.PageSetup.PrintArea;

                if (!string.IsNullOrEmpty(printArea))
                {
                    detectionMethod = 1;
                }

                if (string.IsNullOrEmpty(printArea))
                {
                    // Log diagnostic info for Method 1 failure
                    // NOTE: Must release UsedRange COM object to avoid leaks
                    string? usedRangeAddress = null;
                    try
                    {
                        var usedRange = worksheet.UsedRange;
                        if (usedRange != null)
                        {
                            usedRangeAddress = usedRange.Address;
                            Marshal.ReleaseComObject(usedRange);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug(ex, "[PrintArea:Method1] Failed to read UsedRange for empty sheet.");
                    }

                    _logger.LogInformation(
                        "[PrintArea:Method1] Failed. Worksheet=\"{Name}\", UsedRange={Address}",
                        worksheet.Name,
                        usedRangeAddress ?? "(unavailable)");

                    // Method 2: Search workbook-level Names for Print_Area
                    // Some workbooks (especially older/external ones) store the print area
                    // as a workbook-level name rather than in PageSetup.
                    _logger.LogDebug("[PrintArea:Method2] Searching workbook-level names for Print_Area...");
                    foreach (Excel.Name name in workbook.Names)
                    {
                        try
                        {
                            string nameValue = name.Name ?? "";
                            string? refersTo = name.RefersTo as string;
                            _logger.LogInformation(
                                "[PrintArea:Method2] Workbook Name: \"{Name}\" -> {RefersTo}",
                                nameValue,
                                refersTo ?? "(null)");

                            if (nameValue.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0
                                && !string.IsNullOrEmpty(refersTo))
                            {
                                // Extract the range reference from the RefersTo formula
                                // Format: "=Sheet1!$A$1:$C$10" or "=$A$1:$C$10"
                                int equalsIdx = refersTo.IndexOf('=');
                                string rangePart = equalsIdx >= 0
                                    ? refersTo.Substring(equalsIdx + 1)
                                    : refersTo;

                                // If it includes a sheet name (e.g., "Sheet1!$A$1:$C$10"),
                                // extract only the range portion after '!'
                                int bangIdx = rangePart.IndexOf('!');
                                if (bangIdx >= 0 && bangIdx < rangePart.Length - 1)
                                {
                                    rangePart = rangePart.Substring(bangIdx + 1);
                                }

                                printArea = rangePart;
                                detectionMethod = 2;
                                _logger.LogInformation(
                                    "[PrintArea:Method2] Found via workbook name. RefersTo=\"{RefersTo}\", Extracted=\"{PrintArea}\"",
                                    refersTo,
                                    printArea);
                                break;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(name);
                        }
                    }
                }

                if (string.IsNullOrEmpty(printArea))
                {
                    // Method 3: Search worksheet-level Names for Print_Area
                    // Some workbooks scope the Print_Area name to a specific worksheet.
                    _logger.LogDebug("[PrintArea:Method3] Searching worksheet-level names for Print_Area...");
                    foreach (Excel.Name name in worksheet.Names)
                    {
                        try
                        {
                            string nameValue = name.Name ?? "";
                            string? refersTo = name.RefersTo as string;
                            _logger.LogInformation(
                                "[PrintArea:Method3] Worksheet Name: \"{Name}\" -> {RefersTo}",
                                nameValue,
                                refersTo ?? "(null)");

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

                                printArea = rangePart;
                                detectionMethod = 3;
                                _logger.LogInformation(
                                    "[PrintArea:Method3] Found via worksheet name. RefersTo=\"{RefersTo}\", Extracted=\"{PrintArea}\"",
                                    refersTo,
                                    printArea);
                                break;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(name);
                        }
                    }
                }

                if (string.IsNullOrEmpty(printArea))
                {
                    // Method 4: Force Excel to refresh page setup before reading again.
                    // Some workbooks' PageSetup values are not populated until the sheet
                    // is activated or formulas are recalculated.
                    _logger.LogDebug("[PrintArea:Method4] Attempting page refresh (Activate + Calculate + CalculateFull)...");
                    try
                    {
                        worksheet.Activate();
                        excelApp.Calculate();
                        excelApp.CalculateFull();

                        printArea = worksheet.PageSetup.PrintArea;

                        if (!string.IsNullOrEmpty(printArea))
                        {
                            detectionMethod = 4;
                            _logger.LogInformation(
                                "[PrintArea:Method4] Print area found after page refresh: \"{PrintArea}\"",
                                printArea);
                        }
                        else
                        {
                            _logger.LogInformation(
                                "[PrintArea:Method4] Still empty after page refresh.");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "[PrintArea:Method4] Page refresh failed.");
                    }
                }

                if (string.IsNullOrEmpty(printArea))
                {
                    // Method 5: Log ALL worksheets and their PrintArea values for diagnosis.
                    // This helps identify whether the print area was configured on a
                    // different worksheet than the one being processed.
                    _logger.LogWarning(
                        "[PrintArea:Method5] All detection methods failed. Logging all worksheets...");

                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        try
                        {
                            string? wsUsedRange = null;
                            try
                            {
                                var wsRange = ws.UsedRange;
                                if (wsRange != null)
                                {
                                    wsUsedRange = wsRange.Address;
                                    Marshal.ReleaseComObject(wsRange);
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogDebug(ex, "[PrintArea:Method5] Failed to read UsedRange for sheet \"{Name}\".", ws.Name);
                            }

                            _logger.LogInformation(
                                "[PrintArea:Method5] Sheet#{Index}: \"{Name}\" | PrintArea={PrintArea} | UsedRange={Address}",
                                ws.Index,
                                ws.Name,
                                ws.PageSetup.PrintArea ?? "(none)",
                                wsUsedRange ?? "(unavailable)");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[PrintArea:Method5] Failed to read worksheet.");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(ws);
                        }
                    }

                    // All methods exhausted — no print area truly exists
                    throw new PrintAreaNotConfiguredException(
                        "No print area is configured in this worksheet. " +
                        "Please set a print area (Page Layout > Print Area > Set Print Area) " +
                        "in the Excel file before uploading.");
                }

                _logger.LogInformation(
                    "Print area detected via Method#{Method}: \"{PrintArea}\" on worksheet \"{Name}\"",
                    detectionMethod,
                    printArea,
                    worksheet.Name);

                // --- Step 6a: Read Print Area Range origin and dimensions ---
                // These are used as the coordinate origin for field positioning.
                // All field coordinates are relative to the Print Area, not the worksheet.
                double printAreaLeft = 0, printAreaTop = 0, printAreaWidth = 0, printAreaHeight = 0;
                int printAreaCols = 0, printAreaRows = 0;
                double? firstColWidthPt = null, firstRowHeightPt = null;
                Excel.Range? printRange = null;
                try
                {
                    printRange = worksheet.Range[printArea];
                    printAreaLeft = printRange.Left;
                    printAreaTop = printRange.Top;
                    printAreaWidth = printRange.Width;
                    printAreaHeight = printRange.Height;
                    printAreaCols = printRange.Columns.Count;
                    printAreaRows = printRange.Rows.Count;

                    // Get first column and first row measurements
                    try
                    {
                        var firstCol = printRange.Columns[1];
                        firstColWidthPt = firstCol.Width;
                        Marshal.ReleaseComObject(firstCol);
                    }
                    catch { }
                    try
                    {
                        var firstRow = printRange.Rows[1];
                        firstRowHeightPt = firstRow.Height;
                        Marshal.ReleaseComObject(firstRow);
                    }
                    catch { }

                    _logger.LogInformation(
                        "[AUDIT] PrintArea Range: Address=\"{Address}\" | " +
                        "Origin: Left={Left:F1}pt, Top={Top:F1}pt | " +
                        "Content: Width={Width:F1}pt, Height={Height:F1}pt | " +
                        "Grid: {Cols}col x {Rows}row | " +
                        "FirstColWidth={FCW:F1}pt, FirstRowHeight={FRH:F1}pt",
                        printRange.AddressLocal[false, false],
                        printAreaLeft, printAreaTop,
                        printAreaWidth, printAreaHeight,
                        printAreaCols, printAreaRows,
                        firstColWidthPt ?? 0, firstRowHeightPt ?? 0);
                }
                finally
                {
                    if (printRange != null)
                        Marshal.ReleaseComObject(printRange);
                }

                // --- Step 6a.5: Compute content width from column widths ---
                // Replace COM Range.Width with a width computed from XLSX worksheet XML.
                // This accounts for the Calibri 11pt → Aptos Narrow 11pt font change
                // that affects default column width in centering calculations.
                double contentWidthPt = ComputeContentWidthFromXlsx(
                    excelFilePath, worksheet?.Name, worksheet?.Index ?? 0, printArea, printAreaCols);
                if (contentWidthPt > 0)
                {
                    _logger.LogInformation(
                        "[COLS] Content width override: Range.Width={Old:F2}pt -> Computed={New:F2}pt ({Cols} cols)",
                        printAreaWidth, contentWidthPt, printAreaCols);
                    printAreaWidth = contentWidthPt;
                }
                else
                {
                    _logger.LogInformation(
                        "[COLS] Using Range.Width={Width:F2}pt ({Cols} cols) — no XLSX column data available",
                        printAreaWidth, printAreaCols);
                }

                // --- Step 6b: Read Page Setup and calculate printed page offset ---
                // The PNG shows the PRINTED PAGE (with margins, centering, etc.),
                // not the raw worksheet. We need to determine where the print area
                // content actually appears on the rendered page.
                double leftMarginPt = 0, topMarginPt = 0, rightMarginPt = 0, bottomMarginPt = 0;
                bool centerHorizontally = false, centerVertically = false;
                int zoomSetting = 0;
                int fitToPagesWide = 0, fitToPagesTall = 0;
                double pageWidthPt = 612, pageHeightPt = 792;  // Default to Letter
                XlPageOrientation orientation = XlPageOrientation.xlPortrait;
                double printedOriginXPts = 0, printedOriginYPts = 0;

                try
                {
                    // Read page setup values
                    leftMarginPt = worksheet.PageSetup.LeftMargin;
                    rightMarginPt = worksheet.PageSetup.RightMargin;
                    topMarginPt = worksheet.PageSetup.TopMargin;
                    bottomMarginPt = worksheet.PageSetup.BottomMargin;
                    centerHorizontally = worksheet.PageSetup.CenterHorizontally;
                    centerVertically = worksheet.PageSetup.CenterVertically;
                    orientation = worksheet.PageSetup.Orientation;

                    // Zoom and FitToPages are mutually exclusive
                    try
                    {
                        zoomSetting = worksheet.PageSetup.Zoom;
                    }
                    catch { /* Zoom may be null/0 when FitToPages is active */ }

                    try
                    {
                        fitToPagesWide = worksheet.PageSetup.FitToPagesWide;
                    }
                    catch { }

                    try
                    {
                        fitToPagesTall = worksheet.PageSetup.FitToPagesTall;
                    }
                    catch { }

                    // Map paper size to points
                    (pageWidthPt, pageHeightPt) = GetPaperSizePoints(
                        (XlPaperSize)(int)worksheet.PageSetup.PaperSize,
                        orientation);

                    // Calculate where the print area content starts on the printed page
                    // The printable area is the page minus margins.
                    // If centering is enabled, the content is centered within the printable area.
                    double printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt;
                    double printableHeightPt = pageHeightPt - topMarginPt - bottomMarginPt;

                    // Base origin: top-left of printable area (inside margins)
                    printedOriginXPts = leftMarginPt;
                    printedOriginYPts = topMarginPt;

                    // Apply horizontal centering: shift content right by half the unused space
                    if (centerHorizontally && printAreaWidth < printableWidthPt)
                    {
                        printedOriginXPts += (printableWidthPt - printAreaWidth) / 2.0;
                    }

                    // Apply vertical centering: shift content down by half the unused space
                    if (centerVertically && printAreaHeight < printableHeightPt)
                    {
                        printedOriginYPts += (printableHeightPt - printAreaHeight) / 2.0;
                    }

                    // Log complete page setup diagnostics
                    _logger.LogInformation(
                        "[PAGE] Setup: " +
                        "Paper={PaperSize}x{Orientation} ({PageW:F1}x{PageH:F1}pt) | " +
                        "Margins: L={LM:F1} R={RM:F1} T={TM:F1} B={BM:F1}pt | " +
                        "Printable: {PW:F1}x{PH:F1}pt | " +
                        "CenterH={CH}, CenterV={CV} | " +
                        "Zoom={Zoom}, FitToPages={FTPW}x{FTPH} | " +
                        "PrintAreaContent: {PAW:F1}x{PAH:F1}pt | " +
                        "PrintedOrigin: ({OX:F1},{OY:F1})pt",
                        (int)worksheet.PageSetup.PaperSize, orientation,
                        pageWidthPt, pageHeightPt,
                        leftMarginPt, rightMarginPt, topMarginPt, bottomMarginPt,
                        printableWidthPt, printableHeightPt,
                        centerHorizontally, centerVertically,
                        zoomSetting, fitToPagesWide, fitToPagesTall,
                        printAreaWidth, printAreaHeight,
                        printedOriginXPts, printedOriginYPts);

                    // Warn if FitToPages or Zoom will affect content scaling
                    if (fitToPagesWide > 0 || fitToPagesTall > 0)
                    {
                        _logger.LogWarning(
                            "[PAGE] FitToPages is active ({FTPW}x{FTPH}) — " +
                            "Excel will scale the print area content to fit the page. " +
                            "The DPI/72 pixel scale assumes 100% zoom and the rendered " +
                            "cell dimensions may differ from cellWidth * DPI/72. " +
                            "Field overlay accuracy may be reduced for this workbook.",
                            fitToPagesWide, fitToPagesTall);
                    }
                    else if (zoomSetting > 0 && zoomSetting != 100)
                    {
                        _logger.LogWarning(
                            "[PAGE] Custom Zoom is set to {Zoom}% — " +
                            "Excel will scale the content by {Zoom}%. " +
                            "The DPI/72 pixel scale assumes 100% zoom; fields may " +
                            "need additional scaling by {Zoom/100:F2}x.",
                            zoomSetting, zoomSetting, zoomSetting / 100.0);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[PAGE] Failed to read page setup, using defaults.");
                }

                // Convert printed origin to PNG pixels at DPI
                const double dpi = 300.0;
                const double pointsToPixels = dpi / 72.0;
                double printedOriginX = printedOriginXPts * pointsToPixels;
                double printedOriginY = printedOriginYPts * pointsToPixels;

                _logger.LogInformation(
                    "[PAGE] Printed origin in pixels: ({OX:F1},{OY:F1})px (at {DPI} DPI)",
                    printedOriginX, printedOriginY, dpi);

                // --- Step 7 & 8: Export Print Area to PDF, then convert to PNG ---

                // Generate unique filenames for the intermediate PDF and final PNG
                string localFileId = fileId ?? Guid.NewGuid().ToString("N");
                string pdfFileName = $"page_{localFileId}.pdf";
                string previewFileName = $"page_{localFileId}.png";
                pdfPath = Path.Combine(previewFolder, pdfFileName);
                string previewPath = Path.Combine(previewFolder, previewFileName);

                // Step 7a: Export the worksheet's print area to PDF using Excel's print engine
                cancellationToken.ThrowIfCancellationRequested();
                _logger.LogDebug("Exporting print area to PDF: {PdfPath}", pdfPath);

                worksheet.ExportAsFixedFormat(
                    Type: XlFixedFormatType.xlTypePDF,
                    Filename: pdfPath,
                    Quality: XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: false,
                    IgnorePrintAreas: false,  // Respect the configured print area
                    OpenAfterPublish: false);

                pdfExportTime = sw.ElapsedMilliseconds;

                // Validate the generated PDF exists and has content
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

                _logger.LogInformation("PDF exported in {Duration}ms: {PdfPath}, Size: {Size} bytes",
                    pdfExportTime - workbookOpenTime,
                    pdfPath,
                    pdfFileInfo.Length);

                cancellationToken.ThrowIfCancellationRequested();

                // Step 7b: Convert the PDF to a high-resolution PNG (300 DPI) using PDFtoImage
                // NOTE: PDFtoImage's ToImage(string) overload expects Base64 string, not a file path.
                // We must read the PDF as a byte array and use the byte[] overload.
                _logger.LogDebug("Converting PDF to PNG at 300 DPI: {PngPath}", previewPath);

                // Read the PDF file into a byte array
                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath, cancellationToken);

                // PDFtoImage uses PDFium (Chrome's PDF engine) + SkiaSharp for rendering
                using var bitmap = PDFtoImage.Conversion.ToImage(
                    pdfBytes,
                    page: 0,  // First (and only) page
                    options: new PDFtoImage.RenderOptions
                    {
                        Dpi = 300,  // High resolution for print-quality output
                        WithAnnotations = false
                    });

                // Encode the SKBitmap as PNG and write to disk
                using var pngImage = SkiaSharp.SKImage.FromBitmap(bitmap);
                using var pngData = pngImage.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                await System.IO.File.WriteAllBytesAsync(previewPath, pngData.ToArray(), cancellationToken);

                conversionTime = sw.ElapsedMilliseconds;
                _logger.LogInformation("PNG saved to: {PngPath} (Size: {Size} bytes, conversion: {Duration}ms)",
                    previewPath,
                    new FileInfo(previewPath).Length,
                    conversionTime - pdfExportTime);

                cancellationToken.ThrowIfCancellationRequested();

                // --- Step 10: Read PNG dimensions ---
                int pngWidth = 0, pngHeight = 0;
                try
                {
                    using var pngStream = System.IO.File.OpenRead(previewPath);
                    using var pngBitmap = SkiaSharp.SKBitmap.Decode(pngStream);
                    pngWidth = pngBitmap.Width;
                    pngHeight = pngBitmap.Height;
                    _logger.LogInformation("PNG dimensions: {Width}x{Height}", pngWidth, pngHeight);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read PNG dimensions from: {PngPath}", previewPath);
                }

                // --- Step 11: Compute actual point-to-pixel scale from rendered PNG ---
                // Always use the actual rendered dimensions, not the assumed DPI/72.
                // scaleX = pngWidth / pageWidthPt
                // scaleY = pngHeight / pageHeightPt

                double scaleX = pngWidth > 0 && pageWidthPt > 0
                    ? pngWidth / pageWidthPt
                    : dpi / 72.0;
                double scaleY = pngHeight > 0 && pageHeightPt > 0
                    ? pngHeight / pageHeightPt
                    : dpi / 72.0;

                double theoreticalScale = dpi / 72.0;
                double scaleRatioX = scaleX / theoreticalScale;
                double scaleRatioY = scaleY / theoreticalScale;

                // Recompute printed origin using the ACTUAL scale
                double actualPrintedOriginX = printedOriginXPts * scaleX;
                double actualPrintedOriginY = printedOriginYPts * scaleY;

                _logger.LogInformation(
                    "[SCALE] Actual X={SX:F6} Y={SY:F6} theoretical={T:F6} " +
                    "ratio X={RX:F6} Y={RY:F6} | png={PW}x{PH}px page={PPW:F1}x{PPH:F1}pt",
                    scaleX, scaleY, theoreticalScale,
                    scaleRatioX, scaleRatioY,
                    pngWidth, pngHeight, pageWidthPt, pageHeightPt);

                _logger.LogInformation(
                    "[AUDIT] Final coordinate system:\n" +
                    "  Printed origin: ({OXP:F1},{OYP:F1})pt -> ({OX:F1},{OY:F1})px\n" +
                    "  Scale: {S:F6} px/pt (theoretical {T:F6})",
                    printedOriginXPts, printedOriginYPts,
                    actualPrintedOriginX, actualPrintedOriginY, scaleX, theoreticalScale);

                // --- Step 12: Extract fields using ACTUAL scale and origin ---
                _logger.LogDebug("Extracting field metadata from cell comments...");
                var fields = ExtractFields(
                    worksheet,
                    printAreaLeft, printAreaTop,
                    actualPrintedOriginX, actualPrintedOriginY,
                    scaleX, scaleY,
                    pageWidthPt, pageHeightPt);

                // Log coordinate conversion trace for the first field (comparison table)
                if (fields.Count > 0)
                {
                    var f = fields[0];
                    double offsetLeft = f.ExcelLeft - f.PrintAreaLeft;
                    double offsetTop = f.ExcelTop - f.PrintAreaTop;
                    _logger.LogInformation(
                        "[COORD] First field \"{Cell}\" | " +
                        "PrintedOrigin=({OX:F1},{OY:F1})px + " +
                        "(ExcelPt=({EL:F1},{ET:F1}) - OriginPt=({OL:F1},{OT:F1}))" +
                        " x {Scale:F4}px/pt => " +
                        "Pixel=({PL:F1},{PT:F1})px Size=({PW:F1}x{PH:F1})px",
                        f.Cell,
                        actualPrintedOriginX, actualPrintedOriginY,
                        f.ExcelLeft, f.ExcelTop,
                        f.PrintAreaLeft, f.PrintAreaTop,
                        scaleX,
                        f.Left, f.Top,
                        f.Width, f.Height);
                }

                // ================================================================
                // Phase 2.8 - Render Calibration Test (Bypass PDF)
                // Tests whether Excel->PDF or PDF->PNG introduces the error.
                // ================================================================
                if (fields.Count > 0 && pdfPath != null && System.IO.File.Exists(pdfPath))
                {
                    try
                    {
                        // ================================================================
                        // Test 1: MergeArea geometry from Excel COM
                        // ================================================================
                        try
                        {
                            var testField1 = fields[0];
                            if (worksheet != null && !string.IsNullOrEmpty(testField1.Cell))
                            {
                                Excel.Range? fr = null;
                                try
                                {
                                    fr = worksheet.Range[testField1.Cell];
                                    var ma = fr.MergeArea;
                                    if (ma != null)
                                    {
                                        // COM properties are dynamic; cast to explicit types to avoid CS1973
                                        string cellAddr = ma.Address ?? "";
                                        double maLeft = ma.Left;
                                        double maTop = ma.Top;
                                        double maWidth = ma.Width;
                                        double maHeight = ma.Height;
                                        int maRows = ma.Rows.Count;
                                        int maCols = ma.Columns.Count;
                                        _logger.LogInformation(
                                            "[CALIB:1] MergeArea Geometry for \"{Cell}\" - " +
                                            "Address=\"{Address}\" | " +
                                            "Left={L:F2} Top={T:F2} Width={W:F2} Height={H:F2} | " +
                                            "Rows={Rows} Cols={Cols}",
                                            testField1.Cell, cellAddr,
                                            maLeft, maTop, maWidth, maHeight,
                                            maRows, maCols);
                                        Marshal.ReleaseComObject(ma);
                                    }
                                }
                                finally
                                {
                                    if (fr != null) Marshal.ReleaseComObject(fr);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[CALIB:1] MergeArea diagnostic failed: {Msg}", ex.Message);
                        }

                        // ================================================================
                        // Test 2: Annotated PDF - Draw red rectangle at calculated coords.
                        // If rectangle aligns with cell -> Excel->PDF coords are CORRECT.
                        //   Bug is in PDF->PNG (PDFium rendering layer).
                        // If rectangle is offset -> Bug is in coordinate calculation.
                        // ================================================================
                        try
                        {
                            byte[] pdfBytesAnn = System.IO.File.ReadAllBytes(pdfPath);
                            using var pdfBmp = PDFtoImage.Conversion.ToImage(
                                pdfBytesAnn,
                                page: 0,
                                options: new PDFtoImage.RenderOptions { Dpi = 300, WithAnnotations = false });

                            if (pdfBmp != null)
                            {
                                using var annBmp = new SkiaSharp.SKBitmap(pdfBmp.Width, pdfBmp.Height);
                                using var srcPix = pdfBmp.PeekPixels();
                                using var dstPix = annBmp.PeekPixels();
                                srcPix.ReadPixels(dstPix.Info, dstPix.GetPixels(), dstPix.RowBytes, 0, 0);

                                using var canvas = new SkiaSharp.SKCanvas(annBmp);
                                using var redPen = new SkiaSharp.SKPaint
                                {
                                    Color = new SkiaSharp.SKColor(255, 0, 0, 160),
                                    Style = SkiaSharp.SKPaintStyle.Stroke,
                                    StrokeWidth = 3
                                };
                                using var crossPen = new SkiaSharp.SKPaint
                                {
                                    Color = new SkiaSharp.SKColor(255, 0, 0, 220),
                                    Style = SkiaSharp.SKPaintStyle.Stroke,
                                    StrokeWidth = 1
                                };

                                var tf = fields[0];
                                float x0 = (float)tf.Left;
                                float y0 = (float)tf.Top;
                                float ww = (float)tf.Width;
                                float hh = (float)tf.Height;
                                canvas.DrawRect(x0, y0, ww, hh, redPen);
                                canvas.DrawLine(x0 - 5, y0, x0 + 5, y0, crossPen);
                                canvas.DrawLine(x0, y0 - 5, x0, y0 + 5, crossPen);

                                string annPath = Path.Combine(previewFolder, $"annotated_{localFileId}.png");
                                using var annImg = SkiaSharp.SKImage.FromBitmap(annBmp);
                                using var annData = annImg.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                                System.IO.File.WriteAllBytes(annPath, annData.ToArray());

                                _logger.LogInformation(
                                    "[CALIB:2] Annotated PNG saved: {Path}\n" +
                                    "  Red rectangle at calculated coords:\n" +
                                    "    L={L:F1} T={T:F1} W={W:F1} H={H:F1}\n" +
                                    "  If rectangle ALIGNS with cell -> PDF coords are CORRECT.\n" +
                                    "    Bug is in PDF->PNG (PDFium rendering).\n" +
                                    "  If rectangle is OFFSET -> Bug is in coordinate calc\n" +
                                    "    (Excel->PDF transform or printed origin is wrong).",
                                    annPath,
                                    tf.Left, tf.Top, tf.Width, tf.Height);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[CALIB:2] Annotated PDF test failed: {Msg}", ex.Message);
                        }

                        // ================================================================
                        // Test 4: Multi-DPI rendering - detect if error changes with DPI.
                        // Render PDF at 72/150/300/600 DPI, measure first field dimensions.
                        // If error % CHANGES with DPI -> PDFium applies additional scaling.
                        // If error % STAYS THE SAME -> PDF itself is already transformed.
                        // ================================================================
                        try
                        {
                            byte[] pdfBytesM = System.IO.File.ReadAllBytes(pdfPath);
                            var mf = fields[0];
                            double fwPt = mf.ExcelWidthPt;
                            double fhPt = mf.ExcelHeightPt;

                            int[] dpiVals = { 72, 150, 300, 600 };
                            string dpiLog = "";

                            foreach (int td in dpiVals)
                            {
                                using var db = PDFtoImage.Conversion.ToImage(
                                    pdfBytesM,
                                page: 0,
                                options: new PDFtoImage.RenderOptions { Dpi = td, WithAnnotations = false });
                                if (db == null) continue;

                                double ts = td / 72.0;
                                double exW = fwPt * ts;
                                double exH = fhPt * ts;

                                int dw = db.Width;
                                int dh = db.Height;

                                // Actual scale at this DPI from rendered image
                                double actualScaleForDPI = pageWidthPt > 0 ? (double)dw / pageWidthPt : ts;

                                // Expected position using the actual rendered scale
                                int eL = (int)Math.Round(printedOriginXPts * actualScaleForDPI + (mf.ExcelLeft - printAreaLeft) * actualScaleForDPI);
                                int eT = (int)Math.Round(printedOriginY * td / 300.0 + (mf.ExcelTop - printAreaTop) * ts);
                                int eR = eL + (int)Math.Round(exW);
                                int eB = eT + (int)Math.Round(exH);
                                int eMx = (eL + eR) / 2;
                                int eMy = (eT + eB) / 2;

                                int aL = eL, aT = eT, aR = eR, aB = eB;
                                if (eMx >= 0 && eMx < dw && eMy >= 0 && eMy < dh)
                                {
                                    const int mg = 20;
                                    const byte th = 240;

                                    int sY = Math.Clamp(eMy, 0, dh - 1);
                                    for (int x = Math.Max(0, eL - mg); x <= Math.Min(dw - 1, eL); x++)
                                    { var p = db.GetPixel(x, sY); if (p.Red < th || p.Green < th || p.Blue < th) { aL = x; break; } }
                                    for (int x = Math.Max(0, eR); x < Math.Min(dw, eR + mg); x++)
                                    { var p = db.GetPixel(x, sY); if (p.Red < th || p.Green < th || p.Blue < th) aR = x; else if (x > aR + 3) break; }

                                    int sX = Math.Clamp(eMx, 0, dw - 1);
                                    for (int y = Math.Max(0, eT - mg); y <= Math.Min(dh - 1, eT); y++)
                                    { var p = db.GetPixel(sX, y); if (p.Red < th || p.Green < th || p.Blue < th) { aT = y; break; } }
                                    for (int y = Math.Max(0, eB); y < Math.Min(dh, eB + mg); y++)
                                    { var p = db.GetPixel(sX, y); if (p.Red < th || p.Green < th || p.Blue < th) aB = y; else if (y > aB + 3) break; }
                                }

                                double aW = aR - aL;
                                double aH = aB - aT;
                                double wR = exW > 0 ? aW / exW : 0;
                                double hR = exH > 0 ? aH / exH : 0;

                                dpiLog += string.Format(
                                    "  {0}DPI: thScale={1:F4} actScale={2:F4} expW={3:F1} actW={4:F1} wRatio={5:F6} expH={6:F1} actH={7:F1} hRatio={8:F6}\n",
                                    td, ts, actualScaleForDPI, exW, aW, wR, exH, aH, hR);
                            }

                            _logger.LogInformation(
                                "[CALIB:4] Multi-DPI Rendering Test - Field \"{Cell}\" ({WPt:F1}x{HPt:F1}pt)\n" +
                                "{DpiLog}" +
                                "  ANALYSIS:\n" +
                                "  If ratios are CONSTANT across all DPIs:\n" +
                                "    - PDF itself is already scaled (bug in Excel->PDF export).\n" +
                                "    Fix: Adjust DPI/72 scale by the observed ratio.\n" +
                                "  If ratios CHANGE with DPI:\n" +
                                "    - PDFium applies its own scaling on PDF content.\n" +
                                "    Fix: Investigate PDFium render options.\n" +
                                "  If ALL ratios ~= 1.0:\n" +
                                "    - Coordinate system IS correct. Problem elsewhere.\n" +
                                "  Reference: At 300 DPI, expected scale = 300/72 = {Ref:F4}",
                                mf.Cell ?? "", fwPt, fhPt,
                                dpiLog,
                                scaleX);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[CALIB:4] Multi-DPI test failed: {Msg}", ex.Message);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "[CALIB] Render calibration test failed: {Msg}", ex.Message);
                    }
                }

                // Build and return the complete capture result
                return new CaptureResult
                {
                    ImageUrl = $"/preview/{previewFileName}",
                    Page = new PageInfo
                    {
                        Width = pngWidth,
                        Height = pngHeight
                    },
                    Fields = fields,
                    PageSetup = new PageSetupDebug
                    {
                        PrintedOriginX = Math.Round(actualPrintedOriginX, 2),
                        PrintedOriginY = Math.Round(actualPrintedOriginY, 2),
                        LeftMargin = leftMarginPt,
                        TopMargin = topMarginPt,
                        CenterHorizontally = centerHorizontally,
                        CenterVertically = centerVertically,
                        PageWidthPt = Math.Round(pageWidthPt, 1),
                        PageHeightPt = Math.Round(pageHeightPt, 1),
                        PrintAreaWidthPt = Math.Round(printAreaWidth, 1),
                        PrintAreaHeightPt = Math.Round(printAreaHeight, 1),
                        Scale = dpi / 72.0,
                        ActualScaleX = Math.Round(scaleX, 6),
                        ActualScaleY = Math.Round(scaleY, 6),
                        Zoom = zoomSetting,
                        FitToPagesWide = fitToPagesWide,
                        FitToPagesTall = fitToPagesTall
                    }
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
                _logger.LogWarning("Capture operation cancelled due to timeout or client disconnect for file: {FilePath}", excelFilePath);
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
                // Cleanup: Release COM objects and delete the intermediate PDF
                var cleanupSw = Stopwatch.StartNew();

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

                cleanupTime = cleanupSw.ElapsedMilliseconds;

                // Log final performance summary
                sw.Stop();
                _logger.LogInformation(
                    "Capture completed for {FilePath} - Total: {Total}ms | Excel: {Excel}ms | Open: {Open}ms | " +
                    "PDF: {Pdf}ms | Convert: {Conv}ms | Cleanup: {Clean}ms",
                    excelFilePath,
                    sw.ElapsedMilliseconds,
                    excelStartTime,
                    workbookOpenTime - excelStartTime,
                    pdfExportTime - workbookOpenTime,
                    conversionTime - pdfExportTime,
                    cleanupTime);
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Phase 10 — Live Excel Runtime State Capture
        // ═══════════════════════════════════════════════════════════════════

        #region Phase 10 Runtime State Capture

        /// <summary>Win32 GetDeviceCaps constants.</summary>
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const int HORZRES = 8;
        private const int VERTRES = 10;
        private const int PHYSICALWIDTH = 110;
        private const int PHYSICALHEIGHT = 111;
        private const int PHYSICALOFFSETX = 112;
        private const int PHYSICALOFFSETY = 113;

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int GetDeviceCaps(IntPtr hdc, int index);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        /// <summary>
        /// Captures the complete COM runtime state at the current pipeline stage.
        /// Writes a JSON snapshot to a timestamped file and logs key differences
        /// if a previous snapshot exists.
        /// </summary>
        private void CaptureGeometryTimeline(
            string stage,
            Application? excelApp,
            Workbook? workbook,
            Worksheet? worksheet,
            string? printArea,
            List<ExcelField>? fields,
            string previewFolder)
        {
            try
            {
                var snapshot = new RuntimeStateSnapshot
                {
                    Stage = stage,
                    WorkbookName = workbook?.Name ?? "",
                    WorksheetName = worksheet?.Name ?? "",
                    WorksheetIndex = worksheet?.Index ?? 0,
                };

                if (worksheet != null)
                {
                    DumpPageSetup(snapshot, worksheet);
                    DumpWindowState(snapshot, worksheet);
                    DumpRangeGeometry(snapshot, worksheet, printArea);
                    DumpShapes(snapshot, worksheet);
                    DumpFieldCells(snapshot, worksheet, fields);
                }

                if (excelApp != null)
                {
                    DumpPrinterState(snapshot, excelApp, previewFolder);
                    snapshot.ActivePrinter = excelApp.ActivePrinter;
                    snapshot.Version = excelApp.Version;
                }

                // Auto-diff against previous snapshot
                if (_runtimeSnapshots.Count > 0)
                {
                    var prev = _runtimeSnapshots[^1];
                    LogSnapshotDiff(prev, snapshot);
                }

                _runtimeSnapshots.Add(snapshot);

                // Write JSON to disk
                string capturesDir = Path.Combine(previewFolder, RuntimeCaptureDir);
                Directory.CreateDirectory(capturesDir);
                string fileName = $"snapshot_{stage.Replace(" ", "_")}_{Guid.NewGuid():N}.json";
                string filePath = Path.Combine(capturesDir, fileName);
                var json = JsonSerializer.Serialize(snapshot, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
                });
                System.IO.File.WriteAllText(filePath, json);
                _logger.LogInformation("[PHASE10] Snapshot saved: {Path} ({JsonLen} bytes)", filePath, json.Length);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[PHASE10] Failed to capture runtime state at stage \"{Stage}\"", stage);
            }
        }

        /// <summary>Capture every PageSetup property from COM.</summary>
        private void DumpPageSetup(RuntimeStateSnapshot snap, Worksheet ws)
        {
            var ps = new PageSetupSnapshot();
            try
            {
                var pageSetup = ws.PageSetup;
                ps.LeftMargin = pageSetup.LeftMargin;
                ps.RightMargin = pageSetup.RightMargin;
                ps.TopMargin = pageSetup.TopMargin;
                ps.BottomMargin = pageSetup.BottomMargin;

                try { ps.HeaderMargin = pageSetup.HeaderMargin; } catch { }
                try { ps.FooterMargin = pageSetup.FooterMargin; } catch { }

                ps.CenterHorizontally = pageSetup.CenterHorizontally;
                ps.CenterVertically = pageSetup.CenterVertically;
                ps.PaperSize = (int)(double)pageSetup.PaperSize;
                ps.Orientation = (int)pageSetup.Orientation;

                try { ps.Zoom = pageSetup.Zoom; } catch { }
                ps.FitToPagesWide = pageSetup.FitToPagesWide;
                ps.FitToPagesTall = pageSetup.FitToPagesTall;
                ps.Draft = pageSetup.Draft;
                ps.BlackAndWhite = pageSetup.BlackAndWhite;
                ps.Order = (int)pageSetup.Order;

                ps.PrintQuality = pageSetup.PrintQuality;
                ps.FirstPageNumber = pageSetup.FirstPageNumber;
                ps.PrintTitleRows = pageSetup.PrintTitleRows;
                ps.PrintTitleColumns = pageSetup.PrintTitleColumns;
                ps.PrintArea = pageSetup.PrintArea;

                // Calculate page dimensions
                var (pw, ph) = GetPaperSizePoints((XlPaperSize)ps.PaperSize, (XlPageOrientation)ps.Orientation);
                ps.PageWidthPt = pw;
                ps.PageHeightPt = ph;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[PHASE10] PageSetup dump failed");
            }
            snap.PageSetup = ps;
        }

        /// <summary>Capture window state and view properties.</summary>
        private void DumpWindowState(RuntimeStateSnapshot snap, Worksheet ws)
        {
            try
            {
                var win = ws.Parent?.GetType().GetProperty("ActiveWindow")?.GetValue(ws.Parent) as Microsoft.Office.Interop.Excel.Window;
                if (win == null) return;

                var ws2 = new WindowStateSnapshot
                {
                    Zoom = win.Zoom,
                    View = (int)win.View,
                    ScrollRow = win.ScrollRow,
                    ScrollColumn = win.ScrollColumn,
                    
                    DisplayGridlines = win.DisplayGridlines,
                    DisplayHeadings = win.DisplayHeadings,
                    DisplayZeros = win.DisplayZeros,
                };
                try { ws2.VisibleRange = win.VisibleRange?.Address; } catch { }
                snap.WindowState = ws2;
            }
            catch
            {
                // Window access may fail in non-interactive mode
            }
        }

        /// <summary>Capture UsedRange, PrintArea, and CurrentRegion geometry.</summary>
        private void DumpRangeGeometry(RuntimeStateSnapshot snap, Worksheet ws, string? printArea)
        {
            // UsedRange
            try
            {
                Excel.Range? ur = null;
                try
                {
                    ur = ws.UsedRange;
                    snap.UsedRange = new RangeGeometry
                    {
                        Left = ur.Left,
                        Top = ur.Top,
                        Width = ur.Width,
                        Height = ur.Height,
                        ColumnsCount = ur.Columns.Count,
                        RowsCount = ur.Rows.Count,
                        Address = ur.Address,
                    };
                    snap.UsedRangeAddress = ur.Address;
                }
                finally { if (ur != null) Marshal.ReleaseComObject(ur); }
            }
            catch { }

            // PrintArea range
            if (!string.IsNullOrEmpty(printArea))
            {
                try
                {
                    Excel.Range? pr = null;
                    try
                    {
                        pr = ws.Range[printArea];
                        snap.PrintArea = new RangeGeometry
                        {
                            Left = pr.Left,
                            Top = pr.Top,
                            Width = pr.Width,
                            Height = pr.Height,
                            ColumnsCount = pr.Columns.Count,
                            RowsCount = pr.Rows.Count,
                            Address = pr.Address,
                        };
                        snap.PrintAreaAddress = printArea;
                    }
                    finally { if (pr != null) Marshal.ReleaseComObject(pr); }
                }
                catch { }
            }

            // SheetDimension from first cell
            try
            {
                Excel.Range? c1 = null;
                try
                {
                    c1 = ws.Cells[1, 1];
                    snap.SheetDimension = c1.CurrentRegion?.Address ?? "";
                }
                finally { if (c1 != null) Marshal.ReleaseComObject(c1); }
            }
            catch { }
        }

        /// <summary>Capture all printable shape bounds.</summary>
        private void DumpShapes(RuntimeStateSnapshot snap, Worksheet ws)
        {
            try
            {
                var shapes = ws.Shapes;
                if (shapes == null || shapes.Count == 0) return;

                var list = new List<ShapeSnapshot>();
                foreach (Microsoft.Office.Interop.Excel.Shape shape in shapes)
                {
                    try
                    {
                        list.Add(new ShapeSnapshot
                        {
                            Name = shape.Name,
                            Left = shape.Left,
                            Top = shape.Top,
                            Width = shape.Width,
                            Height = shape.Height,
                            Visible = shape.Visible == Microsoft.Office.Core.MsoTriState.msoTrue,
                            Type = (int)shape.Type,
                            TypeName = shape.Type.ToString(),
                            
                            ZOrderPosition = shape.ZOrderPosition,
                        });
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(shape);
                    }
                }
                snap.Shapes = list;
            }
            catch
            {
                // No shapes or access denied
            }
        }

        /// <summary>Capture per-field cell geometry at the moment of capture.</summary>
        private void DumpFieldCells(RuntimeStateSnapshot snap, Worksheet ws, List<ExcelField>? fields)
        {
            if (fields == null || fields.Count == 0) return;

            var list = new List<FieldCellSnapshot>();
            foreach (var f in fields)
            {
                if (string.IsNullOrEmpty(f.Cell)) continue;
                try
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = ws.Range[f.Cell];
                        var fc = new FieldCellSnapshot
                        {
                            CellAddress = f.Cell,
                            FieldType = f.Type,
                            CellLeft = cell.Left,
                            CellTop = cell.Top,
                            CellWidth = cell.Width,
                            CellHeight = cell.Height,
                        };

                        // MergeArea
                        try
                        {
                            fc.IsMerged = cell.MergeCells;
                            if (fc.IsMerged)
                            {
                                Excel.Range? ma = null;
                                try
                                {
                                    ma = cell.MergeArea;
                                    fc.MergeAddress = ma.Address;
                                    fc.MergeLeft = ma.Left;
                                    fc.MergeTop = ma.Top;
                                    fc.MergeWidth = ma.Width;
                                    fc.MergeHeight = ma.Height;
                                }
                                finally { if (ma != null) Marshal.ReleaseComObject(ma); }
                            }
                        }
                        catch { }

                        // EntireColumn/Row
                        try
                        {
                            Excel.Range? ec = null;
                            try { ec = cell.EntireColumn; fc.EntireColumnLeft = ec.Left; fc.EntireColumnWidth = ec.Width; }
                            finally { if (ec != null) Marshal.ReleaseComObject(ec); }
                        }
                        catch { }
                        try
                        {
                            Excel.Range? er = null;
                            try { er = cell.EntireRow; fc.EntireRowTop = er.Top; fc.EntireRowHeight = er.Height; }
                            finally { if (er != null) Marshal.ReleaseComObject(er); }
                        }
                        catch { }

                        // CurrentRegion
                        try
                        {
                            Excel.Range? cr = null;
                            try { cr = cell.CurrentRegion; fc.CurrentRegionAddress = cr.Address; }
                            finally { if (cr != null) Marshal.ReleaseComObject(cr); }
                        }
                        catch { }

                        list.Add(fc);
                    }
                    finally { if (cell != null) Marshal.ReleaseComObject(cell); }
                }
                catch { }
            }
            snap.FieldCells = list;
        }

        /// <summary>Capture printer device context via Win32 GetDeviceCaps.</summary>
        private void DumpPrinterState(RuntimeStateSnapshot snap, Application excelApp, string previewFolder)
        {
            var ps = new PrinterStateSnapshot();
            try
            {
                ps.ActivePrinter = excelApp.ActivePrinter;

                // Try to get printer DC from the active printer
                // Note: Full GetDeviceCaps requires knowing the printer's DC handle,
                // which is printer-specific. We capture the printer name for manual lookup.
                string? printerName = excelApp.ActivePrinter;
                if (!string.IsNullOrEmpty(printerName))
                {
                    ps.PrinterDriver = printerName;

                    // Log the printer name so we can manually run GetDeviceCaps
                    _logger.LogInformation(
                        "[PHASE10:PRINTER] ActivePrinter=\"{Printer}\"", printerName);
                }

                // Snapshot the initial display DC for reference DPI
                IntPtr hdc = GetDC(IntPtr.Zero);
                if (hdc != IntPtr.Zero)
                {
                    try
                    {
                        ps.LogicalPixelsX = GetDeviceCaps(hdc, LOGPIXELSX);
                        ps.LogicalPixelsY = GetDeviceCaps(hdc, LOGPIXELSY);
                        ps.HorizontalResolution = GetDeviceCaps(hdc, HORZRES);
                        ps.VerticalResolution = GetDeviceCaps(hdc, VERTRES);
                        ps.PhysicalWidth = GetDeviceCaps(hdc, PHYSICALWIDTH);
                        ps.PhysicalHeight = GetDeviceCaps(hdc, PHYSICALHEIGHT);
                        ps.PhysicalOffsetX = GetDeviceCaps(hdc, PHYSICALOFFSETX);
                        ps.PhysicalOffsetY = GetDeviceCaps(hdc, PHYSICALOFFSETY);

                        // Convert pixels to points at display DPI
                        if (ps.LogicalPixelsX > 0)
                        {
                            ps.HardMarginLeftPt = ps.PhysicalOffsetX * 72.0 / ps.LogicalPixelsX;
                        }
                        if (ps.LogicalPixelsY > 0)
                        {
                            ps.HardMarginTopPt = ps.PhysicalOffsetY * 72.0 / ps.LogicalPixelsY;
                        }

                        _logger.LogInformation(
                            "[PHASE10:DC] Display DC: DPI={DpiX}x{DpiY} " +
                            "HORZRES={HRes} VERTRES={VRes} " +
                            "PHYSICALW={PW} PHYSICALH={PH} " +
                            "OFFSET=({OffX},{OffY})px -> HardMargin=({HMLeft:F2},{HMTop:F2})pt",
                            ps.LogicalPixelsX, ps.LogicalPixelsY,
                            ps.HorizontalResolution, ps.VerticalResolution,
                            ps.PhysicalWidth, ps.PhysicalHeight,
                            ps.PhysicalOffsetX, ps.PhysicalOffsetY,
                            ps.HardMarginLeftPt, ps.HardMarginTopPt);
                    }
                    finally
                    {
                        ReleaseDC(IntPtr.Zero, hdc);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[PHASE10:DC] Printer state dump failed");
            }
            snap.Printer = ps;
        }

        /// <summary>
        /// Compares two snapshots and logs every property that changed.
        /// </summary>
        private void LogSnapshotDiff(RuntimeStateSnapshot prev, RuntimeStateSnapshot curr)
        {
            _logger.LogInformation("[PHASE10:DIFF] Comparing \"{From}\" -> \"{To}\"", prev.Stage, curr.Stage);
            bool anyChange = false;

            // PageSetup diff
            if (prev.PageSetup != null && curr.PageSetup != null)
            {
                LogIfChanged("PageSetup.LeftMargin", prev.PageSetup.LeftMargin, curr.PageSetup.LeftMargin);
                LogIfChanged("PageSetup.RightMargin", prev.PageSetup.RightMargin, curr.PageSetup.RightMargin);
                LogIfChanged("PageSetup.Zoom", prev.PageSetup.Zoom, curr.PageSetup.Zoom);
                LogIfChanged("PageSetup.FitToPagesWide", prev.PageSetup.FitToPagesWide, curr.PageSetup.FitToPagesWide);
                LogIfChanged("PageSetup.CenterHorizontally", prev.PageSetup.CenterHorizontally, curr.PageSetup.CenterHorizontally);
                ref bool anyChangeRef2 = ref anyChange;
                LogIfChanged("PageSetup.PrintArea", prev.PageSetup.PrintArea ?? "", curr.PageSetup.PrintArea ?? "");
            }

            // Range diff
            if (prev.UsedRange != null && curr.UsedRange != null)
            {
                LogIfChanged("UsedRange.Left", prev.UsedRange.Left, curr.UsedRange.Left);
                LogIfChanged("UsedRange.Width", prev.UsedRange.Width, curr.UsedRange.Width);
            }
            if (prev.PrintArea != null && curr.PrintArea != null)
            {
                LogIfChanged("PrintArea.Left", prev.PrintArea.Left, curr.PrintArea.Left);
                LogIfChanged("PrintArea.Width", prev.PrintArea.Width, curr.PrintArea.Width);
            }

            // Printer diff
            if (prev.Printer != null && curr.Printer != null)
            {
                LogIfChanged("Printer.ActivePrinter", prev.Printer.ActivePrinter ?? "", curr.Printer.ActivePrinter ?? "");
            }

            // Window diff
            if (prev.WindowState != null && curr.WindowState != null)
            {
                LogIfChanged("Window.Zoom", prev.WindowState.Zoom, curr.WindowState.Zoom);
                LogIfChanged("Window.View", prev.WindowState.View, curr.WindowState.View);
                LogIfChanged("Window.ScrollRow", prev.WindowState.ScrollRow, curr.WindowState.ScrollRow);
                LogIfChanged("Window.ScrollColumn", prev.WindowState.ScrollColumn, curr.WindowState.ScrollColumn);
            }

            // Shapes diff
            int prevShapes = prev.Shapes?.Count ?? -1;
            int currShapes = curr.Shapes?.Count ?? -1;
            if (prevShapes != currShapes)
            {
                _logger.LogInformation("[PHASE10:DIFF] Shapes.Count: {Prev} -> {Curr}", prevShapes, currShapes);
                anyChange = true;
            }

            if (!anyChange)
            {
                _logger.LogInformation("[PHASE10:DIFF] No changes detected between \"{From}\" and \"{To}\"",
                    prev.Stage, curr.Stage);
            }
        }

        private void LogIfChanged(string property, double prev, double curr)
        {
            if (Math.Abs(prev - curr) > 0.001)
            {
                _logger.LogInformation(
                    "[PHASE10:DIFF] {Prop}: {Prev:F4} -> {Curr:F4} (Δ{Delta:+#0.000;-#0.000})",
                    property, prev, curr, curr - prev);
            }
        }

        private void LogIfChanged(string property, int prev, int curr)
        {
            if (prev != curr)
            {
                _logger.LogInformation(
                    "[PHASE10:DIFF] {Prop}: {Prev} -> {Curr}",
                    property, prev, curr);
            }
        }

        private void LogIfChanged(string property, bool prev, bool curr)
        {
            if (prev != curr)
            {
                _logger.LogInformation(
                    "[PHASE10:DIFF] {Prop}: {Prev} -> {Curr}",
                    property, prev, curr);
            }
        }

        private void LogIfChanged(string property, string prev, string curr)
        {
            if (prev != curr)
            {
                _logger.LogInformation(
                    "[PHASE10:DIFF] {Prop}: \"{Prev}\" -> \"{Curr}\"",
                    property, prev, curr);
            }
        }

        #endregion

        /// <summary>
        /// Extracts form field metadata from cell comments in the worksheet.
        /// Uses Excel's SpecialCells to find ONLY cells with comments (NOT all cells).
        /// Each comment is expected to contain a field type (e.g., "Type=Text").
        ///
        /// Coordinate System:
        ///   pixel = printedOrigin + (cellPoint - printAreaOriginPoint) * (DPI / 72)
        ///
        /// Where:
        ///   - printedOrigin = where the print area content starts on the rendered page
        ///     (accounts for page margins and centering)
        ///   - printAreaOrigin = where the print area starts in the worksheet (cell Left/Top)
        ///   - DPI / 72 = points-to-pixels conversion at the rendering resolution
        /// </summary>
        private static List<ExcelField> ExtractFields(
            Excel.Worksheet worksheet,
            double printAreaLeft,
            double printAreaTop,
            double printedOriginX,
            double printedOriginY,
            double scaleX,
            double scaleY,
            double pageWidthPt,
            double pageHeightPt)
        {
            var fields = new List<ExcelField>();

            try
            {
                // Use SpecialCells to get ONLY cells that have comments attached.
                // xlCellTypeComments (-4144) returns only cells with comments/notes.
                // This is MUCH more efficient than iterating every cell in the print area.
                Excel.Range? commentedCells = null;
                try
                {
                    commentedCells = worksheet.Cells.SpecialCells(
                        XlCellType.xlCellTypeComments);
                }
                catch (COMException)
                {
                    // No cells have comments — this is expected, not an error.
                    return fields;
                }

                if (commentedCells == null)
                    return fields;

                // Iterate only over the cells that have comments (typically 1-10)
                foreach (Excel.Range cell in commentedCells)
                {
                    try
                    {
                        if (cell.Comment == null)
                            continue;

                        string commentText = cell.Comment.Text();

                        if (string.IsNullOrWhiteSpace(commentText))
                            continue;

                        string cellAddress = cell.AddressLocal[false, false];
                        string fieldType = ParseFieldType(commentText);

                        // --- Determine geometry source: MergeArea or individual cell ---
                        // When a cell is part of a merged range, cell.Left/Top/Width/Height
                        // only reflects the anchor cell, not the entire merged region.
                        // We must use MergeArea to get the full merged region dimensions.
                        bool isMerged = false;
                        string? mergeAddress = null;
                        double cellLeftPt, cellTopPt, cellWidthPt, cellHeightPt;

                        try
                        {
                            isMerged = cell.MergeCells;
                        }
                        catch
                        {
                            isMerged = false;
                        }

                        if (isMerged)
                        {
                            // Use MergeArea for the complete merged region geometry
                            Excel.Range? mergeArea = null;
                            try
                            {
                                mergeArea = cell.MergeArea;
                                cellLeftPt = mergeArea.Left;
                                cellTopPt = mergeArea.Top;
                                cellWidthPt = mergeArea.Width;
                                cellHeightPt = mergeArea.Height;
                                mergeAddress = mergeArea.AddressLocal[false, false];
                            }
                            finally
                            {
                                if (mergeArea != null)
                                    Marshal.ReleaseComObject(mergeArea);
                            }
                        }
                        else
                        {
                            // Normal (non-merged) cell — use cell's own geometry
                            cellLeftPt = cell.Left;
                            cellTopPt = cell.Top;
                            cellWidthPt = cell.Width;
                            cellHeightPt = cell.Height;
                        }

                        // Calculate content offset within the print area (in points)
                        double offsetLeftPt = cellLeftPt - printAreaLeft;
                        double offsetTopPt = cellTopPt - printAreaTop;

                        // Convert to PNG pixels on the printed page:
                        // pixel = printedOrigin + (cellPoint - printAreaOriginPoint) * scale
                        double pixelLeft = scaleX > 0
                            ? printedOriginX + offsetLeftPt * scaleX
                            : 0;
                        double pixelTop = scaleY > 0
                            ? printedOriginY + offsetTopPt * scaleY
                            : 0;
                        double pixelWidth = scaleX > 0
                            ? cellWidthPt * scaleX
                            : 0;
                        double pixelHeight = scaleY > 0
                            ? cellHeightPt * scaleY
                            : 0;

                        fields.Add(new ExcelField
                        {
                            Id = $"field_{cellAddress}",
                            Cell = cellAddress,
                            Type = fieldType,
                            Left = Math.Round(pixelLeft, 1),
                            Top = Math.Round(pixelTop, 1),
                            Width = Math.Round(pixelWidth, 1),
                            Height = Math.Round(pixelHeight, 1),
                            Comment = commentText,
                            IsMerged = isMerged,
                            MergeAddress = mergeAddress,

                            // Debug metadata for verifying alignment
                            ExcelLeft = cellLeftPt,
                            ExcelTop = cellTopPt,
                            PrintAreaLeft = printAreaLeft,
                            PrintAreaTop = printAreaTop,
                            ExcelWidthPt = cellWidthPt,
                            ExcelHeightPt = cellHeightPt
                        });
                    }
                    finally
                    {
                        // Release each cell's COM range immediately to prevent leaks
                        if (cell != null)
                            Marshal.ReleaseComObject(cell);
                    }
                }

                // Release the commentedCells range
                if (commentedCells != null)
                    Marshal.ReleaseComObject(commentedCells);
            }
            catch (Exception ex) when (ex is not PrintAreaNotConfiguredException)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Warning: Field extraction failed: {ex.Message}");
            }

            return fields;
        }

        /// <summary>
        /// Parses the field type from a cell comment.
        /// Expected format: "Type=Text" or any string containing "Type=X".
        /// Returns "Unknown" if no type can be determined.
        /// </summary>
        private static string ParseFieldType(string comment)
        {
            // Look for "Type=" pattern in the comment
            const string prefix = "Type=";
            int idx = comment.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
            {
                int start = idx + prefix.Length;
                int end = comment.IndexOfAny(['\r', '\n', ' ', ','], start);

                if (end < 0)
                {
                    end = comment.Length;
                }

                if (end > start)
                {
                    return comment[start..end].Trim();
                }
            }

            return "Unknown";
        }

        /// <summary>
        /// Converts an Excel paper size and orientation to page dimensions in points.
        /// 1 point = 1/72 inch.
        /// </summary>
        private static (double width, double height) GetPaperSizePoints(XlPaperSize paperSize, XlPageOrientation orientation)
        {
            (double w, double h) = paperSize switch
            {
                XlPaperSize.xlPaperLetter or XlPaperSize.xlPaperLetterSmall => (612.0, 792.0),       // 8.5 x 11 in
                XlPaperSize.xlPaperLegal => (612.0, 1008.0),       // 8.5 x 14 in
                XlPaperSize.xlPaperA4 or XlPaperSize.xlPaperA4Small => (595.0, 842.0),  // 210 x 297 mm
                XlPaperSize.xlPaperA3 => (842.0, 1191.0),          // 297 x 420 mm
                XlPaperSize.xlPaperA5 => (420.0, 595.0),           // 148 x 210 mm
                XlPaperSize.xlPaperB5 => (499.0, 709.0),           // 176 x 250 mm
                XlPaperSize.xlPaperExecutive => (522.0, 756.0),    // 7.25 x 10.5 in
                XlPaperSize.xlPaperTabloid => (792.0, 1224.0),     // 11 x 17 in
                XlPaperSize.xlPaperLedger => (1224.0, 792.0),      // 17 x 11 in
                XlPaperSize.xlPaperEnvelope10 => (684.0, 360.0),   // 9.5 x 4.125 in
                _ => (612.0, 792.0)  // Default to Letter
            };

            if (orientation == XlPageOrientation.xlLandscape)
                return (h, w);

            return (w, h);
        }

        /// <summary>
        /// Releases a single COM object safely.
        /// </summary>
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

        /// <summary>
        /// Releases all COM objects and ensures Excel process is terminated.
        /// This is critical to prevent orphan EXCEL.EXE processes.
        /// </summary>
        private void CleanupComObjects(Workbook? workbook, Application? excelApp)
        {
            // Close and release the workbook
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

            // Quit and release the Excel application
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

            // Force garbage collection to release any remaining COM wrappers
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #region Coordinate Helpers

        /// <summary>
        /// Computes the print area content width from XLSX worksheet column definitions.
        /// For worksheets with explicit &lt;cols&gt; elements, sums the column widths.
        /// For default-column worksheets, uses 50.1pt per column (Calibri 11pt legacy width).
        /// Returns 0 if the XLSX cannot be read (caller falls back to Range.Width).
        /// </summary>
        private double ComputeContentWidthFromXlsx(
            string excelFilePath, string? worksheetName, int worksheetIndex, string printArea, int printAreaCols)
        {
            _logger.LogDebug("[COLS] Computing content width from XLSX: Sheet=\"{Sheet}\" Index={Index} PA=\"{PA}\" Cols={Cols}",
                worksheetName ?? "(null)", worksheetIndex, printArea, printAreaCols);

            // Parse print area address to determine column range
            if (!TryParsePrintAreaColumns(printArea, out int firstCol, out int lastCol))
            {
                firstCol = 1;
                lastCol = printAreaCols;
            }

            int colsInRange = lastCol - firstCol + 1;
            _logger.LogDebug("[COLS] Column range: {First} to {Last} ({Count} cols)", firstCol, lastCol, colsInRange);

            try
            {
                using var archive = new ZipArchive(
                    System.IO.File.OpenRead(excelFilePath), ZipArchiveMode.Read);

                // Find the worksheet XML entry matching the COM-selected worksheet
                // Uses worksheet Index (sheetId) as primary resolution key,
                // falling back to name-based matching for robustness.
                string? wsPath = ResolveWorksheetPath(archive, worksheetName, worksheetIndex);
                if (wsPath == null)
                {
                    _logger.LogWarning("[COLS] Could not resolve worksheet path for \"{Sheet}\"", worksheetName);
                    return 0;
                }

                var wsEntry = archive.GetEntry(wsPath);
                if (wsEntry == null)
                {
                    _logger.LogWarning("[COLS] Worksheet entry not found: {Path}", wsPath);
                    return 0;
                }

                using var reader = new StreamReader(wsEntry.Open());
                var wsXml = XDocument.Load(reader);
                XNamespace ns = wsXml.Root!.GetDefaultNamespace();

                // Check for explicit &lt;cols&gt; element
                var colsElement = wsXml.Descendants(ns + "cols").FirstOrDefault();

                if (colsElement != null)
                {
                    // Explicit column widths defined in XLSX
                    double totalWidth = 0;
                    int colsFound = 0;

                    foreach (var col in colsElement.Elements(ns + "col"))
                    {
                        int min = (int)col.Attribute("min")!;
                        int max = (int)col.Attribute("max")!;
                        double width = (double)col.Attribute("width")!;
                        bool hidden = (bool?)col.Attribute("hidden") ?? false;

                        // Only include columns within the print area range
                        int overlapMin = Math.Max(min, firstCol);
                        int overlapMax = Math.Min(max, lastCol);

                        if (overlapMin <= overlapMax && !hidden)
                        {
                            int count = overlapMax - overlapMin + 1;

                            // Convert from character units to points using the OOXML formula:
                            //   pixelWidth = charWidth * maxDigitWidth + padding
                            //   pointWidth = pixelWidth * 72 / 96
                            // maxDigitWidth ≈ 7.33 for Calibri 11pt (legacy Normal font)
                            // padding = 5 pixels (standard Excel constant at 96 DPI)
                            // Reference: ECMA-376, 18.3.1.13 (col element)
                            const double maxDigitWidth = 7.33;
                            const double padding = 5.0;
                            double pointWidth = (width * maxDigitWidth + padding) * 72.0 / 96.0;

                            totalWidth += pointWidth * count;
                            colsFound += count;

                            _logger.LogDebug(
                                "[COLS] Col {Min}-{Max}: charWidth={W:F2} -> {PW:F2}pt custom={C} hidden={H} -> +{Count} col(s), running total={Total:F2}pt",
                                min, max, width, pointWidth,
                                (bool?)col.Attribute("customWidth") ?? false,
                                hidden, count, totalWidth);
                        }
                    }

                    if (colsFound > 0)
                    {
                        _logger.LogInformation(
                            "[COLS] Explicit column widths: {Found} cols in range {First}-{Last}, total={Total:F2}pt",
                            colsFound, firstCol, lastCol, totalWidth);
                        return totalWidth;
                    }

                    _logger.LogInformation(
                        "[COLS] &lt;cols&gt; element found but no columns matched print area range {First}-{Last}, falling back to default",
                        firstCol, lastCol);
                }
                else
                {
                    _logger.LogDebug("[COLS] No &lt;cols&gt; element found in worksheet XML");
                }

                // No explicit columns or none in range — use default column width
                double defaultWidth = colsInRange * 50.1;
                _logger.LogInformation(
                    "[COLS] Default column width: {Cols} cols x 50.1pt = {Total:F2}pt (Calibri 11pt legacy width)",
                    colsInRange, defaultWidth);
                return defaultWidth;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[COLS] Failed to read XLSX column widths, falling back to Range.Width");
                return 0;
            }
        }

        /// <summary>
        /// Resolves a worksheet name to its XLSX internal path (e.g., "xl/worksheets/sheet1.xml").
        /// Uses the COM worksheet Index (sheetId) as the primary resolution key,
        /// falling back to name-based matching for robustness.
        /// Reads workbook.xml and relationships to map sheet names/IDs to files.
        /// </summary>
        private static string? ResolveWorksheetPath(ZipArchive archive, string? worksheetName, int worksheetIndex)
        {
            // =============================
            // WORKSHEET RESOLUTION LOGGING
            // =============================
            System.Diagnostics.Debug.WriteLine("\n=========================");
            System.Diagnostics.Debug.WriteLine("WORKSHEET RESOLUTION");
            System.Diagnostics.Debug.WriteLine("=========================");
            System.Diagnostics.Debug.WriteLine($"COM Name:   {worksheetName ?? "(null)"}");
            System.Diagnostics.Debug.WriteLine($"COM Index:  {worksheetIndex}");

            var wbEntry = archive.GetEntry("xl/workbook.xml");
            if (wbEntry == null)
            {
                System.Diagnostics.Debug.WriteLine("FAILED: xl/workbook.xml not found in archive");
                return null;
            }

            using var wbReader = new StreamReader(wbEntry.Open());
            var wbXml = XDocument.Load(wbReader);
            XNamespace wbNs = wbXml.Root!.GetDefaultNamespace();

            // Collect all sheet definitions from workbook.xml
            var sheetEntries = new List<(string name, string? sheetId, string? relId)>();
            foreach (var sheet in wbXml.Descendants(wbNs + "sheet"))
            {
                string? name = sheet.Attribute("name")?.Value;
                string? sid = sheet.Attribute("sheetId")?.Value;
                string? rid = sheet.Attribute(
                    XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))?.Value;
                sheetEntries.Add((name ?? "?", sid, rid));
            }

            // Log workbook.xml sheets
            System.Diagnostics.Debug.WriteLine("\nWorkbook.xml Sheets:");
            foreach (var (name, sid, rid) in sheetEntries)
            {
                System.Diagnostics.Debug.WriteLine($"  Sheet=\"{name}\" sheetId={sid} {rid}");
            }

            // Read workbook.xml.rels to map relationship IDs to target paths
            var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry == null)
            {
                System.Diagnostics.Debug.WriteLine("FAILED: xl/_rels/workbook.xml.rels not found");
                return null;
            }

            using var relsReader = new StreamReader(relsEntry.Open());
            var relsXml = XDocument.Load(relsReader);
            XNamespace relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var relsMap = new Dictionary<string, string>();
            foreach (var rel in relsXml.Descendants(relsNs + "Relationship"))
            {
                string? id = rel.Attribute("Id")?.Value;
                string? target = rel.Attribute("Target")?.Value;
                if (id != null && target != null)
                {
                    relsMap[id] = target.Replace('\\', '/');
                }
            }

            // Log relationships
            System.Diagnostics.Debug.WriteLine("\nRelationships:");
            foreach (var (rid, target) in relsMap)
            {
                if (target.StartsWith("worksheets/", StringComparison.OrdinalIgnoreCase) ||
                    target.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase))
                {
                    System.Diagnostics.Debug.WriteLine($"  {rid}  {target}");
                }
            }

            // Strategy 1: Match by sheetId == worksheetIndex (primary)
            string? targetIndex = worksheetIndex.ToString();
            string? matchedRelId = null;
            foreach (var (name, sid, rid) in sheetEntries)
            {
                if (sid == targetIndex && rid != null)
                {
                    matchedRelId = rid;
                    System.Diagnostics.Debug.WriteLine($"\nStrategy 1 (Index match): sheetId={sid} -> {rid}");
                    break;
                }
            }

            // Strategy 2: Fallback to exact name match
            if (matchedRelId == null && !string.IsNullOrEmpty(worksheetName))
            {
                foreach (var (name, _, rid) in sheetEntries)
                {
                    if (name == worksheetName && rid != null)
                    {
                        matchedRelId = rid;
                        System.Diagnostics.Debug.WriteLine($"\nStrategy 2 (Name match): \"{name}\" -> {rid}");
                        break;
                    }
                }
            }

            // Strategy 3: Fallback to case-insensitive name match
            if (matchedRelId == null && !string.IsNullOrEmpty(worksheetName))
            {
                foreach (var (name, _, rid) in sheetEntries)
                {
                    if (name != null && name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase) && rid != null)
                    {
                        matchedRelId = rid;
                        System.Diagnostics.Debug.WriteLine($"\nStrategy 3 (CI name match): \"{name}\" -> {rid}");
                        break;
                    }
                }
            }

            if (matchedRelId == null)
            {
                System.Diagnostics.Debug.WriteLine("FAILED: No matching sheet found in workbook.xml");
                return null;
            }

            // Resolve the relationship ID to a target path
            if (relsMap.TryGetValue(matchedRelId, out string? targetPath) && targetPath != null)
            {
                string resolved = targetPath.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                    ? targetPath
                    : "xl/" + targetPath;

                System.Diagnostics.Debug.WriteLine($"\nResolved XML: {resolved}");
                System.Diagnostics.Debug.WriteLine("SUCCESS");
                return resolved;
            }

            System.Diagnostics.Debug.WriteLine($"FAILED: Relationship {matchedRelId} not found in workbook.xml.rels");
            return null;
        }

        /// <summary>
        /// Converts a column letter (e.g., "A", "Z", "AA") to a 1-based column index.
        /// </summary>
        private static int ColumnLetterToIndex(string letters)
        {
            int result = 0;
            foreach (char c in letters.ToUpperInvariant())
            {
                result = result * 26 + (c - 'A' + 1);
            }
            return result;
        }

        /// <summary>
        /// Parses a print area address like "$A$1:$D$10" to extract column indices (1-based).
        /// Returns false if the address cannot be parsed.
        /// </summary>
        private static bool TryParsePrintAreaColumns(string printArea, out int firstCol, out int lastCol)
        {
            firstCol = 0;
            lastCol = 0;

            var match = Regex.Match(printArea, @"\$?([A-Z]+)\$?\d+:\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase);
            if (!match.Success || match.Groups.Count < 3)
                return false;

            firstCol = ColumnLetterToIndex(match.Groups[1].Value);
            lastCol = ColumnLetterToIndex(match.Groups[2].Value);
            return true;
        }

        #endregion
    }
}
