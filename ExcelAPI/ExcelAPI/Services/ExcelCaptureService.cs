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
        public async Task<CaptureResult> CapturePrintAreaAsync(string excelFilePath, CancellationToken cancellationToken = default)
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
                string fileId = Guid.NewGuid().ToString("N");
                string pdfFileName = $"page_{fileId}.pdf";
                string previewFileName = $"page_{fileId}.png";
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

                // --- Step 11: Verify the point-to-pixel scale ---
                // The scale from Excel points to PNG pixels at the rendering DPI is:
                //   scale = DPI / 72 ≈ 4.1667
                // This was computed in Step 6b as `pointsToPixels`.
                //
                // Additionally, the PNG shows the FULL PAGE (including margins).
                // The PRINTED ORIGIN offset (computed in Step 6b) shifts the
                // print area content from the raw worksheet origin to where it
                // actually appears on the page.
                //
                // Final formula:
                //   pixel = printedOrigin + (cellPoint - printAreaOriginPoint) * (DPI / 72)

                double scaleX = dpi / 72.0;
                double scaleY = dpi / 72.0;

                _logger.LogInformation(
                    "[AUDIT] Final coordinate system:\n" +
                    "  Printed origin:         ({OX:F1},{OY:F1}) px\n" +
                    "  Point-to-pixel scale:   {Scale:F4} px/pt (at {DPI} DPI)\n" +
                    "  Formula: pixel = printedOrigin + (cellPoint - printAreaOrigin) * scale",
                    printedOriginX, printedOriginY, scaleX, dpi);

                // --- Step 12: Extract field metadata from cell comments ---
                // Coordinates now include the printed page offset.
                _logger.LogDebug("Extracting field metadata from cell comments...");
                var fields = ExtractFields(
                    worksheet,
                    printAreaLeft, printAreaTop,
                    printedOriginX, printedOriginY,
                    scaleX, scaleY);
                _logger.LogInformation("Extracted {Count} field(s) from cell comments", fields.Count);

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
                        printedOriginX, printedOriginY,
                        f.ExcelLeft, f.ExcelTop,
                        f.PrintAreaLeft, f.PrintAreaTop,
                        scaleX,
                        f.Left, f.Top,
                        f.Width, f.Height);
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
                        PrintedOriginX = Math.Round(printedOriginX, 1),
                        PrintedOriginY = Math.Round(printedOriginY, 1),
                        LeftMargin = leftMarginPt,
                        TopMargin = topMarginPt,
                        CenterHorizontally = centerHorizontally,
                        CenterVertically = centerVertically,
                        PageWidthPt = Math.Round(pageWidthPt, 1),
                        PageHeightPt = Math.Round(pageHeightPt, 1),
                        PrintAreaWidthPt = Math.Round(printAreaWidth, 1),
                        PrintAreaHeightPt = Math.Round(printAreaHeight, 1),
                        Scale = dpi / 72.0,
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
            double scaleY)
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

                        // Raw cell coordinates in Excel points
                        double cellLeftPt = cell.Left;
                        double cellTopPt = cell.Top;
                        double cellWidthPt = cell.Width;
                        double cellHeightPt = cell.Height;

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

                            // Debug metadata for verifying alignment
                            ExcelLeft = cellLeftPt,
                            ExcelTop = cellTopPt,
                            PrintAreaLeft = printAreaLeft,
                            PrintAreaTop = printAreaTop
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
    }
}
