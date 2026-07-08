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
    /// Scale Factor: 1 Excel point = 1/72 inch. PNG at 300 DPI.
    /// PNG pixel = Excel point * (300 / 72) = Excel point * 4.1667
    /// </summary>
    public class ExcelCaptureService : IExcelCaptureService
    {
        // Scale factor from Excel points to PNG pixels at 300 DPI
        // 1 inch = 72 points (Excel) = 300 pixels (PNG)
        // So: PNG pixels = Excel points * (300 / 72)
        private const double PointsToPixels = 300.0 / 72.0;

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

                // --- Step 9: Extract field metadata from cell comments ---
                _logger.LogDebug("Extracting field metadata from cell comments...");
                var fields = ExtractFields(worksheet, printArea);
                _logger.LogInformation("Extracted {Count} field(s) from cell comments", fields.Count);

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

                // Build and return the complete capture result
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
        /// Coordinates are converted from Excel points to PNG pixels at 300 DPI.
        /// </summary>
        private static List<ExcelField> ExtractFields(Excel.Worksheet worksheet, string printArea)
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

                        double leftPt = cell.Left;
                        double topPt = cell.Top;
                        double widthPt = cell.Width;
                        double heightPt = cell.Height;

                        fields.Add(new ExcelField
                        {
                            Id = $"field_{cellAddress}",
                            Cell = cellAddress,
                            Type = fieldType,
                            Left = Math.Round(leftPt * PointsToPixels, 1),
                            Top = Math.Round(topPt * PointsToPixels, 1),
                            Width = Math.Round(widthPt * PointsToPixels, 1),
                            Height = Math.Round(heightPt * PointsToPixels, 1),
                            Comment = commentText
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
