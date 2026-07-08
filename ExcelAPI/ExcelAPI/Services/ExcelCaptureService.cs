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
    ///
    /// This approach ensures correct margins, scaling, fonts, borders, and page breaks.
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
        public async Task<string> CapturePrintAreaAsync(string excelFilePath, CancellationToken cancellationToken = default)
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

                // --- Step 5: Select the first worksheet ---
                _logger.LogDebug("Selecting first worksheet...");
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // --- Step 6: Read the configured Print Area ---
                string? printArea = worksheet.PageSetup.PrintArea;
                _logger.LogInformation("Print area detected: {PrintArea}", printArea ?? "(none)");

                if (string.IsNullOrEmpty(printArea))
                {
                    throw new PrintAreaNotConfiguredException(
                        "No print area is configured in this worksheet. " +
                        "Please set a print area (Page Layout > Print Area > Set Print Area) " +
                        "in the Excel file before uploading.");
                }

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

                // Return the relative URL path
                return $"/preview/{previewFileName}";
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
                // Re-throw user-domain exceptions as-is for the controller to handle
                throw;
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("Capture operation cancelled due to timeout or client disconnect for file: {FilePath}", excelFilePath);
                throw;
            }
            catch (InvalidOperationException)
            {
                // Re-throw domain exceptions as-is
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
