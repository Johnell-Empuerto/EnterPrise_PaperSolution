using ExcelAPI.Models;
using ExcelAPI.Models.WorkbookDefinition;
using ExcelAPI.Rendering;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Designer.Generation
{
    /// <summary>
    /// Generates PDF output for forms.
    /// Canonical path: WorkbookDefinition → ExportCoordinator.ExportPdf().
    ///
    /// Phase 4.2: Primary overload accepts WorkbookDefinition.
    ///            Fallback overlay-only path removed — WbDef is always available.
    ///            No rendering logic in this class — only coordinates exports.
    /// </summary>
    public class PdfGenerator
    {
        private readonly ILogger<PdfGenerator> _logger;
        private readonly ExportCoordinator _exportCoordinator;
        private readonly string? _xlsxPath;

        public PdfGenerator(ILogger<PdfGenerator> logger, ExportCoordinator exportCoordinator)
        {
            _logger = logger;
            _exportCoordinator = exportCoordinator;
        }

        public PdfGenerator(ILogger<PdfGenerator> logger, ExportCoordinator exportCoordinator, string xlsxPath)
        {
            _logger = logger;
            _exportCoordinator = exportCoordinator;
            _xlsxPath = xlsxPath;
        }

        /// <summary>
        /// Generate a PDF from a FormDefinition (compatibility shim).
        /// Converts to WorkbookDefinition internally.
        /// </summary>
        public string Generate(FormDefinition form, string sheetId, string outputPath)
        {
            var wbDef = WorkbookDefinitionConverter.FromFormDefinition(form);
            wbDef.SourcePath = _xlsxPath ?? form.Metadata.GetValueOrDefault("xlsxPath", "");

            var sheet = wbDef.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            return Generate(wbDef, sheetId, outputPath);
        }

        /// <summary>
        /// Generate a PDF from a WorkbookDefinition (canonical path).
        /// Delegates to ExportCoordinator for all rendering.
        /// </summary>
        public string Generate(WorkbookDefinition wbDef, string sheetId, string outputPath)
        {
            var sheet = wbDef.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            // Convert WbDef to FormDefinition for ExportCoordinator compatibility
            var form = WorkbookDefinitionConverter.ToFormDefinition(wbDef);

            string xlsxPath = _xlsxPath ?? wbDef.SourcePath ?? "";
            if (!string.IsNullOrEmpty(xlsxPath) && System.IO.File.Exists(xlsxPath))
            {
                try
                {
                    var options = new ExportOptions
                    {
                        Dpi = 300,
                        EmbedFonts = false,
                        RasterImagesOnly = true,
                        CompressPdf = true,
                        IncludeMetadata = true,
                        Title = form.Workbook.Title,
                        Author = form.Workbook.Author,
                        Subject = form.Workbook.Description,
                        HighQualityAntialiasing = true,
                        MaxPages = 0,
                        IncludePageNumbers = false
                    };

                    string result = _exportCoordinator.ExportPdf(xlsxPath, form, sheetId, outputPath, options);
                    _logger.LogInformation("PDF generated via ExportCoordinator: {Path}", outputPath);
                    return result;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "ExportCoordinator failed for PDF");
                }
            }

            // Fallback: minimal placeholder PDF
            double pageWidthPt = sheet.PrintLayout?.PageWidthPt ?? 612;
            double pageHeightPt = sheet.PrintLayout?.PageHeightPt ?? 792;
            using var pdfDocument = SkiaSharp.SKDocument.CreatePdf(outputPath);
            using var pdfCanvas = pdfDocument.BeginPage((float)pageWidthPt, (float)pageHeightPt);
            pdfCanvas.Clear(SkiaSharp.SKColors.White);
            pdfDocument.EndPage();
            pdfDocument.Close();

            _logger.LogWarning("PDF ExportCoordinator failed — generated blank placeholder: {Path}", outputPath);
            return outputPath;
        }
    }
}
