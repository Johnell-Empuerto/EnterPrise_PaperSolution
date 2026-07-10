using ExcelAPI.Models;
using ExcelAPI.Rendering;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Generators
{
    /// <summary>
    /// Generates PDF output for forms.
    /// Delegates all rendering to ExportCoordinator (Phase 11G Production Export Engine).
    ///
    /// No rendering logic should exist in this class — it only coordinates
    /// the export options and delegates to ExportCoordinator.ExportPdf().
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
        /// Generate a PDF for the given form and sheet.
        /// Delegates to ExportCoordinator.ExportPdf() with default PDF options.
        /// Falls back to overlay-only PDF if XLSX is unavailable or rendering fails.
        /// </summary>
        public string Generate(FormDefinition form, string sheetId, string outputPath)
        {
            var sheet = form.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            string xlsxPath = _xlsxPath ?? form.Metadata.GetValueOrDefault("xlsxPath", "");
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
                        MaxPages = 0, // unlimited
                        IncludePageNumbers = false
                    };

                    string result = _exportCoordinator.ExportPdf(xlsxPath, form, sheetId, outputPath, options);
                    _logger.LogInformation("PDF generated via ExportCoordinator: {Path}", outputPath);
                    return result;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "ExportCoordinator failed, falling back to overlay-only PDF");
                }
            }

            // Fallback: overlay-only PDF
            return GenerateOverlayOnly(form, sheetId, sheet, outputPath);
        }

        private string GenerateOverlayOnly(FormDefinition form, string sheetId, SheetDefinition sheet, string outputPath)
        {
            const double PointsToPixels = 300.0 / 72.0;

            double pageWidthPt = sheet.PageSettings.WidthPt;
            double pageHeightPt = sheet.PageSettings.HeightPt;

            int width = (int)Math.Round(pageWidthPt * PointsToPixels);
            int height = (int)Math.Round(pageHeightPt * PointsToPixels);

            using var bitmap = new SkiaSharp.SKBitmap(width, height);
            using var canvas = new SkiaSharp.SKCanvas(bitmap);

            canvas.Clear(SkiaSharp.SKColors.White);

            // Draw cluster overlays
            var sheetClusters = form.Clusters.Where(c => c.SheetId == sheetId).ToList();
            foreach (var cluster in sheetClusters)
            {
                DrawCluster(canvas, cluster, PointsToPixels);
            }

            // Create PDF from rendered bitmap
            using var pdfDocument = SkiaSharp.SKDocument.CreatePdf(outputPath);
            using var pdfCanvas = pdfDocument.BeginPage((float)pageWidthPt, (float)pageHeightPt);
            pdfCanvas.DrawBitmap(bitmap, 0, 0);
            pdfDocument.EndPage();
            pdfDocument.Close();

            _logger.LogInformation("Overlay-only PDF generated: {Path}", outputPath);
            return outputPath;
        }

        private static void DrawCluster(SkiaSharp.SKCanvas canvas, ClusterDefinition cluster, double pointsToPixels)
        {
            float left = (float)(cluster.LeftPt * pointsToPixels);
            float top = (float)(cluster.TopPt * pointsToPixels);
            float width = (float)(cluster.WidthPt * pointsToPixels);
            float height = (float)(cluster.HeightPt * pointsToPixels);

            using var fillPaint = new SkiaSharp.SKPaint
            {
                Color = new SkiaSharp.SKColor(200, 200, 255, 30),
                Style = SkiaSharp.SKPaintStyle.Fill
            };

            using var borderPaint = new SkiaSharp.SKPaint
            {
                Color = new SkiaSharp.SKColor(100, 100, 200, 150),
                Style = SkiaSharp.SKPaintStyle.Stroke,
                StrokeWidth = 1
            };

            canvas.DrawRect(left, top, width, height, fillPaint);
            canvas.DrawRect(left, top, width, height, borderPaint);
        }
    }
}
