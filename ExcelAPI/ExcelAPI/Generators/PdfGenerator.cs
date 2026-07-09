using ExcelAPI.Models;
using Microsoft.Extensions.Logging;
using SkiaSharp;

namespace ExcelAPI.Generators
{
    public class PdfGenerator
    {
        private readonly ILogger<PdfGenerator> _logger;
        private const double DPI = 300.0;
        private const double PointsToPixels = DPI / 72.0;

        public PdfGenerator(ILogger<PdfGenerator> logger)
        {
            _logger = logger;
        }

        public string Generate(FormDefinition form, string sheetId, string outputPath)
        {
            var sheet = form.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            if (sheet.PrintArea == null)
                throw new InvalidOperationException("No print area configured for this sheet.");

            double pageWidthPt = sheet.PageSettings.WidthPt;
            double pageHeightPt = sheet.PageSettings.HeightPt;

            int width = (int)Math.Round(pageWidthPt * PointsToPixels);
            int height = (int)Math.Round(pageHeightPt * PointsToPixels);

            using var bitmap = new SKBitmap(width, height);
            using var canvas = new SKCanvas(bitmap);

            canvas.Clear(SKColors.White);

            double printAreaLeftPx = sheet.PrintArea.LeftPt * PointsToPixels;
            double printAreaTopPx = sheet.PrintArea.TopPt * PointsToPixels;
            double printAreaWidthPx = sheet.PrintArea.WidthPt * PointsToPixels;
            double printAreaHeightPx = sheet.PrintArea.HeightPt * PointsToPixels;

            DrawPrintAreaContent(canvas, sheet, printAreaLeftPx, printAreaTopPx);

            // Draw cluster overlays
            var sheetClusters = form.Clusters.Where(c => c.SheetId == sheetId).ToList();
            foreach (var cluster in sheetClusters)
            {
                DrawCluster(canvas, cluster);
            }

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);

            // Save as PNG first, then we could convert to PDF
            string pngPath = outputPath.Replace(".pdf", ".png");
            System.IO.File.WriteAllBytes(pngPath, data.ToArray());

            // Use PDFtoImage-like approach or SkiaSharp PDF
            using var pdfBitmap = SKBitmap.Decode(pngPath);
            if (pdfBitmap == null)
                throw new InvalidOperationException("Failed to decode preview for PDF generation");

            using var pdfDocument = SkiaSharp.SKDocument.CreatePdf(outputPath);
            using var pdfCanvas = pdfDocument.BeginPage((float)pageWidthPt, (float)pageHeightPt);

            pdfCanvas.DrawBitmap(pdfBitmap, 0, 0);

            pdfDocument.EndPage();
            pdfDocument.Close();

            _logger.LogInformation("PDF generated: {Path}", outputPath);
            return outputPath;
        }

        private static void DrawPrintAreaContent(SKCanvas canvas, SheetDefinition sheet,
            double leftPx, double topPx)
        {
            using var gridPaint = new SKPaint
            {
                Color = new SKColor(220, 220, 220, 100),
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 0.5f
            };

            double x = leftPx;
            double y = topPx;

            int maxRow = 0;
            if (sheet.RowHeights.Count > 0)
                maxRow = sheet.RowHeights.Keys.Max();
            if (sheet.PrintArea != null)
                maxRow = Math.Max(maxRow, sheet.PrintArea.Rows);

            int maxCol = 0;
            if (sheet.ColumnWidths.Count > 0)
                maxCol = sheet.ColumnWidths.Keys.Max();
            if (sheet.PrintArea != null)
                maxCol = Math.Max(maxCol, sheet.PrintArea.Cols);

            for (int row = 1; row <= maxRow; row++)
            {
                double h = sheet.RowHeights.TryGetValue(row, out var rh) ? rh : 15;
                h *= PointsToPixels;
                canvas.DrawLine((float)x, (float)(y + h), (float)(x + 1000), (float)(y + h), gridPaint);
                y += h;
            }
        }

        private static void DrawCluster(SKCanvas canvas, ClusterDefinition cluster)
        {
            float left = (float)(cluster.LeftPt * PointsToPixels);
            float top = (float)(cluster.TopPt * PointsToPixels);
            float width = (float)(cluster.WidthPt * PointsToPixels);
            float height = (float)(cluster.HeightPt * PointsToPixels);

            using var fillPaint = new SKPaint
            {
                Color = new SKColor(200, 200, 255, 30),
                Style = SKPaintStyle.Fill
            };

            using var borderPaint = new SKPaint
            {
                Color = new SKColor(100, 100, 200, 150),
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 1
            };

            canvas.DrawRect(left, top, width, height, fillPaint);
            canvas.DrawRect(left, top, width, height, borderPaint);
        }
    }
}
