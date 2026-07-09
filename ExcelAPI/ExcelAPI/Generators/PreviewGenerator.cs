using ExcelAPI.Models;
using SkiaSharp;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Generators
{
    public class PreviewGenerator
    {
        private readonly ILogger<PreviewGenerator> _logger;
        private const double DPI = 300.0;
        private const double PointsToPixels = DPI / 72.0;

        public PreviewGenerator(ILogger<PreviewGenerator> logger)
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

            double pageWidthPx = sheet.PageSettings.WidthPt * PointsToPixels;
            double pageHeightPx = sheet.PageSettings.HeightPt * PointsToPixels;

            int width = (int)Math.Round(pageWidthPx);
            int height = (int)Math.Round(pageHeightPx);

            using var bitmap = new SKBitmap(width, height);
            using var canvas = new SKCanvas(bitmap);

            canvas.Clear(SKColors.White);

            // Draw content within print area
            double printAreaLeftPx = sheet.PrintArea.LeftPt * PointsToPixels;
            double printAreaTopPx = sheet.PrintArea.TopPt * PointsToPixels;
            double printAreaWidthPx = sheet.PrintArea.WidthPt * PointsToPixels;
            double printAreaHeightPx = sheet.PrintArea.HeightPt * PointsToPixels;

            DrawPrintArea(canvas, sheet, printAreaLeftPx, printAreaTopPx, printAreaWidthPx, printAreaHeightPx);

            // Draw clusters as overlays
            var sheetClusters = form.Clusters.Where(c => c.SheetId == sheetId).ToList();
            foreach (var cluster in sheetClusters)
            {
                DrawClusterOverlay(canvas, cluster, pointsToPixels: PointsToPixels);
            }

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(outputPath, data.ToArray());

            _logger.LogInformation("Preview generated: {Path} ({W}x{H})", outputPath, width, height);
            return outputPath;
        }

        private static void DrawPrintArea(SKCanvas canvas, SheetDefinition sheet,
            double leftPx, double topPx, double widthPx, double heightPx)
        {
            using var borderPaint = new SKPaint
            {
                Color = new SKColor(200, 200, 200, 128),
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 1
            };

            canvas.DrawRect((float)leftPx, (float)topPx, (float)widthPx, (float)heightPx, borderPaint);

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

                using var linePaint = new SKPaint
                {
                    Color = new SKColor(220, 220, 220, 100),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = 0.5f
                };
                canvas.DrawLine((float)x, (float)(y + h), (float)(x + widthPx), (float)(y + h), linePaint);

                y += h;
            }

            y = topPx;
            for (int col = 1; col <= maxCol; col++)
            {
                double w = sheet.ColumnWidths.TryGetValue(col, out var cw) ? cw : 64;
                w *= PointsToPixels;

                using var linePaint = new SKPaint
                {
                    Color = new SKColor(220, 220, 220, 100),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = 0.5f
                };
                canvas.DrawLine((float)(x + w), (float)y, (float)(x + w), (float)(y + heightPx), linePaint);

                x += w;
            }
        }

        private static void DrawClusterOverlay(SKCanvas canvas, ClusterDefinition cluster, double pointsToPixels)
        {
            float left = (float)(cluster.LeftPt * pointsToPixels);
            float top = (float)(cluster.TopPt * pointsToPixels);
            float width = (float)(cluster.WidthPt * pointsToPixels);
            float height = (float)(cluster.HeightPt * pointsToPixels);

            var color = cluster.Type switch
            {
                "text" => new SKColor(59, 130, 246, 40),
                "date" => new SKColor(16, 185, 129, 40),
                "checkbox" => new SKColor(245, 158, 11, 40),
                "signature" => new SKColor(139, 92, 246, 40),
                "number" => new SKColor(239, 68, 68, 40),
                _ => new SKColor(255, 212, 0, 40)
            };

            var borderColor = color.WithAlpha(180);

            using var fillPaint = new SKPaint
            {
                Color = color,
                Style = SKPaintStyle.Fill
            };
            using var borderPaint = new SKPaint
            {
                Color = borderColor,
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 2
            };

            canvas.DrawRect(left, top, width, height, fillPaint);
            canvas.DrawRect(left, top, width, height, borderPaint);

            // Draw cluster label
            using var textPaint = new SKPaint
            {
                Color = borderColor.WithAlpha(220),
                TextSize = 10,
                IsAntialias = true
            };
            canvas.DrawText($"{cluster.Name} ({cluster.Type})", left + 2, top + 12, textPaint);
        }
    }
}
