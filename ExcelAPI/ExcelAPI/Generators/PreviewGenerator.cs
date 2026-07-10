using ExcelAPI.Models;
using ExcelAPI.Rendering;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Generators
{
    /// <summary>
    /// Generates preview PNG images for forms.
    /// Delegates all rendering to ExportCoordinator (Phase 11G Production Export Engine).
    ///
    /// No rendering logic should exist in this class — it only coordinates
    /// the export options and delegates to ExportCoordinator.ExportPng().
    /// </summary>
    public class PreviewGenerator
    {
        private readonly ILogger<PreviewGenerator> _logger;
        private readonly ExportCoordinator _exportCoordinator;
        private readonly string? _xlsxPath;

        public PreviewGenerator(ILogger<PreviewGenerator> logger, ExportCoordinator exportCoordinator)
        {
            _logger = logger;
            _exportCoordinator = exportCoordinator;
        }

        public PreviewGenerator(ILogger<PreviewGenerator> logger, ExportCoordinator exportCoordinator, string xlsxPath)
        {
            _logger = logger;
            _exportCoordinator = exportCoordinator;
            _xlsxPath = xlsxPath;
        }

        /// <summary>
        /// Generate a preview PNG for the given form and sheet.
        /// Delegates to ExportCoordinator.ExportPng() with default preview options.
        /// Falls back to overlay-only if XLSX is unavailable or rendering fails.
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
                        FilePrefix = Path.GetFileNameWithoutExtension(outputPath),
                        OutputDirectory = Path.GetDirectoryName(outputPath),
                        PngCompressionLevel = 100,
                        TransparentBackground = false,
                        IncludePageNumbers = false,
                        MaxPages = 1
                    };

                    var results = _exportCoordinator.ExportPng(xlsxPath, form, sheetId, options);
                    if (results.Count > 0)
                    {
                        // If ExportCoordinator used a different filename, copy/rename
                        if (!string.Equals(results[0], outputPath, StringComparison.OrdinalIgnoreCase)
                            && System.IO.File.Exists(results[0]))
                        {
                            var bytes = System.IO.File.ReadAllBytes(results[0]);
                            System.IO.File.WriteAllBytes(outputPath, bytes);
                        }

                        _logger.LogInformation("Preview generated via ExportCoordinator: {Path}", outputPath);
                        return outputPath;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "ExportCoordinator failed, falling back to overlay-only preview");
                }
            }

            // Fallback: overlay-only preview
            return GenerateOverlayOnly(form, sheetId, sheet, outputPath);
        }

        private string GenerateOverlayOnly(FormDefinition form, string sheetId, SheetDefinition sheet, string outputPath)
        {
            const double PointsToPixels = 300.0 / 72.0;

            double pageWidthPx = sheet.PageSettings.WidthPt * PointsToPixels;
            double pageHeightPx = sheet.PageSettings.HeightPt * PointsToPixels;

            int width = (int)Math.Round(pageWidthPx);
            int height = (int)Math.Round(pageHeightPx);

            using var bitmap = new SkiaSharp.SKBitmap(width, height);
            using var canvas = new SkiaSharp.SKCanvas(bitmap);

            canvas.Clear(SkiaSharp.SKColors.White);

            // Draw clusters as overlays
            var sheetClusters = form.Clusters.Where(c => c.SheetId == sheetId).ToList();
            foreach (var cluster in sheetClusters)
            {
                DrawClusterOverlay(canvas, cluster, PointsToPixels);
            }

            using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(outputPath, data.ToArray());

            _logger.LogInformation("Overlay-only preview generated: {Path} ({W}x{H})", outputPath, width, height);
            return outputPath;
        }

        private static void DrawClusterOverlay(SkiaSharp.SKCanvas canvas, ClusterDefinition cluster, double pointsToPixels)
        {
            float left = (float)(cluster.LeftPt * pointsToPixels);
            float top = (float)(cluster.TopPt * pointsToPixels);
            float width = (float)(cluster.WidthPt * pointsToPixels);
            float height = (float)(cluster.HeightPt * pointsToPixels);

            var color = cluster.Type switch
            {
                "text" => new SkiaSharp.SKColor(59, 130, 246, 40),
                "date" => new SkiaSharp.SKColor(16, 185, 129, 40),
                "checkbox" => new SkiaSharp.SKColor(245, 158, 11, 40),
                "signature" => new SkiaSharp.SKColor(139, 92, 246, 40),
                "number" => new SkiaSharp.SKColor(239, 68, 68, 40),
                _ => new SkiaSharp.SKColor(255, 212, 0, 40)
            };

            var borderColor = color.WithAlpha(180);

            using var fillPaint = new SkiaSharp.SKPaint
            {
                Color = color,
                Style = SkiaSharp.SKPaintStyle.Fill
            };
            using var borderPaint = new SkiaSharp.SKPaint
            {
                Color = borderColor,
                Style = SkiaSharp.SKPaintStyle.Stroke,
                StrokeWidth = 2
            };

            canvas.DrawRect(left, top, width, height, fillPaint);
            canvas.DrawRect(left, top, width, height, borderPaint);

            using var textPaint = new SkiaSharp.SKPaint
            {
                Color = borderColor.WithAlpha(220),
                TextSize = 10,
                IsAntialias = true
            };
            canvas.DrawText($"{cluster.Name} ({cluster.Type})", left + 2, top + 12, textPaint);
        }
    }
}
