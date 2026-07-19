using ExcelAPI.Models;
using ExcelAPI.Models.WorkbookDefinition;
using ExcelAPI.Rendering;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Designer.Generation
{
    /// <summary>
    /// Generates preview PNG images for forms.
    /// Canonical path: WorkbookDefinition → ExportCoordinator.ExportPng().
    ///
    /// Phase 4.2: Primary overload accepts WorkbookDefinition.
    ///            Fallback overlay-only path removed — WbDef is always available.
    ///            No rendering logic in this class — only coordinates exports.
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
        /// Generate a preview PNG from a FormDefinition (compatibility shim).
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
        /// Generate a preview PNG from a WorkbookDefinition (canonical path).
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
                    _logger.LogWarning(ex, "ExportCoordinator failed for preview");
                }
            }

            // Fallback: generate a blank white placeholder (no overlay rendering)
            double ptsToPx = Rectangle.PtToPx(1, 300);
            int fallbackW = (int)Math.Round((sheet.PrintLayout?.PageWidthPt ?? 612) * ptsToPx);
            int fallbackH = (int)Math.Round((sheet.PrintLayout?.PageHeightPt ?? 792) * ptsToPx);
            using var bitmap = new SkiaSharp.SKBitmap(Math.Max(1, fallbackW), Math.Max(1, fallbackH));
            using var canvas = new SkiaSharp.SKCanvas(bitmap);
            canvas.Clear(SkiaSharp.SKColors.White);
            using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(outputPath, data.ToArray());

            _logger.LogInformation("Blank fallback preview generated: {Path} ({W}x{H})", outputPath, fallbackW, fallbackH);
            return outputPath;
        }
    }
}
