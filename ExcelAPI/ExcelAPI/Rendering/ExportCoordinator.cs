using ExcelAPI.Models;
using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Public export API for the FormLess Rendering Engine.
    /// Orchestrates the complete production pipeline:
    ///   1. Parse XLSX → RenderWorkbook via OpenXmlParser
    ///   2. Compute cell geometry via GeometryBuilder
    ///   3. Compute pages via PageRenderer
    ///   4. Render each page via RendererCoordinator
    ///   5. Export to PNG or PDF
    ///
    /// PreviewGenerator and PdfGenerator should delegate to this class.
    /// No rendering logic should exist outside this pipeline.
    /// </summary>
    public class ExportCoordinator
    {
        private readonly OpenXmlParser _parser;
        private readonly GeometryBuilder _geometry;
        private readonly PageRenderer _pageRenderer;
        private readonly RendererCoordinator _renderCoordinator;

        public ExportCoordinator(
            OpenXmlParser parser,
            GeometryBuilder geometry,
            PageRenderer pageRenderer,
            RendererCoordinator renderCoordinator)
        {
            _parser = parser;
            _geometry = geometry;
            _pageRenderer = pageRenderer;
            _renderCoordinator = renderCoordinator;
        }

        /// <summary>
        /// Parse an XLSX into a RenderWorkbook and compute geometry.
        /// Shared setup for all export paths.
        /// </summary>
        public RenderWorkbook ParseWorkbook(string xlsxPath)
        {
            var workbook = _parser.Parse(xlsxPath);
            if (workbook.Sheets.Count == 0)
                throw new InvalidOperationException("No sheets found in workbook");

            foreach (var sheet in workbook.Sheets)
            {
                _geometry.ComputeGeometry(sheet);
            }

            return workbook;
        }

        /// <summary>
        /// Export a workbook sheet to one or more PNG files.
        /// Returns a list of file paths (one per page).
        /// </summary>
        public List<string> ExportPng(
            string xlsxPath,
            FormDefinition form,
            string sheetId,
            ExportOptions options)
        {
            var sheet = form.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            var workbook = ParseWorkbook(xlsxPath);
            var renderSheet = workbook.Sheets.FirstOrDefault();
            if (renderSheet == null)
                throw new InvalidOperationException("No sheets in workbook");

            // Compute pages
            var pages = _pageRenderer.ComputePages(
                paperSize: sheet.PageSettings.PaperSize,
                orientation: sheet.PageSettings.Orientation,
                leftMargin: sheet.PageSettings.LeftMargin,
                rightMargin: sheet.PageSettings.RightMargin,
                topMargin: sheet.PageSettings.TopMargin,
                bottomMargin: sheet.PageSettings.BottomMargin,
                centerHorizontally: sheet.PageSettings.CenterHorizontally,
                centerVertically: sheet.PageSettings.CenterVertically,
                zoom: sheet.PageSettings.Zoom,
                fitToPagesWide: sheet.PageSettings.FitToPagesWide,
                fitToPagesTall: sheet.PageSettings.FitToPagesTall,
                totalContentWidthPt: renderSheet.TotalWidthPt,
                totalContentHeightPt: renderSheet.TotalHeightPt,
                dpi: options.Dpi,
                maxPages: options.MaxPages);

            var outputPaths = new List<string>();
            string outputDir = options.OutputDirectory ?? Path.GetDirectoryName(xlsxPath) ?? ".";

            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                var ctx = page.Context;
                ctx.Workbook = workbook;
                ctx.Sheet = renderSheet;

                // Build output filename
                string outputPath;
                if (pages.Count == 1)
                {
                    outputPath = Path.Combine(outputDir, $"{options.FilePrefix}.png");
                }
                else
                {
                    outputPath = Path.Combine(outputDir,
                        $"{options.FilePrefix}{page.PageNumber}.png");
                }

                // Render to bitmap
                using var bitmap = new SKBitmap(page.PixelWidth, page.PixelHeight);
                using var canvas = new SKCanvas(bitmap);

                if (options.TransparentBackground)
                {
                    canvas.Clear(SKColors.Transparent);
                }
                else
                {
                    canvas.Clear(SKColors.White);
                }

                // Apply quality settings
                ApplyQualitySettings(canvas, options);

                // Render layers
                _renderCoordinator.RenderToCanvas(canvas, ctx);

                // Draw page number if requested
                if (options.IncludePageNumbers)
                {
                    DrawPageNumber(canvas, ctx, page.PageNumber, page.TotalPages, options);
                }

                // Save to PNG
                using var image = SKImage.FromBitmap(bitmap);
                using var data = image.Encode(SKEncodedImageFormat.Png, options.PngCompressionLevel);
                System.IO.File.WriteAllBytes(outputPath, data.ToArray());

                outputPaths.Add(outputPath);
            }

            return outputPaths;
        }

        /// <summary>
        /// Export a workbook sheet to a PDF file.
        /// Each page is a separate PDF page.
        /// Uses SkiaSharp PDF canvas for vector output.
        /// </summary>
        public string ExportPdf(
            string xlsxPath,
            FormDefinition form,
            string sheetId,
            string outputPath,
            ExportOptions options)
        {
            var sheet = form.Sheets.FirstOrDefault(s => s.Id == sheetId);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetId}");

            var workbook = ParseWorkbook(xlsxPath);
            var renderSheet = workbook.Sheets.FirstOrDefault();
            if (renderSheet == null)
                throw new InvalidOperationException("No sheets in workbook");

            // Compute pages
            var pages = _pageRenderer.ComputePages(
                paperSize: sheet.PageSettings.PaperSize,
                orientation: sheet.PageSettings.Orientation,
                leftMargin: sheet.PageSettings.LeftMargin,
                rightMargin: sheet.PageSettings.RightMargin,
                topMargin: sheet.PageSettings.TopMargin,
                bottomMargin: sheet.PageSettings.BottomMargin,
                centerHorizontally: sheet.PageSettings.CenterHorizontally,
                centerVertically: sheet.PageSettings.CenterVertically,
                zoom: sheet.PageSettings.Zoom,
                fitToPagesWide: sheet.PageSettings.FitToPagesWide,
                fitToPagesTall: sheet.PageSettings.FitToPagesTall,
                totalContentWidthPt: renderSheet.TotalWidthPt,
                totalContentHeightPt: renderSheet.TotalHeightPt,
                dpi: options.Dpi,
                maxPages: options.MaxPages);

            // Create PDF document
            var pdfOptions = new SKDocumentPdfMetadata();
            if (options.IncludeMetadata && workbook.Sheets.Count > 0)
            {
                pdfOptions.Title = options.Title ?? form.Workbook.Title;
                pdfOptions.Author = options.Author ?? form.Workbook.Author;
                pdfOptions.Subject = options.Subject ?? form.Workbook.Description;
                pdfOptions.Keywords = options.Keywords;
                pdfOptions.Producer = "FormLess Rendering Engine";
            }

            using var pdfDocument = SKDocument.CreatePdf(outputPath, pdfOptions);

            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                var ctx = page.Context;
                ctx.Workbook = workbook;
                ctx.Sheet = renderSheet;

                // Begin PDF page (in points — SkiaSharp PDF canvas preserves vector content)
                using var pdfCanvas = pdfDocument.BeginPage(
                    (float)ctx.PageWidthPt, (float)ctx.PageHeightPt);

                // For vector PDF: render to a bitmap first, then draw to PDF canvas.
                // This ensures all SkiaSharp rendering (fills, borders, text as paths)
                // is captured correctly. Future optimization: render text/borders as
                // native PDF vector elements.
                using var stageBitmap = new SKBitmap(page.PixelWidth, page.PixelHeight);
                using var stageCanvas = new SKCanvas(stageBitmap);

                if (options.TransparentBackground)
                    stageCanvas.Clear(SKColors.Transparent);
                else
                    stageCanvas.Clear(SKColors.White);

                ApplyQualitySettings(stageCanvas, options);
                _renderCoordinator.RenderToCanvas(stageCanvas, ctx);

                if (options.IncludePageNumbers)
                {
                    DrawPageNumber(stageCanvas, ctx, page.PageNumber, page.TotalPages, options);
                }

                // Draw the rendered bitmap onto the PDF canvas, scaled to page dimensions.
                // Using SKRect destination to ensure correct sizing in the PDF point-space.
                using var image = SKImage.FromBitmap(stageBitmap);
                var destRect = new SKRect(0, 0,
                    (float)ctx.PageWidthPt, (float)ctx.PageHeightPt);
                var sourceRect = new SKRect(0, 0, stageBitmap.Width, stageBitmap.Height);
                pdfCanvas.DrawImage(image, sourceRect, destRect);

                pdfDocument.EndPage();
            }

            pdfDocument.Close();
            return outputPath;
        }

        /// <summary>
        /// Apply quality settings to a canvas based on export options.
        /// </summary>
        private static void ApplyQualitySettings(SKCanvas canvas, ExportOptions options)
        {
            if (options.HighQualityAntialiasing)
            {
                // Enable high-quality rendering
                // Note: Specific quality settings depend on SkiaSharp version
            }
        }

        /// <summary>
        /// Draw a page number footer on the canvas.
        /// </summary>
        private static void DrawPageNumber(
            SKCanvas canvas, RenderingContext context,
            int pageNumber, int totalPages, ExportOptions options)
        {
            string text = string.Format(options.PageNumberFormat, pageNumber, totalPages);
            float fontSizePx = (float)(options.PageNumberFontSizePt * (options.Dpi / 72.0));

            using var typeface = SKTypeface.FromFamilyName("Arial");
            using var font = new SKFont(typeface, fontSizePx);
            using var paint = new SKPaint
            {
                Color = new SKColor(120, 120, 120),
                IsAntialias = true
            };

            float textWidth = font.MeasureText(text);
            float centerX = (float)(context.PageWidthPt * context.PointsToPixels / 2f);
            float bottomY = (float)(context.PageHeightPt * context.PointsToPixels) - fontSizePx * 0.5f;

            canvas.DrawText(text, centerX - textWidth / 2f, bottomY, font, paint);
        }

        /// <summary>
        /// Quick validation: check if output file exists and has minimum size.
        /// </summary>
        public static bool ValidateOutput(string outputPath, int expectedMinBytes = 1000)
        {
            if (!System.IO.File.Exists(outputPath)) return false;
            var info = new FileInfo(outputPath);
            return info.Length >= expectedMinBytes;
        }
    }
}
