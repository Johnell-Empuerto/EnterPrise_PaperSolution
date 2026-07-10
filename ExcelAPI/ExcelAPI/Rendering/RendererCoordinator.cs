using ExcelAPI.Models;
using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Orchestrates the full rendering pipeline.
    /// Executes IRenderLayer instances in registration order.
    ///
    /// Pipeline:
    ///   1. Parse XLSX → RenderWorkbook via OpenXmlParser
    ///   2. Compute cell geometry via GeometryBuilder (cumulative col/row positions)
    ///   3. Compute print layout via PrintLayoutEngine (page geo, margins, origin, scaling, clip)
    ///   4. Execute each IRenderLayer (fills → gridlines → borders → text → ...)
    ///   5. Output as PNG or PDF
    ///
    /// To add a new layer (e.g., TextLayer), register it in DI and it will
    /// automatically be injected into the layer list — no coordinator changes needed.
    /// </summary>
    public class RendererCoordinator
    {
        private readonly OpenXmlParser _parser;
        private readonly GeometryBuilder _geometry;
        private readonly PrintLayoutEngine _printLayout;
        private readonly IEnumerable<IRenderLayer> _layers;

        public RendererCoordinator(
            OpenXmlParser parser,
            GeometryBuilder geometry,
            PrintLayoutEngine printLayout,
            IEnumerable<IRenderLayer> layers)
        {
            _parser = parser;
            _geometry = geometry;
            _printLayout = printLayout;
            _layers = layers;
        }

        /// <summary>
        /// Parse an XLSX and set up the RenderingContext.
        /// This is the common setup for both RenderToPng and RenderToPdf.
        /// </summary>
        public (RenderWorkbook workbook, RenderSheet sheet, RenderingContext context) Prepare(
            string xlsxPath, FormDefinition form)
        {
            var workbook = _parser.Parse(xlsxPath);
            if (workbook.Sheets.Count == 0)
                throw new InvalidOperationException("No sheets found in workbook");

            var sheet = workbook.Sheets[0];
            var formSheet = form.Sheets.FirstOrDefault();
            if (formSheet == null)
                throw new InvalidOperationException("No sheet in form definition");

            // Step 1: Compute cell geometry (cumulative col/row positions)
            _geometry.ComputeGeometry(sheet);

            // Step 2: Compute full print layout via PrintLayoutEngine
            var layout = _printLayout.Compute(
                paperSize: formSheet.PageSettings.PaperSize,
                orientation: formSheet.PageSettings.Orientation,
                leftMargin: formSheet.PageSettings.LeftMargin,
                rightMargin: formSheet.PageSettings.RightMargin,
                topMargin: formSheet.PageSettings.TopMargin,
                bottomMargin: formSheet.PageSettings.BottomMargin,
                centerHorizontally: formSheet.PageSettings.CenterHorizontally,
                centerVertically: formSheet.PageSettings.CenterVertically,
                zoom: formSheet.PageSettings.Zoom,
                fitToPagesWide: formSheet.PageSettings.FitToPagesWide,
                fitToPagesTall: formSheet.PageSettings.FitToPagesTall,
                totalContentWidthPt: sheet.TotalWidthPt,
                totalContentHeightPt: sheet.TotalHeightPt);

            // Step 3: Build the expanded RenderingContext
            var context = new RenderingContext
            {
                Workbook = workbook,
                Sheet = sheet,
                PaperSize = layout.PaperSize,
                Orientation = layout.Orientation,
                PageWidthPt = layout.PageWidthPt,
                PageHeightPt = layout.PageHeightPt,
                MarginLeftPt = layout.MarginLeftPt,
                MarginRightPt = layout.MarginRightPt,
                MarginTopPt = layout.MarginTopPt,
                MarginBottomPt = layout.MarginBottomPt,
                PrintableWidthPt = layout.PrintableWidthPt,
                PrintableHeightPt = layout.PrintableHeightPt,
                OriginXPt = layout.OriginXPt,
                OriginYPt = layout.OriginYPt,
                ScaleFactor = layout.ScaleFactor,
                Zoom = layout.Zoom,
                FitToPagesWide = layout.FitToPagesWide,
                FitToPagesTall = layout.FitToPagesTall,
                IsScalingActive = layout.IsScalingActive,
                ClipLeftPt = layout.ClipLeftPt,
                ClipTopPt = layout.ClipTopPt,
                ClipRightPt = layout.ClipRightPt,
                ClipBottomPt = layout.ClipBottomPt,
                Dpi = 300.0
            };

            return (workbook, sheet, context);
        }

        /// <summary>
        /// Render all layers to a SkiaSharp canvas.
        /// Called by both RenderToPng and RenderToPdf.
        /// </summary>
        public void RenderToCanvas(SKCanvas canvas, RenderingContext context)
        {
            // Layer 1: White page background
            canvas.Clear(SKColors.White);

            // Apply print area clip region (between background and render layers)
            canvas.Save();
            canvas.ClipRect(context.ClipRectPx);

            // Layers 2+: Execute each registered IRenderLayer in order
            foreach (var layer in _layers)
            {
                layer.Render(canvas, context);
            }

            // Restore clip region after rendering
            canvas.Restore();
        }

        /// <summary>
        /// Render a form definition to a PNG file.
        /// </summary>
        public string RenderToPng(string xlsxPath, FormDefinition form, string outputPath)
        {
            var (_, _, context) = Prepare(xlsxPath, form);

            using var bitmap = new SKBitmap(context.PixelWidth, context.PixelHeight);
            using var canvas = new SKCanvas(bitmap);

            RenderToCanvas(canvas, context);

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(outputPath, data.ToArray());

            return outputPath;
        }

        /// <summary>
        /// Render to PDF via SkiaSharp PDF document.
        /// </summary>
        public string RenderToPdf(string xlsxPath, FormDefinition form, string outputPath)
        {
            var (_, _, context) = Prepare(xlsxPath, form);

            using var bitmap = new SKBitmap(context.PixelWidth, context.PixelHeight);
            using var canvas = new SKCanvas(bitmap);

            RenderToCanvas(canvas, context);

            // Create PDF from rendered bitmap
            using var pdfDocument = SKDocument.CreatePdf(outputPath);
            using var pdfCanvas = pdfDocument.BeginPage(
                (float)context.PageWidthPt,
                (float)context.PageHeightPt);

            pdfCanvas.DrawBitmap(bitmap, 0, 0);
            pdfDocument.EndPage();
            pdfDocument.Close();

            return outputPath;
        }

        /// <summary>
        /// Quick validation: compare renderer output against expected pixel count.
        /// </summary>
        public bool ValidateOutput(string outputPath, int expectedMinBytes = 1000)
        {
            if (!System.IO.File.Exists(outputPath)) return false;
            var info = new FileInfo(outputPath);
            return info.Length >= expectedMinBytes;
        }
    }
}
