using ExcelAPI.Rendering;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Orchestrates the construction of a RuntimeForm from a parsed RenderWorkbook.
    ///
    /// Reuses the rendering engine's CoordinateEngine, GeometryBuilder, and StyleResolver
    /// for all coordinate and style information — never duplicates parsing or geometry computation.
    ///
    /// Workflow:
    ///   1. Accept a parsed RenderWorkbook (from OpenXmlParser)
    ///   2. Compute geometry via GeometryBuilder (if not already computed)
    ///   3. For each sheet, detect editable fields via FieldDetector
    ///   4. Build RuntimeForm with all metadata
    ///   5. Serialize to JSON for the frontend API
    /// </summary>
    public class FormRuntimeBuilder
    {
        private readonly GeometryBuilder _geometry;
        private readonly CoordinateEngine _coords;
        private readonly FieldDetector _fieldDetector;

        public FormRuntimeBuilder(
            GeometryBuilder geometry,
            CoordinateEngine coords,
            FieldDetector fieldDetector)
        {
            _geometry = geometry;
            _coords = coords;
            _fieldDetector = fieldDetector;
        }

        /// <summary>
        /// Build a RuntimeForm from a parsed workbook.
        /// </summary>
        /// <param name="workbook">Parsed RenderWorkbook from OpenXmlParser.</param>
        /// <param name="dpi">Rendering DPI for pixel coordinates.</param>
        /// <param name="originXPt">Page origin X in points (margin + centering).</param>
        /// <param name="originYPt">Page origin Y in points (margin + centering).</param>
        /// <returns>Complete RuntimeForm ready for JSON serialization.</returns>
        public RuntimeForm Build(
            RenderWorkbook workbook,
            int dpi = 300,
            double originXPt = 0,
            double originYPt = 0)
        {
            // Ensure geometry is computed for all sheets
            foreach (var sheet in workbook.Sheets)
            {
                if (sheet.TotalWidthPt <= 0)
                    _geometry.ComputeGeometry(sheet);
            }

            var runtimeForm = new RuntimeForm
            {
                WorkbookName = Path.GetFileNameWithoutExtension(workbook.FilePath),
                Title = Path.GetFileNameWithoutExtension(workbook.FilePath),
                Dpi = dpi,
                Scale = 1.0,
                Version = "1.0"
            };

            double ptsToPx = dpi / 72.0;

            foreach (var sheet in workbook.Sheets)
            {
                var runtimeSheet = new RuntimeSheet
                {
                    Name = sheet.Name,
                    Index = sheet.Index,
                    PageWidthPx = (int)Math.Round(sheet.TotalWidthPt * ptsToPx),
                    PageHeightPx = (int)Math.Round(sheet.TotalHeightPt * ptsToPx)
                };

                // Detect editable fields using CoordinateEngine pixel bounds
                runtimeSheet.Fields = _fieldDetector.DetectFields(
                    sheet, originXPt, originYPt, dpi, sheet.Index + 1);

                // Convert drawing objects to runtime image/shape references
                if (sheet.DrawingObjects != null)
                {
                    foreach (var drawObj in sheet.DrawingObjects)
                    {
                        if (drawObj.IsImage && drawObj.ImageData != null)
                        {
                            runtimeSheet.Images.Add(new RuntimeImage
                            {
                                Name = drawObj.ImageData.FileName,
                                LeftPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.LeftEmu)),
                                TopPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.TopEmu)),
                                WidthPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.RightEmu - drawObj.LeftEmu)),
                                HeightPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.BottomEmu - drawObj.TopEmu)),
                                ContentType = drawObj.ImageData.ContentType
                            });
                        }
                    }
                }

                // Set page dimensions from the first sheet
                if (runtimeForm.Sheets.Count == 0)
                {
                    runtimeForm.PageWidth = runtimeSheet.PageWidthPx;
                    runtimeForm.PageHeight = runtimeSheet.PageHeightPx;
                }

                runtimeForm.Sheets.Add(runtimeSheet);
            }

            return runtimeForm;
        }
    }
}
