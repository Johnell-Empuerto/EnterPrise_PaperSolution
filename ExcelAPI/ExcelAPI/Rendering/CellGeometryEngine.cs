using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Maps cell coordinates to canvas pixel positions.
    /// Now delegates to CoordinateEngine as the canonical coordinate source (Phase 11D).
    /// Retained for backward compatibility with existing render layers.
    /// </summary>
    public class CellGeometryEngine
    {
        private readonly CoordinateEngine _coord;

        public CellGeometryEngine()
        {
            _coord = new CoordinateEngine();
        }

        public CellGeometryEngine(CoordinateEngine coord)
        {
            _coord = coord;
        }

        /// <summary>Get the page-relative pixel bounds of a cell.</summary>
        public SKRect GetCellPixelBounds(RenderSheet sheet, uint col, uint row,
            double originXPt, double originYPt)
        {
            return _coord.GetCellPixelBounds(sheet, col, row, originXPt, originYPt);
        }

        /// <summary>Get the page-relative pixel bounds of a merged range.</summary>
        public SKRect GetMergePixelBounds(RenderMerge merge, double originXPt, double originYPt)
        {
            return _coord.GetMergePixelBounds(merge, originXPt, originYPt);
        }

        /// <summary>Get the page-relative pixel bounds at cell/merge location.</summary>
        public SKRect GetCellOrMergePixelBounds(RenderSheet sheet, RenderCell cell,
            double originXPt, double originYPt)
        {
            return _coord.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt);
        }

        /// <summary>Convert character width to points using ECMA-376 formula.</summary>
        public static double CharWidthToPoints(double charWidth, double maxDigitWidth = 7.33)
        {
            return CoordinateEngine.CharWidthToPoints(charWidth, maxDigitWidth);
        }
    }
}
