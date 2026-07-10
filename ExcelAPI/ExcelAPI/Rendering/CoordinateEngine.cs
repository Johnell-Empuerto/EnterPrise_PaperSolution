using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Canonical coordinate mapping engine.
    /// Converts between coordinate spaces: cell → points → pixels → page.
    ///
    /// Every render layer consumes coordinates from this engine.
    /// No layer independently computes cell bounds, merge bounds, or pixel conversion.
    ///
    /// Coordinate flow:
    ///   Cell (col, row)
    ///     ↓ GeometryBuilder.ComputedColCumLeftPt / ComputedRowCumTopPt
    ///   Points (worksheet-relative)
    ///     ↓ + originXPt (margin + centering)
    ///   Points (page-relative)
    ///     ↓ * PointsToPixels (DPI/72)
    ///   Pixels (canvas)
    /// </summary>
    public class CoordinateEngine
    {
        /// <summary>Points-to-pixels conversion at 300 DPI.</summary>
        public const double Dpi = 300.0;
        public const double PointsToPixels = Dpi / 72.0;

        /// <summary>Convert character width to points using ECMA-376 formula.</summary>
        public static double CharWidthToPoints(double charWidth, double maxDigitWidth = 7.33)
        {
            return (charWidth * maxDigitWidth + 5.0) * 72.0 / 96.0;
        }

        /// <summary>Convert points to pixels at 300 DPI.</summary>
        public static double PtToPx(double pt) => pt * PointsToPixels;

        /// <summary>Convert pixels to points at 300 DPI.</summary>
        public static double PxToPt(double px) => px / PointsToPixels;

        /// <summary>Get the page-relative point position of a column's left edge.</summary>
        public double GetColLeftPt(RenderSheet sheet, uint col, double originXPt)
        {
            return originXPt + sheet.ComputedColCumLeftPt.GetValueOrDefault(col, 0);
        }

        /// <summary>Get the page-relative point position of a row's top edge.</summary>
        public double GetRowTopPt(RenderSheet sheet, uint row, double originYPt)
        {
            return originYPt + sheet.ComputedRowCumTopPt.GetValueOrDefault(row, 0);
        }

        /// <summary>Get the page-relative point width of a column.</summary>
        public double GetColWidthPt(RenderSheet sheet, uint col)
        {
            return sheet.ComputedColWidthsPt.GetValueOrDefault(col,
                CharWidthToPoints(sheet.DefaultColumnWidth, sheet.MaxDigitWidth));
        }

        /// <summary>Get the page-relative point height of a row.</summary>
        public double GetRowHeightPt(RenderSheet sheet, uint row)
        {
            return sheet.ComputedRowHeightsPt.GetValueOrDefault(row, sheet.DefaultRowHeight);
        }

        /// <summary>Get the page-relative pixel bounds of a cell.</summary>
        public SKRect GetCellPixelBounds(RenderSheet sheet, uint col, uint row,
            double originXPt, double originYPt)
        {
            double leftPt = GetColLeftPt(sheet, col, originXPt);
            double topPt = GetRowTopPt(sheet, row, originYPt);
            double wPt = GetColWidthPt(sheet, col);
            double hPt = GetRowHeightPt(sheet, row);

            return new SKRect(
                (float)PtToPx(leftPt),
                (float)PtToPx(topPt),
                (float)PtToPx(leftPt + wPt),
                (float)PtToPx(topPt + hPt));
        }

        /// <summary>Get the page-relative pixel bounds of a merged range.</summary>
        public SKRect GetMergePixelBounds(RenderMerge merge, double originXPt, double originYPt)
        {
            return new SKRect(
                (float)PtToPx(originXPt + merge.LeftPt),
                (float)PtToPx(originYPt + merge.TopPt),
                (float)PtToPx(originXPt + merge.LeftPt + merge.WidthPt),
                (float)PtToPx(originYPt + merge.TopPt + merge.HeightPt));
        }

        /// <summary>Get the page-relative pixel bounds at cell/merge location.</summary>
        public SKRect GetCellOrMergePixelBounds(RenderSheet sheet, RenderCell cell,
            double originXPt, double originYPt)
        {
            if (cell.MergeIndex.HasValue && cell.MergeIndex.Value >= 0
                && cell.MergeIndex.Value < sheet.Merges.Count)
            {
                return GetMergePixelBounds(sheet.Merges[cell.MergeIndex.Value],
                    originXPt, originYPt);
            }
            return GetCellPixelBounds(sheet, cell.ColumnIndex, cell.RowIndex,
                originXPt, originYPt);
        }
    }
}
