using System.Text.Json.Serialization;

// ────────────────────────────────────────────────────────────────────────────
// CoordinateModel — Geometry Primitives for WorkbookDefinition
//
// These are value types used throughout the WorkbookDefinition hierarchy to
// describe positions and sizes in points (Excel's native unit), ratios
// (for responsive layouts), and cell references.
//
// Ownership: Shared
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// A 2D point in Excel points (1 point = 1/72 inch).
    /// </summary>
    public class Point
    {
        public double X { get; set; }
        public double Y { get; set; }

        public Point() { }

        public Point(double x, double y)
        {
            X = x;
            Y = y;
        }
    }

    /// <summary>
    /// A rectangle in Excel points, describing position and size.
    /// </summary>
    public class Rectangle
    {
        /// <summary>Left edge in points (worksheet-relative or page-relative).</summary>
        public double Left { get; set; }

        /// <summary>Top edge in points.</summary>
        public double Top { get; set; }

        /// <summary>Width in points.</summary>
        public double Width { get; set; }

        /// <summary>Height in points.</summary>
        public double Height { get; set; }

        /// <summary>Computed right edge (Left + Width).</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double Right => Left + Width;

        /// <summary>Computed bottom edge (Top + Height).</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double Bottom => Top + Height;

                /// <summary>
        /// Convert a rectangle from points to pixels at the given DPI.
        /// </summary>
        public Rectangle ToPixels(double dpi)
        {
            double ptsToPx = dpi / 72.0;
            return new Rectangle(
                Left * ptsToPx,
                Top * ptsToPx,
                Width * ptsToPx,
                Height * ptsToPx
            );
        }

        /// <summary>Convert points to pixels at the given DPI.</summary>
        public static double PtToPx(double pt, double dpi) => pt * (dpi / 72.0);

        /// <summary>Convert pixels to points at the given DPI.</summary>
        public static double PxToPt(double px, double dpi) => px / (dpi / 72.0);

        public Rectangle() { }

        public Rectangle(double left, double top, double width, double height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }

    /// <summary>
    /// A rectangle expressed as ratios of a container (typically page width/height).
    /// Used for responsive layouts and legacy ConMas compatibility.
    /// All values are in the range [0, 1].
    /// </summary>
    public class RatioRectangle
    {
        /// <summary>Left edge as ratio of container width (0–1).</summary>
        public double LeftRatio { get; set; }

        /// <summary>Top edge as ratio of container height (0–1).</summary>
        public double TopRatio { get; set; }

        /// <summary>Width as ratio of container width (0–1).</summary>
        public double WidthRatio { get; set; }

        /// <summary>Height as ratio of container height (0–1).</summary>
        public double HeightRatio { get; set; }

        /// <summary>Computed right ratio (LeftRatio + WidthRatio).</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double RightRatio => LeftRatio + WidthRatio;

        /// <summary>Computed bottom ratio (TopRatio + HeightRatio).</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double BottomRatio => TopRatio + HeightRatio;

        public RatioRectangle() { }

        public RatioRectangle(double leftRatio, double topRatio, double widthRatio, double heightRatio)
        {
            LeftRatio = leftRatio;
            TopRatio = topRatio;
            WidthRatio = widthRatio;
            HeightRatio = heightRatio;
        }

        /// <summary>
        /// Convert to pixel coordinates at a given container size and DPI.
        /// </summary>
        public Rectangle ToPixelRect(double containerWidthPt, double containerHeightPt, double dpi)
        {
            double ptsToPx = dpi / 72.0;
            return new Rectangle(
                LeftRatio * containerWidthPt * ptsToPx,
                TopRatio * containerHeightPt * ptsToPx,
                WidthRatio * containerWidthPt * ptsToPx,
                HeightRatio * containerHeightPt * ptsToPx
            );
        }
    }

    /// <summary>
    /// A cell reference in A1-style notation with resolved row/column indices.
    /// </summary>
    public class CellReference
    {
        /// <summary>A1-style address (e.g., "B5", "AA123").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>1-based row index.</summary>
        public uint RowIndex { get; set; }

        /// <summary>1-based column index.</summary>
        public uint ColumnIndex { get; set; }

        /// <summary>Column letter(s) (e.g., "A", "B", "AA").</summary>
        public string ColumnLetter { get; set; } = string.Empty;

        public CellReference() { }

        public CellReference(string address, uint row, uint col)
        {
            Address = address;
            RowIndex = row;
            ColumnIndex = col;
            ColumnLetter = ColumnIndexToLetters(col);
        }

        /// <summary>
        /// Convert a 1-based column index to column letters (A, B, ..., Z, AA, AB, ...).
        /// </summary>
        public static string ColumnIndexToLetters(uint colIndex)
        {
            string letters = string.Empty;
            while (colIndex > 0)
            {
                colIndex--;
                letters = (char)('A' + colIndex % 26) + letters;
                colIndex /= 26;
            }
            return letters;
        }
    }

    /// <summary>
    /// A range of cells (e.g., "A1:C3").
    /// </summary>
    public class CellRange
    {
        /// <summary>Standard range address (e.g., "A1:C3").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>First (top-left) cell in the range.</summary>
        public CellReference? FirstCell { get; set; }

        /// <summary>Last (bottom-right) cell in the range.</summary>
        public CellReference? LastCell { get; set; }

        /// <summary>Number of columns spanned.</summary>
        public uint ColumnSpan => LastCell != null && FirstCell != null
            ? LastCell.ColumnIndex - FirstCell.ColumnIndex + 1
            : 0;

        /// <summary>Number of rows spanned.</summary>
        public uint RowSpan => LastCell != null && FirstCell != null
            ? LastCell.RowIndex - FirstCell.RowIndex + 1
            : 0;
    }
}
