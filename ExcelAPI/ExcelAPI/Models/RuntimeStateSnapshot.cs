using System.Text.Json.Serialization;

namespace ExcelAPI.Models
{
    /// <summary>
    /// Serializable snapshot of every COM property existing at a single point in
    /// the rendering pipeline. Captured at each stage of CaptureGeometryTimeline().
    /// </summary>
    public class RuntimeStateSnapshot
    {
        public string Stage { get; set; } = "";
        public string WorkbookName { get; set; } = "";
        public string WorksheetName { get; set; } = "";
        public int WorksheetIndex { get; set; }
        public string SheetDimension { get; set; } = "";

        // ── Worksheet-level geometry ────────────────────────────────────
        public RangeGeometry? UsedRange { get; set; }
        public RangeGeometry? PrintArea { get; set; }
        public RangeGeometry? CurrentRegion { get; set; }
        public string? PrintAreaAddress { get; set; }
        public string? UsedRangeAddress { get; set; }

        // ── PageSetup — ALL properties ──────────────────────────────────
        public PageSetupSnapshot? PageSetup { get; set; }

        // ── Window State ────────────────────────────────────────────────
        public WindowStateSnapshot? WindowState { get; set; }

        // ── Shapes ──────────────────────────────────────────────────────
        public List<ShapeSnapshot>? Shapes { get; set; }

        // ── Printer ─────────────────────────────────────────────────────
        public PrinterStateSnapshot? Printer { get; set; }

        // ── Field Cells ─────────────────────────────────────────────────
        public List<FieldCellSnapshot>? FieldCells { get; set; }

        // ── Application ─────────────────────────────────────────────────
        public string? ActivePrinter { get; set; }
        public string? Version { get; set; }
        public string? CalculationVersion { get; set; }
    }

    /// <summary>Geometry properties common to all Excel Range objects.</summary>
    public class RangeGeometry
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public int ColumnsCount { get; set; }
        public int RowsCount { get; set; }
        public string? Address { get; set; }

        /// <summary>EntireColumn geometry (cumulative boundary).</summary>
        public double EntireColumnLeft { get; set; }
        public double EntireColumnWidth { get; set; }

        /// <summary>EntireRow geometry.</summary>
        public double EntireRowTop { get; set; }
        public double EntireRowHeight { get; set; }
    }

    /// <summary>Every PageSetup property exposed by Excel COM.</summary>
    public class PageSetupSnapshot
    {
        // Margins
        public double LeftMargin { get; set; }
        public double RightMargin { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double HeaderMargin { get; set; }
        public double FooterMargin { get; set; }

        // Centering
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }

        // Paper & orientation
        public int PaperSize { get; set; }
        public int Orientation { get; set; }

        // Scaling
        public int Zoom { get; set; }
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
        public bool Draft { get; set; }
        public bool BlackAndWhite { get; set; }

        // Print order
        public int Order { get; set; }

        // Print quality
        public int PrintQuality { get; set; }

        // Page numbering
        public int FirstPageNumber { get; set; }

        // Print titles
        public string? PrintTitleRows { get; set; }
        public string? PrintTitleColumns { get; set; }

        // Print area
        public string? PrintArea { get; set; }

        // Page dimensions (calculated from PaperSize + Orientation)
        public double PageWidthPt { get; set; }
        public double PageHeightPt { get; set; }
    }

    /// <summary>Window state properties.</summary>
    public class WindowStateSnapshot
    {
        public double Zoom { get; set; }
        public int View { get; set; }
        public int ScrollRow { get; set; }
        public int ScrollColumn { get; set; }
        public string? VisibleRange { get; set; }
        public bool DisplayPageBreaks { get; set; }
        public bool DisplayGridlines { get; set; }
        public bool DisplayHeadings { get; set; }
        public bool DisplayZeros { get; set; }
    }

    /// <summary>Shape/object bounds.</summary>
    public class ShapeSnapshot
    {
        public string? Name { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool Visible { get; set; }
        public int Type { get; set; }
        public string? TypeName { get; set; }
        public bool PrintObject { get; set; }
        public int ZOrderPosition { get; set; }
    }

    /// <summary>Printer DC metrics.</summary>
    public class PrinterStateSnapshot
    {
        public string? ActivePrinter { get; set; }
        public string? PrinterDriver { get; set; }

        // Win32 GetDeviceCaps
        public int LogicalPixelsX { get; set; }
        public int LogicalPixelsY { get; set; }
        public int HorizontalResolution { get; set; }
        public int VerticalResolution { get; set; }
        public int PhysicalWidth { get; set; }
        public int PhysicalHeight { get; set; }
        public int PhysicalOffsetX { get; set; }
        public int PhysicalOffsetY { get; set; }

        // Converted to points
        public double HardMarginLeftPt { get; set; }
        public double HardMarginTopPt { get; set; }
    }

    /// <summary>Per-field cell geometry at the moment of capture.</summary>
    public class FieldCellSnapshot
    {
        public string? CellAddress { get; set; }
        public string? FieldType { get; set; }

        // Cell-level geometry
        public double CellLeft { get; set; }
        public double CellTop { get; set; }
        public double CellWidth { get; set; }
        public double CellHeight { get; set; }

        // MergeArea geometry (may differ from cell if merged)
        public bool IsMerged { get; set; }
        public string? MergeAddress { get; set; }
        public double MergeLeft { get; set; }
        public double MergeTop { get; set; }
        public double MergeWidth { get; set; }
        public double MergeHeight { get; set; }

        // EntireColumn/Row
        public double EntireColumnLeft { get; set; }
        public double EntireColumnWidth { get; set; }
        public double EntireRowTop { get; set; }
        public double EntireRowHeight { get; set; }

        // Range-level
        public string? CurrentRegionAddress { get; set; }
    }
}
