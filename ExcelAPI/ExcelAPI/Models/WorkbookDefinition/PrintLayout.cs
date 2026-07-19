using System.Text.Json.Serialization;

// ────────────────────────────────────────────────────────────────────────────
// PrintLayout — Page Layout Definition
//
// Describes the print layout configuration of a worksheet: paper size,
// orientation, margins, scaling, print area, and centering.
//
// Ownership: Shared (populated by Designer, consumed by Rendering)
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// Complete print layout for a worksheet.
    /// Maps to Excel's PageSetup dialog.
    /// </summary>
    public class PrintLayout
    {
        /// <summary>Paper size definition.</summary>
        public PaperSize PaperSize { get; set; } = new();

        /// <summary>Page orientation.</summary>
        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

        /// <summary>Print area bounds (null = entire sheet is print area).</summary>
        public PrintAreaDefinition? PrintArea { get; set; }

        /// <summary>Page margins.</summary>
        public Margins Margins { get; set; } = new();

        /// <summary>Scaling / zoom configuration.</summary>
        public ScalingDefinition Scaling { get; set; } = new();

        /// <summary>Rendering DPI (typically 300 for production, 96 for screen).
        /// Controls the pixel resolution when this layout is rendered.</summary>
        public int Dpi { get; set; } = 300;

        /// <summary>Page width in points (from PaperSize × Orientation).</summary>
        public double PageWidthPt => Orientation == PageOrientation.Landscape
            ? PaperSize.HeightPt
            : PaperSize.WidthPt;

        /// <summary>Page height in points (from PaperSize × Orientation).</summary>
        public double PageHeightPt => Orientation == PageOrientation.Landscape
            ? PaperSize.WidthPt
            : PaperSize.HeightPt;
    }

    /// <summary>
    /// Paper size definition. Supports standard sizes and custom dimensions.
    /// </summary>
    public class PaperSize
    {
        /// <summary>Human-readable name (e.g., "Letter", "A4", "Legal", "Custom").</summary>
        public string Name { get; set; } = "Letter";

        /// <summary>Excel COM PaperSize enum value (1 = Letter, 5 = Legal, 9 = A4, etc.).</summary>
        public int ExcelCode { get; set; } = 1;

        /// <summary>Paper width in points (portrait orientation).</summary>
        public double WidthPt { get; set; } = 612.0;

        /// <summary>Paper height in points (portrait orientation).</summary>
        public double HeightPt { get; set; } = 792.0;
    }

    /// <summary>
    /// Page orientation.
    /// </summary>
    public enum PageOrientation
    {
        Portrait,
        Landscape
    }

    /// <summary>
    /// Print area bounds on the worksheet.
    /// </summary>
    public class PrintAreaDefinition
    {
        /// <summary>A1-style address (e.g., "A1:H40").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>The cell range of the print area.</summary>
        public CellRange? Range { get; set; }

        /// <summary>Print area bounds in points (worksheet-relative).</summary>
        public Rectangle? BoundsPt { get; set; }
    }

    /// <summary>
    /// Page margins in points (1 point = 1/72 inch).
    /// </summary>
    public class Margins
    {
        /// <summary>Left margin in points. Excel default = 0.7in = 50.4pt.</summary>
        public double LeftPt { get; set; } = 50.4;

        /// <summary>Right margin in points. Excel default = 0.7in = 50.4pt.</summary>
        public double RightPt { get; set; } = 50.4;

        /// <summary>Top margin in points. Excel default = 0.75in = 54pt.</summary>
        public double TopPt { get; set; } = 54.0;

        /// <summary>Bottom margin in points. Excel default = 0.75in = 54pt.</summary>
        public double BottomPt { get; set; } = 54.0;

        /// <summary>Header margin in points. Excel default = 0.5in = 36pt.</summary>
        public double HeaderPt { get; set; } = 36.0;

        /// <summary>Footer margin in points. Excel default = 0.5in = 36pt.</summary>
        public double FooterPt { get; set; } = 36.0;
    }

    /// <summary>
    /// Scaling and centering configuration for printing.
    /// </summary>
    public class ScalingDefinition
    {
        /// <summary>Zoom percentage (10–400, 0 means FitToPages is active).</summary>
        public int Zoom { get; set; } = 100;

        /// <summary>Fit to N pages wide (0 = not set, meaning no fit constraint).</summary>
        public int FitToPagesWide { get; set; }

        /// <summary>Fit to N pages tall (0 = not set).</summary>
        public int FitToPagesTall { get; set; }

        /// <summary>Whether to center horizontally on the page.</summary>
        public bool CenterHorizontally { get; set; }

        /// <summary>Whether to center vertically on the page.</summary>
        public bool CenterVertically { get; set; }

        /// <summary>
        /// Whether FitToPages scaling is active (either dimension set).
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public bool IsFitToPagesActive => FitToPagesWide > 0 || FitToPagesTall > 0;

        /// <summary>
        /// Whether any scaling is active (non-default zoom or FitToPages).
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public bool IsScalingActive => Zoom != 100 || IsFitToPagesActive;

        /// <summary>
        /// Effective scale factor (1.0 = 100%, 0.5 = 50%).
        /// Returns 0 when FitToPages is active (actual scale depends on content).
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double ScaleFactor => IsFitToPagesActive ? 0.0 : Zoom / 100.0;
    }
}
