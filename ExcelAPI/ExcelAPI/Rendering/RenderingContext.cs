using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Shared rendering context passed to every IRenderLayer.
    /// Every layer receives the same context — never recalculated.
    ///
    /// Contains pre-computed layout geometry from PrintLayoutEngine and CoordinateEngine.
    /// No renderer should independently compute margins, origin, or scaling.
    /// </summary>
    public class RenderingContext
    {
        /// <summary>The parsed workbook model (cells, columns, rows, merges, styles).</summary>
        public RenderWorkbook Workbook { get; set; } = new();

        /// <summary>The current sheet being rendered.</summary>
        public RenderSheet Sheet { get; set; } = new();

        // ── Page Geometry ────────────────────────────────────────────────

        /// <summary>Paper size name (e.g., "Letter", "A4").</summary>
        public string PaperSize { get; init; } = "Letter";

        /// <summary>Page orientation ("portrait" or "landscape").</summary>
        public string Orientation { get; init; } = "portrait";

        /// <summary>Page width in points.</summary>
        public double PageWidthPt { get; init; } = 612;

        /// <summary>Page height in points.</summary>
        public double PageHeightPt { get; init; } = 792;

        // ── Margins ─────────────────────────────────────────────────────

        /// <summary>Left margin in points.</summary>
        public double MarginLeftPt { get; init; }

        /// <summary>Right margin in points.</summary>
        public double MarginRightPt { get; init; }

        /// <summary>Top margin in points.</summary>
        public double MarginTopPt { get; init; }

        /// <summary>Bottom margin in points.</summary>
        public double MarginBottomPt { get; init; }

        // ── Printable Area ──────────────────────────────────────────────

        /// <summary>Printable width in points (page width - left - right margins).</summary>
        public double PrintableWidthPt { get; init; }

        /// <summary>Printable height in points (page height - top - bottom margins).</summary>
        public double PrintableHeightPt { get; init; }

        // ── Content Origin ──────────────────────────────────────────────

        /// <summary>Printable origin X in points (margin left + centering offset).</summary>
        public double OriginXPt { get; init; }

        /// <summary>Printable origin Y in points (margin top + centering offset).</summary>
        public double OriginYPt { get; init; }

        // ── Scaling ─────────────────────────────────────────────────────

        /// <summary>Effective scale factor (1.0 = 100%, 0.5 = 50%).</summary>
        public double ScaleFactor { get; init; } = 1.0;

        /// <summary>Zoom percentage (100 = default, 0 if FitToPages is active).</summary>
        public int Zoom { get; init; } = 100;

        /// <summary>Fit to N pages wide (0 if not set).</summary>
        public int FitToPagesWide { get; init; }

        /// <summary>Fit to N pages tall (0 if not set).</summary>
        public int FitToPagesTall { get; init; }

        /// <summary>Whether scaling is active (ScaleFactor ≠ 1.0).</summary>
        public bool IsScalingActive { get; init; }

        // ── Clip Region ─────────────────────────────────────────────────

        /// <summary>Clip region left in points (page-relative).</summary>
        public double ClipLeftPt { get; init; }

        /// <summary>Clip region top in points (page-relative).</summary>
        public double ClipTopPt { get; init; }

        /// <summary>Clip region right in points (page-relative).</summary>
        public double ClipRightPt { get; init; }

        /// <summary>Clip region bottom in points (page-relative).</summary>
        public double ClipBottomPt { get; init; }

        /// <summary>Clip region as SkiaSharp SKRect in pixels.</summary>
        public SKRect ClipRectPx => new(
            (float)(ClipLeftPt * PointsToPixels),
            (float)(ClipTopPt * PointsToPixels),
            (float)(ClipRightPt * PointsToPixels),
            (float)(ClipBottomPt * PointsToPixels));

        // ── Rendering Constants ─────────────────────────────────────────

        /// <summary>Rendering DPI (typically 300).</summary>
        public double Dpi { get; init; } = 300.0;

        /// <summary>Computed points-to-pixels multiplier (Dpi / 72).</summary>
        public double PointsToPixels => Dpi / 72.0;

        /// <summary>Pixel width of the output canvas.</summary>
        public int PixelWidth => (int)Math.Round(PageWidthPt * PointsToPixels);

        /// <summary>Pixel height of the output canvas.</summary>
        public int PixelHeight => (int)Math.Round(PageHeightPt * PointsToPixels);
    }
}
