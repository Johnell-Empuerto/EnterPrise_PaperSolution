using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Implements the verified legacy ConMas coordinate generation algorithm.
    /// 
    /// Transforms Excel COM worksheet geometry into page-relative coordinates
    /// using a single deterministic source (Excel COM) and the verified formula:
    ///   origin = margin + (printable - content_width) / 2
    ///   pixel = origin + (cellPt - printAreaPt) * scale
    ///
    /// Phase 14: Added ConMas formula with effective dimensions for correct centering.
    /// The effective dimensions (effW, effH) are measured from the rendered PNG content
    /// bounding box, giving the same centering as the legacy ConMas system.
    ///
    /// Reference: comprehensive_forensics_report.md — Q4 Hidden Transformations
    ///            report-conmas-background-generation.md — Step 14 Coordinate Origin
    ///            Investigation_546/com_pipeline_dump.json — ConMas formula verification
    /// </summary>
    public class CoordinateTransformer
    {
        private readonly ILogger<CoordinateTransformer> _logger;

        public CoordinateTransformer(ILogger<CoordinateTransformer> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Compute the printed page origin in points using the legacy centering formula.
        /// origin = margin + (printable - content_width) / 2
        /// </summary>
        /// <param name="pageWidthPt">Total page width in points.</param>
        /// <param name="pageHeightPt">Total page height in points.</param>
        /// <param name="leftMarginPt">Left margin in points.</param>
        /// <param name="rightMarginPt">Right margin in points.</param>
        /// <param name="topMarginPt">Top margin in points.</param>
        /// <param name="bottomMarginPt">Bottom margin in points.</param>
        /// <param name="contentWidthPt">Content (print area) width in points.</param>
        /// <param name="contentHeightPt">Content (print area) height in points.</param>
        /// <param name="centerHorizontally">Whether to center horizontally.</param>
        /// <param name="centerVertically">Whether to center vertically.</param>
        /// <returns>(originXPt, originYPt) in points.</returns>
        public (double originXPt, double originYPt) ComputePrintedOrigin(
            double pageWidthPt, double pageHeightPt,
            double leftMarginPt, double rightMarginPt,
            double topMarginPt, double bottomMarginPt,
            double contentWidthPt, double contentHeightPt,
            bool centerHorizontally, bool centerVertically)
        {
            double printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt;
            double printableHeightPt = pageHeightPt - topMarginPt - bottomMarginPt;

            double originXPt = leftMarginPt;
            double originYPt = topMarginPt;

            if (centerHorizontally && contentWidthPt < printableWidthPt)
                originXPt += (printableWidthPt - contentWidthPt) / 2.0;

            if (centerVertically && contentHeightPt < printableHeightPt)
                originYPt += (printableHeightPt - contentHeightPt) / 2.0;

            _logger.LogInformation(
                "[COORDS] Origin: ({OX:F2},{OY:F2})pt | page={PW:F1}x{PH:F1} " +
                "margin=({LM:F1},{RM:F1},{TM:F1},{BM:F1}) " +
                "printable={PW:F1}x{PH:F1} content={CW:F1}x{CH:F1} centerH={CH} centerV={CV}",
                originXPt, originYPt,
                pageWidthPt, pageHeightPt,
                leftMarginPt, rightMarginPt, topMarginPt, bottomMarginPt,
                printableWidthPt, printableHeightPt,
                contentWidthPt, contentHeightPt,
                centerHorizontally, centerVertically);

            return (originXPt, originYPt);
        }

        /// <summary>
        /// Compute the printed page origin using EFFECTIVE content dimensions
        /// measured from the rendered PNG (ConMas formula).
        /// 
        /// This matches the legacy ConMas system where the content bounding box
        /// on the printed page differs from COM Range.Width/Height due to
        /// Excel's page layout engine (FitToPages, font metrics, etc.).
        /// 
        /// Formula: origin = margin + (printable - effective) / 2
        /// </summary>
        public (double originXPt, double originYPt) ComputePrintedOriginFromEffective(
            double pageWidthPt, double pageHeightPt,
            double leftMarginPt, double rightMarginPt,
            double topMarginPt, double bottomMarginPt,
            double effectiveWidthPt, double effectiveHeightPt)
        {
            double printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt;
            double printableHeightPt = pageHeightPt - topMarginPt - bottomMarginPt;

            double originXPt = leftMarginPt + (printableWidthPt - effectiveWidthPt) / 2.0;
            double originYPt = topMarginPt + (printableHeightPt - effectiveHeightPt) / 2.0;

            _logger.LogInformation(
                "[CONMAS] Effective origin: ({OX:F2},{OY:F2})pt | " +
                "effW={EW:F2}pt effH={EH:F2}pt | printable={PW:F1}x{PH:F1}",
                originXPt, originYPt,
                effectiveWidthPt, effectiveHeightPt,
                printableWidthPt, printableHeightPt);

            return (originXPt, originYPt);
        }

        /// <summary>
        /// Compute page-relative cell position using the ConMas formula.
        /// This applies the effective-dimension scaling to each cell:
        ///   pagePt = originPt + (cellPt - printAreaPt) * (effectiveDim / rangeDim)
        /// 
        /// The (effectiveDim / rangeDim) factor accounts for Excel's page layout
        /// scaling when FitToPages or other non-100% zoom settings are active.
        /// </summary>
        public double CellToPagePt(
            double cellPt, double printAreaPt,
            double originPt, double effectiveDim, double rangeDim)
        {
            if (rangeDim <= 0 || effectiveDim <= 0)
                return originPt + (cellPt - printAreaPt);

            double scale = effectiveDim / rangeDim;
            return originPt + (cellPt - printAreaPt) * scale;
        }

        /// <summary>
        /// Convert a cell position from worksheet-relative points to page-relative pixels.
        /// pixel = (originPt + (cellPt - printAreaPt)) * scale
        /// </summary>
        public double CellToPixel(double cellPt, double printAreaPt, double originPt, double scale)
        {
            return originPt + (cellPt - printAreaPt) * scale;
        }

        /// <summary>
        /// Convert a point dimension to pixels at the given scale.
        /// </summary>
        public double PtToPx(double pt, double scale) => pt * scale;

        /// <summary>
        /// Compute point-to-pixel scale from actual rendered PNG dimensions.
        /// Falls back to canonical Rectangle.PtToPx scale factor if PNG dimensions are unavailable.
        /// </summary>
        public double ComputeScale(int pngDimension, double pageDimensionPt, double dpi)
        {
            return pngDimension > 0 && pageDimensionPt > 0
                ? pngDimension / pageDimensionPt
                : WbDef.Rectangle.PtToPx(1, dpi);
        }
    }
}
