namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Orchestrates the full print layout computation for a sheet.
    /// Combines PageGeometryResolver + MarginResolver + ScalingResolver
    /// + origin calculation into a single authoritative pipeline.
    ///
    /// Every render layer consumes the output via RenderingContext.
    /// No layer independently computes margins, origin, or scaling.
    /// </summary>
    public class PrintLayoutEngine
    {
        private readonly PageGeometryResolver _pageGeometry;
        private readonly MarginResolver _margins;
        private readonly ScalingResolver _scaling;

        public PrintLayoutEngine(
            PageGeometryResolver pageGeometry,
            MarginResolver margins,
            ScalingResolver scaling)
        {
            _pageGeometry = pageGeometry;
            _margins = margins;
            _scaling = scaling;
        }

        /// <summary>
        /// Compute the complete print layout for a sheet,
        /// given page settings and computed content dimensions.
        /// Produces all the fields needed to populate a RenderingContext.
        /// </summary>
        public PrintLayoutResult Compute(
            string paperSize,
            string orientation,
            double leftMargin,
            double rightMargin,
            double topMargin,
            double bottomMargin,
            bool centerHorizontally,
            bool centerVertically,
            int zoom,
            int fitToPagesWide,
            int fitToPagesTall,
            double totalContentWidthPt,
            double totalContentHeightPt)
        {
            // Step 1: Page geometry
            var pageGeo = _pageGeometry.Resolve(paperSize, orientation);

            // Step 2: Margins
            var marginInfo = _margins.Resolve(leftMargin, rightMargin, topMargin, bottomMargin);

            // Step 3: Printable area
            double printableWidthPt = marginInfo.PrintableWidthPt(pageGeo.WidthPt);
            double printableHeightPt = marginInfo.PrintableHeightPt(pageGeo.HeightPt);

            // Step 4: Scaling
            var scaleInfo = _scaling.Resolve(
                zoom, fitToPagesWide, fitToPagesTall,
                totalContentWidthPt, printableWidthPt,
                totalContentHeightPt, printableHeightPt);

            // Step 5: Printable origin (margin + centering)
            double originXPt = marginInfo.LeftPt;
            double originYPt = marginInfo.TopPt;

            // Content width/height after scaling
            double scaledContentWidthPt = totalContentWidthPt * scaleInfo.ScaleFactor;
            double scaledContentHeightPt = totalContentHeightPt * scaleInfo.ScaleFactor;

            if (centerHorizontally && scaledContentWidthPt < printableWidthPt)
                originXPt += (printableWidthPt - scaledContentWidthPt) / 2.0;

            if (centerVertically && scaledContentHeightPt < printableHeightPt)
                originYPt += (printableHeightPt - scaledContentHeightPt) / 2.0;

            // Step 6: Clip region (content within printable area)
            double clipLeftPt = originXPt;
            double clipTopPt = originYPt;
            double clipRightPt = Math.Min(originXPt + scaledContentWidthPt, marginInfo.LeftPt + printableWidthPt);
            double clipBottomPt = Math.Min(originYPt + scaledContentHeightPt, marginInfo.TopPt + printableHeightPt);

            return new PrintLayoutResult
            {
                PaperSize = pageGeo.PaperSize,
                Orientation = pageGeo.Orientation,
                PageWidthPt = pageGeo.WidthPt,
                PageHeightPt = pageGeo.HeightPt,
                MarginLeftPt = marginInfo.LeftPt,
                MarginRightPt = marginInfo.RightPt,
                MarginTopPt = marginInfo.TopPt,
                MarginBottomPt = marginInfo.BottomPt,
                PrintableWidthPt = printableWidthPt,
                PrintableHeightPt = printableHeightPt,
                OriginXPt = originXPt,
                OriginYPt = originYPt,
                ScaleFactor = scaleInfo.ScaleFactor,
                Zoom = scaleInfo.Zoom,
                FitToPagesWide = scaleInfo.FitToPagesWide,
                FitToPagesTall = scaleInfo.FitToPagesTall,
                IsScalingActive = scaleInfo.IsScalingActive,
                ClipLeftPt = clipLeftPt,
                ClipTopPt = clipTopPt,
                ClipRightPt = clipRightPt,
                ClipBottomPt = clipBottomPt
            };
        }
    }

    /// <summary>
    /// Complete result of print layout computation.
    /// All values consumed directly by RenderingContext.
    /// </summary>
    public class PrintLayoutResult
    {
        public string PaperSize { get; init; } = "Letter";
        public string Orientation { get; init; } = "portrait";
        public double PageWidthPt { get; init; }
        public double PageHeightPt { get; init; }
        public double MarginLeftPt { get; init; }
        public double MarginRightPt { get; init; }
        public double MarginTopPt { get; init; }
        public double MarginBottomPt { get; init; }
        public double PrintableWidthPt { get; init; }
        public double PrintableHeightPt { get; init; }
        public double OriginXPt { get; init; }
        public double OriginYPt { get; init; }
        public double ScaleFactor { get; init; } = 1.0;
        public int Zoom { get; init; } = 100;
        public int FitToPagesWide { get; init; }
        public int FitToPagesTall { get; init; }
        public bool IsScalingActive { get; init; }
        public double ClipLeftPt { get; init; }
        public double ClipTopPt { get; init; }
        public double ClipRightPt { get; init; }
        public double ClipBottomPt { get; init; }
    }
}
