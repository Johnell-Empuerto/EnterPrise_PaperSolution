using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Splits a worksheet into printable pages and manages page-level rendering.
    /// Respects PrintArea, margins, scaling, FitToPage, page breaks, centering,
    /// and orientation.
    ///
    /// Each page gets its own RenderingContext with appropriate clipping
    /// and content offset for that page region.
    ///
    /// Supports unlimited printable pages.
    /// </summary>
    public class PageRenderer
    {
        private readonly PrintLayoutEngine _printLayout;
        private readonly GeometryBuilder _geometry;

        public PageRenderer(PrintLayoutEngine printLayout, GeometryBuilder geometry)
        {
            _printLayout = printLayout;
            _geometry = geometry;
        }

        /// <summary>
        /// Describes a single rendered page.
        /// </summary>
        public class PageInfo
        {
            /// <summary>1-based page number.</summary>
            public int PageNumber { get; init; }

            /// <summary>Total page count.</summary>
            public int TotalPages { get; init; }

            /// <summary>Rendering context for this page (may have adjusted origin/clip).</summary>
            public RenderingContext Context { get; init; } = new();

            /// <summary>Pixel width of the output canvas for this page.</summary>
            public int PixelWidth { get; init; }

            /// <summary>Pixel height of the output canvas for this page.</summary>
            public int PixelHeight { get; init; }
        }

        /// <summary>
        /// Compute all pages for a sheet given page settings and content dimensions.
        /// Returns a list of PageInfo objects, one per printable page.
        /// </summary>
        public List<PageInfo> ComputePages(
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
            double totalContentHeightPt,
            int dpi,
            int maxPages = 0)
        {
            // Compute base layout for the first page
            var layout = _printLayout.Compute(
                paperSize, orientation,
                leftMargin, rightMargin, topMargin, bottomMargin,
                centerHorizontally, centerVertically,
                zoom, fitToPagesWide, fitToPagesTall,
                totalContentWidthPt, totalContentHeightPt);

            // Determine page dimensions in points
            double pageWidthPt = layout.PageWidthPt;
            double pageHeightPt = layout.PageHeightPt;
            double printableWidthPt = layout.PrintableWidthPt;
            double printableHeightPt = layout.PrintableHeightPt;
            double ptsToPx = dpi / 72.0;

            // Calculate the effective content area per page (after scaling)
            double scaledContentWidthPt = totalContentWidthPt * layout.ScaleFactor;
            double scaledContentHeightPt = totalContentHeightPt * layout.ScaleFactor;

            // Determine number of pages
            int pagesWide = fitToPagesWide > 0
                ? fitToPagesWide
                : Math.Max(1, (int)Math.Ceiling(scaledContentWidthPt / printableWidthPt));

            int pagesTall = fitToPagesTall > 0
                ? fitToPagesTall
                : Math.Max(1, (int)Math.Ceiling(scaledContentHeightPt / printableHeightPt));

            int totalPages = pagesWide * pagesTall;
            if (maxPages > 0)
                totalPages = Math.Min(totalPages, maxPages);

            var pages = new List<PageInfo>();

            for (int pageY = 0; pageY < pagesTall && pages.Count < (maxPages > 0 ? maxPages : int.MaxValue); pageY++)
            {
                for (int pageX = 0; pageX < pagesWide && pages.Count < (maxPages > 0 ? maxPages : int.MaxValue); pageX++)
                {
                    int pageNum = pageY * pagesWide + pageX + 1;

                    // Compute the content offset for this page (which portion of the sheet)
                    double pageContentLeft = layout.OriginXPt + pageX * printableWidthPt;
                    double pageContentTop = layout.OriginYPt + pageY * printableHeightPt;

                    // Clip region for this page (within printable area)
                    double clipLeftPt = layout.OriginXPt + pageX * printableWidthPt;
                    double clipTopPt = layout.OriginYPt + pageY * printableHeightPt;
                    double clipRightPt = Math.Min(clipLeftPt + printableWidthPt,
                        layout.OriginXPt + scaledContentWidthPt);
                    double clipBottomPt = Math.Min(clipTopPt + printableHeightPt,
                        layout.OriginYPt + scaledContentHeightPt);

                    var context = new RenderingContext
                    {
                        PaperSize = layout.PaperSize,
                        Orientation = layout.Orientation,
                        PageWidthPt = pageWidthPt,
                        PageHeightPt = pageHeightPt,
                        MarginLeftPt = layout.MarginLeftPt,
                        MarginRightPt = layout.MarginRightPt,
                        MarginTopPt = layout.MarginTopPt,
                        MarginBottomPt = layout.MarginBottomPt,
                        PrintableWidthPt = printableWidthPt,
                        PrintableHeightPt = printableHeightPt,
                        OriginXPt = pageContentLeft,
                        OriginYPt = pageContentTop,
                        ScaleFactor = layout.ScaleFactor,
                        Zoom = layout.Zoom,
                        FitToPagesWide = layout.FitToPagesWide,
                        FitToPagesTall = layout.FitToPagesTall,
                        IsScalingActive = layout.IsScalingActive,
                        ClipLeftPt = clipLeftPt,
                        ClipTopPt = clipTopPt,
                        ClipRightPt = clipRightPt,
                        ClipBottomPt = clipBottomPt,
                        Dpi = dpi
                    };

                    pages.Add(new PageInfo
                    {
                        PageNumber = pageNum,
                        TotalPages = totalPages,
                        Context = context,
                        PixelWidth = (int)Math.Round(pageWidthPt * ptsToPx),
                        PixelHeight = (int)Math.Round(pageHeightPt * ptsToPx)
                    });
                }
            }

            return pages;
        }
    }
}
