namespace ExcelAPI.Models
{
    /// <summary>
    /// Debug information about the page setup used for coordinate calculation.
    /// Returned temporarily to verify pixel-perfect alignment.
    /// </summary>
    public class PageSetupDebug
    {
        /// <summary>X origin of the printed content on the PNG page (in pixels).</summary>
        public double PrintedOriginX { get; set; }

        /// <summary>Y origin of the printed content on the PNG page (in pixels).</summary>
        public double PrintedOriginY { get; set; }

        /// <summary>Left page margin in points (1/72 inch).</summary>
        public double LeftMargin { get; set; }

        /// <summary>Top page margin in points (1/72 inch).</summary>
        public double TopMargin { get; set; }

        /// <summary>Whether content is centered horizontally on the page.</summary>
        public bool CenterHorizontally { get; set; }

        /// <summary>Whether content is centered vertically on the page.</summary>
        public bool CenterVertically { get; set; }

        /// <summary>Page width in points.</summary>
        public double PageWidthPt { get; set; }

        /// <summary>Page height in points.</summary>
        public double PageHeightPt { get; set; }

        /// <summary>Print area content width in points.</summary>
        public double PrintAreaWidthPt { get; set; }

        /// <summary>Print area content height in points.</summary>
        public double PrintAreaHeightPt { get; set; }

        /// <summary>Point-to-pixel scale factor (DPI / 72).</summary>
        public double Scale { get; set; }

        /// <summary>Actual X scale from rendered PNG (pngWidth / pageWidthPt).</summary>
        public double ActualScaleX { get; set; }

        /// <summary>Actual Y scale from rendered PNG (pngHeight / pageHeightPt).</summary>
        public double ActualScaleY { get; set; }

        /// <summary>Zoom setting (0 if using FitToPages).</summary>
        public int Zoom { get; set; }

        /// <summary>Fit to pages wide setting (0 if not set).</summary>
        public int FitToPagesWide { get; set; }

        /// <summary>Fit to pages tall setting (0 if not set).</summary>
        public int FitToPagesTall { get; set; }
    }
}
