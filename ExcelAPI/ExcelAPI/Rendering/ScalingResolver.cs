namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Resolves the effective rendering scale factor from Zoom and FitToPages settings.
    ///
    /// Excel scaling priority:
    ///   1. FitToPagesWide/Tall — scale content to fit N pages wide x M pages tall
    ///   2. Zoom percentage (100 = 100%)
    ///   3. No scaling (scale = 1.0)
    ///
    /// The scale factor affects ALL rendering layers equally (fills, borders, text).
    /// </summary>
    public class ScalingResolver
    {
        /// <summary>
        /// Resolved scaling information.
        /// </summary>
        public record ScalingInfo(
            /// <summary>Scale factor: 1.0 = 100%, 0.5 = 50%, 2.0 = 200%.</summary>
            double ScaleFactor,
            /// <summary>Zoom percentage (0 if using FitToPages).</summary>
            int Zoom,
            /// <summary>Fit to pages wide (0 if not set).</summary>
            int FitToPagesWide,
            /// <summary>Fit to pages tall (0 if not set).</summary>
            int FitToPagesTall,
            /// <summary>Whether scaling is active (scale ≠ 1.0).</summary>
            bool IsScalingActive)
        {
            /// <summary>Effective DPI when scaling is applied.</summary>
            public double EffectiveDpi(double baseDpi) => baseDpi * ScaleFactor;
        }

        /// <summary>
        /// Resolve the effective scale factor.
        /// </summary>
        /// <param name="zoom">Zoom percentage (0 if FitToPages is active, 100 for default).</param>
        /// <param name="fitToPagesWide">Fit to N pages wide (0 if not set).</param>
        /// <param name="fitToPagesTall">Fit to N pages tall (0 if not set).</param>
        /// <param name="contentWidthPt">Total content width in points.</param>
        /// <param name="printableWidthPt">Printable page width in points.</param>
        /// <param name="contentHeightPt">Total content height in points.</param>
        /// <param name="printableHeightPt">Printable page height in points.</param>
        public ScalingInfo Resolve(
            int zoom, int fitToPagesWide, int fitToPagesTall,
            double contentWidthPt, double printableWidthPt,
            double contentHeightPt, double printableHeightPt)
        {
            // Priority 1: FitToPages
            if (fitToPagesWide > 0 || fitToPagesTall > 0)
            {
                double scaleW = 1.0;
                double scaleH = 1.0;

                if (fitToPagesWide > 0 && printableWidthPt > 0)
                    scaleW = printableWidthPt / (contentWidthPt * fitToPagesWide);

                if (fitToPagesTall > 0 && printableHeightPt > 0)
                    scaleH = printableHeightPt / (contentHeightPt * fitToPagesTall);

                // Excel uses the smaller of the two to ensure content fits
                double scale = Math.Min(scaleW, scaleH);
                scale = Math.Clamp(scale, 0.1, 10.0); // Reasonable bounds

                return new ScalingInfo(scale, 0, fitToPagesWide, fitToPagesTall, Math.Abs(scale - 1.0) > 0.001);
            }

            // Priority 2: Custom Zoom
            if (zoom > 0 && zoom != 100)
            {
                double scale = zoom / 100.0;
                return new ScalingInfo(scale, zoom, 0, 0, true);
            }

            // Priority 3: No scaling
            return new ScalingInfo(1.0, 100, 0, 0, false);
        }
    }
}
