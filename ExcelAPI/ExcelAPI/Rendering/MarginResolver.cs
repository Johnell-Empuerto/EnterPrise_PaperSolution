namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Resolves page margins.
    /// Single authoritative source for margin geometry — no layer recalculates this.
    /// </summary>
    public class MarginResolver
    {
        /// <summary>
        /// Resolved page margins in points.
        /// </summary>
        public record PageMargins(
            double LeftPt,
            double RightPt,
            double TopPt,
            double BottomPt)
        {
            /// <summary>Printable width = page width - left - right margin.</summary>
            public double PrintableWidthPt(double pageWidthPt) =>
                pageWidthPt - LeftPt - RightPt;

            /// <summary>Printable height = page height - top - bottom margin.</summary>
            public double PrintableHeightPt(double pageHeightPt) =>
                pageHeightPt - TopPt - BottomPt;
        }

        /// <summary>
        /// Resolve the effective margins from page setup values.
        /// Uses 51.02pt (ConMas default) when a margin is zero or unset.
        /// </summary>
        public PageMargins Resolve(
            double leftMargin,
            double rightMargin,
            double topMargin,
            double bottomMargin)
        {
            const double defaultMargin = 51.02; // ConMas default margin

            return new PageMargins(
                LeftPt: leftMargin > 0 ? leftMargin : defaultMargin,
                RightPt: rightMargin > 0 ? rightMargin : defaultMargin,
                TopPt: topMargin > 0 ? topMargin : defaultMargin,
                BottomPt: bottomMargin > 0 ? bottomMargin : defaultMargin);
        }
    }
}
