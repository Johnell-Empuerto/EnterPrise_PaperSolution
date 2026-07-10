namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Resolves paper size and orientation to page dimensions in points (1/72 inch).
    /// Single authoritative source for paper geometry — no layer recalculates this.
    /// </summary>
    public class PageGeometryResolver
    {
        /// <summary>
        /// Result of page geometry resolution.
        /// </summary>
        public record PageGeometry(
            string PaperSize,
            string Orientation,
            double WidthPt,
            double HeightPt);

        /// <summary>
        /// Resolve page dimensions from paper size name and orientation.
        /// </summary>
        public PageGeometry Resolve(string? paperSize, string? orientation)
        {
            string ps = (paperSize ?? "Letter").Trim();
            string ori = (orientation ?? "portrait").Trim().ToLowerInvariant();

            (double w, double h) = ps.ToLowerInvariant() switch
            {
                "letter" => (612.0, 792.0),         // 8.5 x 11 in
                "a4" => (595.0, 842.0),             // 210 x 297 mm
                "legal" => (612.0, 1008.0),         // 8.5 x 14 in
                "a3" => (842.0, 1191.0),            // 297 x 420 mm
                "a5" => (420.0, 595.0),             // 148 x 210 mm
                "b5" => (499.0, 709.0),             // 176 x 250 mm
                "executive" => (522.0, 756.0),      // 7.25 x 10.5 in
                "tabloid" => (792.0, 1224.0),       // 11 x 17 in
                "ledger" => (1224.0, 792.0),        // 17 x 11 in
                "envelope" or "envelope10" => (684.0, 360.0), // 9.5 x 4.125 in
                _ => (612.0, 792.0)                 // Default to Letter
            };

            // Swap for landscape
            if (ori == "landscape")
                return new PageGeometry(ps, ori, Math.Max(w, h), Math.Min(w, h));

            return new PageGeometry(ps, ori, Math.Min(w, h), Math.Max(w, h));
        }
    }
}
