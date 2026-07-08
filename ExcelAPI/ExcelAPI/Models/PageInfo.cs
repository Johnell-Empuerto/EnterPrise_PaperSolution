namespace ExcelAPI.Models
{
    /// <summary>
    /// Represents the dimensions of the generated PNG image (at 300 DPI).
    /// Used by the frontend to correctly position field overlays.
    /// </summary>
    public class PageInfo
    {
        /// <summary>PNG image width in pixels.</summary>
        public int Width { get; set; }

        /// <summary>PNG image height in pixels.</summary>
        public int Height { get; set; }
    }
}
