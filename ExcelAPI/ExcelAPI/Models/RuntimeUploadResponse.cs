namespace ExcelAPI.Models
{
    /// <summary>
    /// Clean response model returned by POST /api/runtime/upload.
    /// Contains the rendered PNG pages and overlay definitions for interactive fields.
    ///
    /// This is the primary API contract between the PaperLess backend (COM engine)
    /// and the frontend (Next.js runtime viewer).
    /// </summary>
    public class RuntimeUploadResponse
    {
        /// <summary>List of rendered pages (one per print area page / worksheet).</summary>
        public List<RuntimePageInfo> Pages { get; set; } = new();

        /// <summary>List of overlay definitions (one per Excel comment).</summary>
        public List<RuntimeOverlayInfo> Overlays { get; set; } = new();
    }

    /// <summary>
    /// A single rendered page from the Excel print area.
    /// </summary>
    public class RuntimePageInfo
    {
        /// <summary>Worksheet name (e.g., "Sheet1", "Invoice").</summary>
        public string SheetName { get; set; } = "";

        /// <summary>Relative URL to the background PNG image.</summary>
        public string BackgroundImage { get; set; } = "";

        /// <summary>PNG width in pixels.</summary>
        public int Width { get; set; }

        /// <summary>PNG height in pixels.</summary>
        public int Height { get; set; }
    }

    /// <summary>
    /// An interactive overlay field positioned on a rendered page.
    /// </summary>
    public class RuntimeOverlayInfo
    {
        /// <summary>Unique field identifier (e.g., "field_B5").</summary>
        public string Id { get; set; } = "";

        /// <summary>Worksheet this field belongs to.</summary>
        public string SheetName { get; set; } = "";

        /// <summary>Field type: "textbox", "checkbox", "signature", "date", "number".</summary>
        public string Type { get; set; } = "textbox";

        /// <summary>Left edge in pixels (relative to PNG top-left).</summary>
        public double Left { get; set; }

        /// <summary>Top edge in pixels (relative to PNG top-left).</summary>
        public double Top { get; set; }

        /// <summary>Width in pixels.</summary>
        public double Width { get; set; }

        /// <summary>Height in pixels.</summary>
        public double Height { get; set; }

        /// <summary>Original Excel cell reference (e.g., "B5").</summary>
        public string Cell { get; set; } = "";

        /// <summary>Whether this cell is part of a merged range.</summary>
        public bool IsMerged { get; set; }

        /// <summary>Merge range address if merged (e.g., "A1:B2").</summary>
        public string? MergeAddress { get; set; }
    }
}
