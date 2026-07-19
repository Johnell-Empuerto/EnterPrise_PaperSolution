// ────────────────────────────────────────────────────────────────────────────
// ImageDefinition and ShapeDefinition — Visual Elements
//
// Images are embedded pictures from xl/media/. Shapes are DrawingML objects
// (rectangles, text boxes, arrows, etc.) from xl/drawings/.
//
// Ownership: Shared (Designer → Rendering + Runtime)
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// An embedded image (picture) on a worksheet.
    /// Maps to a drawing object with an image relationship.
    /// </summary>
    public class ImageDefinition
    {
        /// <summary>Unique image identifier within the workbook.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>Image name / title from DrawingML.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Content type (e.g., "image/png", "image/jpeg", "image/gif").</summary>
        public string ContentType { get; set; } = "image/png";

        /// <summary>Raw image file data.</summary>
        public byte[]? Data { get; set; }

        /// <summary>Image bounds in points (from DrawingML anchor).</summary>
        public Rectangle BoundsPt { get; set; } = new();

        /// <summary>Pixel width of the decoded image.</summary>
        public int PixelWidth { get; set; }

        /// <summary>Pixel height of the decoded image.</summary>
        public int PixelHeight { get; set; }

        /// <summary>
        /// Whether the image is decorative (not an interactive element).
        /// </summary>
        public bool IsDecorative { get; set; } = true;
    }

    /// <summary>
    /// A drawing shape from DrawingML (xl/drawings/).
    /// Can be a rectangle, text box, arrow, ellipse, line, grouped shape, etc.
    /// </summary>
    public class ShapeDefinition
    {
        /// <summary>Unique shape identifier within the workbook.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>Shape name from DrawingML (may be empty).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Shape type: "rectangle", "roundedRect", "ellipse", "line", "arrow",
        /// "textBox", "polygon", "freeform", "group", "picture".
        /// </summary>
        public string ShapeType { get; set; } = "rectangle";

        /// <summary>Shape bounds in points (from DrawingML anchor, worksheet-relative).</summary>
        public Rectangle BoundsPt { get; set; } = new();

        /// <summary>Rotation in degrees (from DrawingML transform).</summary>
        public double RotationDegrees { get; set; }

        /// <summary>Fill / background style.</summary>
        public FillDefinition? Fill { get; set; }

        /// <summary>Border / outline style.</summary>
        public BorderEdge? Border { get; set; }

        /// <summary>Text content (for textBox shapes).</summary>
        public string? Text { get; set; }

        /// <summary>Text font properties.</summary>
        public FontDefinition? Font { get; set; }

        /// <summary>Text alignment within the shape.</summary>
        public AlignmentDefinition? Alignment { get; set; }

        /// <summary>Whether text wraps within the shape.</summary>
        public bool WrapText { get; set; }
    }
}
