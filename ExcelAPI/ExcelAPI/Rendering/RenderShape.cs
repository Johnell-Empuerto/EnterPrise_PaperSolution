namespace ExcelAPI.Rendering
{
    /// <summary>
    /// A resolved shape parsed from DrawingML (drawing.xml).
    /// Ready for consumption by ShapeEngine.
    /// </summary>
    public class RenderShape
    {
        /// <summary>Shape name from DrawingML (may be empty).</summary>
        public string Name { get; set; } = "";

        /// <summary>Shape type: rectangle, roundedRect, ellipse, line, arrow, textBox, polygon.</summary>
        public string ShapeType { get; set; } = "rectangle";

        // ── Geometry (worksheet-relative points, from anchor) ──────────

        /// <summary>Left edge in worksheet points (EMU-converted).</summary>
        public double LeftPt { get; set; }

        /// <summary>Top edge in worksheet points.</summary>
        public double TopPt { get; set; }

        /// <summary>Width in worksheet points.</summary>
        public double WidthPt { get; set; }

        /// <summary>Height in worksheet points.</summary>
        public double HeightPt { get; set; }

        /// <summary>Rotation in degrees (from DrawingML transform).</summary>
        public double RotationDegrees { get; set; }

        // ── Fill ──────────────────────────────────────────────────────

        /// <summary>Fill color in #AARRGGBB format (null = no fill / transparent).</summary>
        public string? FillColorArgb { get; set; }

        /// <summary>Fill opacity 0.0-1.0.</summary>
        public double FillOpacity { get; set; } = 1.0;

        // ── Border / Line ─────────────────────────────────────────────

        /// <summary>Border color in #AARRGGBB format (null = default black).</summary>
        public string? BorderColorArgb { get; set; }

        /// <summary>Border width in points (from DrawingML ln/w).</summary>
        public double BorderWidthPt { get; set; } = 0.5;

        /// <summary>Border dash style: solid, dash, dot, dashDot, etc.</summary>
        public string BorderDashStyle { get; set; } = "solid";

        // ── Text (for text boxes) ─────────────────────────────────────

        /// <summary>Text content (for textBox shapes).</summary>
        public string? Text { get; set; }

        /// <summary>Font name.</summary>
        public string FontName { get; set; } = "Calibri";

        /// <summary>Font size in points.</summary>
        public double FontSizePt { get; set; } = 11;

        /// <summary>Font color in #AARRGGBB format.</summary>
        public string? FontColorArgb { get; set; }

        /// <summary>Bold flag.</summary>
        public bool Bold { get; set; }

        /// <summary>Italic flag.</summary>
        public bool Italic { get; set; }

        /// <summary>Horizontal alignment: left, center, right.</summary>
        public string HorizontalAlignment { get; set; } = "left";

        /// <summary>Vertical alignment: top, center, bottom.</summary>
        public string VerticalAlignment { get; set; } = "top";

        /// <summary>Wrap text flag.</summary>
        public bool WrapText { get; set; }
    }
}
