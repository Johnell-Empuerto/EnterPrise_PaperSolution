using System.Text.Json.Serialization;

// ────────────────────────────────────────────────────────────────────────────
// StyleDefinition — Cell Style, Font, Border, Fill, Alignment
//
// Canonical style primitives consumed by FieldDefinition and CellStyleInfo.
// These match the resolved style model in the Rendering layer (ResolvedCellStyle)
// and the COM-read styles (CellStyleInfo).
//
// Ownership: Shared
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// Complete cell style definition: font, border, fill, alignment.
    /// Mirrors the Rendering layer's ResolvedCellStyle as the canonical form.
    /// </summary>
    public class CellStyle
    {
        /// <summary>Font properties.</summary>
        public FontDefinition Font { get; set; } = new();

        /// <summary>Border definition (top, bottom, left, right, diagonal).</summary>
        public BorderDefinition? Border { get; set; }

        /// <summary>Fill / background pattern.</summary>
        public FillDefinition? Fill { get; set; }

        /// <summary>Text alignment.</summary>
        public AlignmentDefinition Alignment { get; set; } = new();

        /// <summary>Whether text wraps within the cell.</summary>
        public bool WrapText { get; set; }

        /// <summary>Indentation level (number of characters).</summary>
        public uint Indent { get; set; }

        /// <summary>Text rotation in degrees (0 = horizontal, 90 = vertical, -90 = rotated).</summary>
        public double TextRotation { get; set; }
    }

    /// <summary>
    /// Font properties — color in #AARRGGBB format.
    /// </summary>
    public class FontDefinition
    {
        /// <summary>Font family name (e.g., "Calibri", "Arial").</summary>
        public string Name { get; set; } = "Calibri";

        /// <summary>Font size in points (e.g., 11).</summary>
        public double SizePt { get; set; } = 11;

        /// <summary>Whether the font is bold.</summary>
        public bool Bold { get; set; }

        /// <summary>Whether the font is italic.</summary>
        public bool Italic { get; set; }

        /// <summary>Whether the font is underlined.</summary>
        public bool Underline { get; set; }

        /// <summary>Whether the font is strikethrough.</summary>
        public bool Strikeout { get; set; }

        /// <summary>Font color as #AARRGGBB (null = default / auto).</summary>
        public string? ColorArgb { get; set; }
    }

    /// <summary>
    /// Complete border definition with all six edges.
    /// </summary>
    public class BorderDefinition
    {
        /// <summary>Top edge border.</summary>
        public BorderEdge? Top { get; set; }

        /// <summary>Bottom edge border.</summary>
        public BorderEdge? Bottom { get; set; }

        /// <summary>Left edge border.</summary>
        public BorderEdge? Left { get; set; }

        /// <summary>Right edge border.</summary>
        public BorderEdge? Right { get; set; }

        /// <summary>Diagonal-up border (bottom-left to top-right).</summary>
        public BorderEdge? DiagonalUp { get; set; }

        /// <summary>Diagonal-down border (top-left to bottom-right).</summary>
        public BorderEdge? DiagonalDown { get; set; }

        /// <summary>
        /// Whether any edge has a border defined.
        /// </summary>
        public bool HasBorder =>
            Top != null || Bottom != null || Left != null || Right != null ||
            DiagonalUp != null || DiagonalDown != null;
    }

    /// <summary>
    /// A single border edge with style, color, and computed width.
    /// </summary>
    public class BorderEdge
    {
        /// <summary>
        /// Border style name. One of:
        /// "none", "thin", "medium", "thick", "double", "dotted", "dashed",
        /// "dashDot", "dashDotDot", "hair", "mediumDashed", "mediumDashDot",
        /// "mediumDashDotDot", "slantDashDot".
        /// </summary>
        public string Style { get; set; } = "thin";

        /// <summary>Border color as #AARRGGBB (null = default/auto, typically black).</summary>
        public string? ColorArgb { get; set; }

        /// <summary>Pre-computed border width in points.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public double WidthPt
        {
            get
            {
                return Style switch
                {
                    "hair" => 0.25,
                    "dotted" or "dashed" or "dashDot" or "dashDotDot" => 0.5,
                    "thin" => 0.5,
                    "medium" or "mediumDashed" or "mediumDashDot" or "mediumDashDotDot" => 1.0,
                    "thick" or "double" => 2.0,
                    _ => 0.5, // default for unrecognized styles
                };
            }
        }
    }

    /// <summary>
    /// Fill / background pattern.
    /// </summary>
    public class FillDefinition
    {
        /// <summary>
        /// Pattern type: "none", "solid", "gray125", "gray0625", "gray50",
        /// "gray75", "gray25", "horizontalStripe", "verticalStripe",
        /// "reverseDiagonalStripe", "diagonalStripe", "thinHorizontalStripe",
        /// "thinVerticalStripe", "thinReverseDiagonalStripe",
        /// "thinDiagonalStripe", "thinHorizontalCrossHatch",
        /// "thinDiagonalCrossHatch", etc.
        /// </summary>
        public string PatternType { get; set; } = "none";

        /// <summary>Foreground (pattern) color as #AARRGGBB.</summary>
        public string? ColorArgb { get; set; }

        /// <summary>Background color as #AARRGGBB (used only for patterned fills).</summary>
        public string? PatternColorArgb { get; set; }

        /// <summary>
        /// Whether this fill has any visible color.
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.Always)]
        public bool HasFill =>
            !string.IsNullOrEmpty(ColorArgb) && PatternType != "none";
    }

    /// <summary>
    /// Horizontal and vertical text alignment.
    /// </summary>
    public class AlignmentDefinition
    {
        /// <summary>
        /// Horizontal alignment. One of:
        /// "general", "left", "center", "right", "fill", "justify",
        /// "centerAcrossSelection", "distributed".
        /// </summary>
        public string Horizontal { get; set; } = "general";

        /// <summary>
        /// Vertical alignment. One of:
        /// "top", "center", "bottom", "justify", "distributed".
        /// </summary>
        public string Vertical { get; set; } = "bottom";
    }
}
