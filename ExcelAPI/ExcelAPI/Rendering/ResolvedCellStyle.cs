namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Canonical resolved cell style. Every render layer (FillEngine, BorderEngine,
    /// TextEngine) consumes this — never OpenXml structures directly.
    ///
    /// Resolved once by StyleResolver, cached by StyleCache, stored on RenderCell.
    /// No layer independently resolves fonts, fills, borders, or alignment.
    /// </summary>
    public class ResolvedCellStyle
    {
        // ── Font ────────────────────────────────────────────────────────
        public string FontName { get; set; } = "Calibri";
        public double FontSize { get; set; } = 11;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
        public string? FontColorArgb { get; set; }  // #AARRGGBB

        // ── Fill ────────────────────────────────────────────────────────
        public string? FillColorArgb { get; set; }   // #AARRGGBB
        public string? PatternType { get; set; }     // "none", "solid", "gray125", etc.

        // ── Border ──────────────────────────────────────────────────────
        public ResolvedBorder? Border { get; set; }

        // ── Alignment ───────────────────────────────────────────────────
        public string? HorizontalAlignment { get; set; }  // "general", "left", "center", etc.
        public string? VerticalAlignment { get; set; }    // "bottom", "top", "center", etc.
        public bool WrapText { get; set; }
        public uint Indent { get; set; }
        public double TextRotation { get; set; }
    }

    /// <summary>
    /// Resolved border with pre-computed style name and color.
    /// </summary>
    public class ResolvedBorder
    {
        public ResolvedBorderItem? Top { get; set; }
        public ResolvedBorderItem? Bottom { get; set; }
        public ResolvedBorderItem? Left { get; set; }
        public ResolvedBorderItem? Right { get; set; }
        public ResolvedBorderItem? DiagonalUp { get; set; }
        public ResolvedBorderItem? DiagonalDown { get; set; }
    }

    /// <summary>
    /// Single border edge with resolved style name and color.
    /// </summary>
    public class ResolvedBorderItem
    {
        /// <summary>Border style: "thin", "medium", "thick", "double", "dotted", etc.</summary>
        public string Style { get; set; } = "thin";

        /// <summary>Border color as #AARRGGBB, or null for default (black).</summary>
        public string? ColorArgb { get; set; }
    }
}
