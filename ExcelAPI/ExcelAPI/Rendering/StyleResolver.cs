using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Orchestrates all style resolution from the OpenXml Stylesheet.
    /// Produces ResolvedCellStyle objects via StyleCache.
    ///
    /// Replaces OpenXmlParser.ApplyCellStyle() — style logic is now centralized here.
    /// Every render layer consumes ResolvedCellStyle instead of parsing OpenXml structures.
    /// </summary>
    public class StyleResolver
    {
        private readonly ColorResolver _colorResolver;
        private readonly StyleCache _cache;

        public StyleResolver(ColorResolver colorResolver, StyleCache cache)
        {
            _colorResolver = colorResolver;
            _cache = cache;
        }

        /// <summary>
        /// Resolve a cell's style based on its style index, using the workbook's stylesheet.
        /// Results are cached so each CellFormat is resolved only once.
        /// </summary>
        public ResolvedCellStyle Resolve(Stylesheet styleSheet, uint? styleIndex,
            string defaultFontName, double defaultFontSize)
        {
            int idx = (int)(styleIndex ?? 0);

            return _cache.GetOrAdd(idx, (_) =>
            {
                var style = new ResolvedCellStyle
                {
                    FontName = defaultFontName,
                    FontSize = defaultFontSize
                };

                var cellFormats = styleSheet.CellFormats;
                if (cellFormats == null || idx >= cellFormats.Count)
                    return style;

                var cf = cellFormats.ChildElements[idx] as CellFormat;
                if (cf == null) return style;

                // Resolve font
                if (cf.FontId?.Value != null && styleSheet.Fonts != null
                    && cf.FontId.Value < styleSheet.Fonts.Count)
                {
                    var font = styleSheet.Fonts.ChildElements[(int)cf.FontId.Value] as Font;
                    if (font != null)
                    {
                        style.FontName = font.FontName?.Val?.Value ?? style.FontName;
                        style.FontSize = (double)(font.FontSize?.Val?.Value ?? style.FontSize);
                        style.Bold = font.Bold?.Val ?? false;
                        style.Italic = font.Italic?.Val ?? false;
                        style.Underline = font.Underline != null;
                        style.Strikeout = font.Strike?.Val ?? false;

                        if (font.Color != null)
                            style.FontColorArgb = _colorResolver.Resolve(font.Color);
                    }
                }

                // Resolve fill
                if (cf.FillId?.Value != null && styleSheet.Fills != null
                    && cf.FillId.Value < styleSheet.Fills.Count)
                {
                    var fill = styleSheet.Fills.ChildElements[(int)cf.FillId.Value] as Fill;
                    if (fill?.PatternFill != null)
                    {
                        var pf = fill.PatternFill;
                        if (pf.PatternType?.Value != null)
                            style.PatternType = pf.PatternType.Value.ToString();

                        if (pf.ForegroundColor != null)
                            style.FillColorArgb = _colorResolver.Resolve(pf.ForegroundColor);
                        else if (pf.BackgroundColor != null)
                            style.FillColorArgb = _colorResolver.Resolve(pf.BackgroundColor);
                    }
                }

                // Resolve border
                if (cf.BorderId?.Value != null && styleSheet.Borders != null
                    && cf.BorderId.Value < styleSheet.Borders.Count)
                {
                    var border = styleSheet.Borders.ChildElements[(int)cf.BorderId.Value]
                        as DocumentFormat.OpenXml.Spreadsheet.Border;
                    if (border != null)
                    {
                        var rb = new ResolvedBorder
                        {
                            Left = ResolveBorderItem(border.LeftBorder),
                            Right = ResolveBorderItem(border.RightBorder),
                            Top = ResolveBorderItem(border.TopBorder),
                            Bottom = ResolveBorderItem(border.BottomBorder),
                        };

                        // Diagonal borders
                        var diag = border.GetFirstChild<DiagonalBorder>();
                        if (diag != null)
                        {
                            rb.DiagonalUp = ResolveBorderItem(diag);
                        }

                        style.Border = rb;
                    }
                }

                // Resolve alignment
                if (cf.Alignment != null)
                {
                    var align = cf.Alignment;
                    if (align.Horizontal?.Value != null)
                    {
                        string hVal = align.Horizontal.Value.ToString();
                        style.HorizontalAlignment = char.ToLowerInvariant(hVal[0]) + hVal[1..];
                    }
                    if (align.Vertical?.Value != null)
                    {
                        string vVal = align.Vertical.Value.ToString();
                        style.VerticalAlignment = char.ToLowerInvariant(vVal[0]) + vVal[1..];
                    }
                    style.WrapText = align.WrapText?.Value ?? false;
                    style.Indent = align.Indent?.Value ?? 0;
                    style.TextRotation = align.TextRotation?.Value ?? 0;
                }

                return style;
            });
        }

        /// <summary>
        /// Resolve a single border edge.
        /// </summary>
        private ResolvedBorderItem? ResolveBorderItem(BorderPropertiesType? bp)
        {
            if (bp?.Style?.Value == null) return null;

            string styleName = bp.Style.Value.ToString();
            styleName = char.ToLowerInvariant(styleName[0]) + styleName[1..];

            return new ResolvedBorderItem
            {
                Style = styleName,
                ColorArgb = bp.Color != null ? _colorResolver.Resolve(bp.Color) : null
            };
        }

        /// <summary>
        /// Pre-resolve all styles in the stylesheet (eager caching).
        /// Called once per workbook after parsing.
        /// </summary>
        public void PreResolveAll(Stylesheet styleSheet,
            string defaultFontName, double defaultFontSize)
        {
            var cellFormats = styleSheet.CellFormats;
            if (cellFormats == null) return;

            for (int i = 0; i < cellFormats.Count; i++)
            {
                Resolve(styleSheet, (uint)i, defaultFontName, defaultFontSize);
            }
        }

        /// <summary>
        /// Load theme colors from the workbook part.
        /// Must be called from the parser after opening the workbook.
        /// </summary>
        public void LoadTheme(WorkbookPart wbPart)
        {
            _colorResolver.LoadTheme(wbPart);
        }
    }
}
