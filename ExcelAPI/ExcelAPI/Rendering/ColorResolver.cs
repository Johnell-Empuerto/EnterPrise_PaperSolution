namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Resolves Excel color values (RGB, Indexed, Theme, Auto) to #AARRGGBB strings.
    /// Theme colors require a ThemeResolver to resolve.
    /// Single authoritative source for color resolution — no other class duplicates this.
    /// </summary>
    public class ColorResolver
    {
        private readonly ThemeResolver _theme;

        public ColorResolver(ThemeResolver theme)
        {
            _theme = theme;
        }

        /// <summary>
        /// Load theme from workbook part to enable theme color resolution.
        /// </summary>
        public void LoadTheme(DocumentFormat.OpenXml.Packaging.WorkbookPart wbPart)
        {
            _theme.LoadFromWorkbook(wbPart);
        }

        /// <summary>
        /// Resolve an Excel color to a #AARRGGBB string.
        /// Priority: RGB → Theme → Indexed → Auto → null.
        /// </summary>
        public string? Resolve(DocumentFormat.OpenXml.Spreadsheet.ColorType? color)
        {
            if (color == null) return null;

            // Priority 1: Explicit RGB value
            if (color.Rgb?.Value != null)
            {
                string rgb = color.Rgb.Value;
                return rgb.Length == 6 ? "#" + rgb : "#" + rgb;
            }

            // Priority 2: Theme color (requires ThemeResolver)
            if (color.Theme?.Value != null)
            {
                double? tint = color.Tint?.Value;
                return _theme.ResolveThemeColor((int)color.Theme.Value, tint);
            }

            // Priority 3: Indexed palette
            if (color.Indexed?.Value != null)
            {
                return IndexedToArgb((int)color.Indexed.Value);
            }

            // Priority 4: Auto color (default black)
            if (color.Auto?.Value == true)
            {
                return "#FF000000";
            }

            return null;
        }

        /// <summary>
        /// Convert an RGB string (6 or 8 hex chars) to #AARRGGBB.
        /// </summary>
        public static string RgbToArgb(string? rgb)
        {
            if (string.IsNullOrEmpty(rgb)) return "#FF000000";
            string clean = rgb.TrimStart('#');
            return clean.Length switch
            {
                6 => "#FF" + clean,
                8 => "#" + clean,
                _ => "#FF000000"
            };
        }

        /// <summary>
        /// Standard Excel 56-color indexed palette.
        /// </summary>
        public static string? IndexedToArgb(int index)
        {
            return index switch
            {
                0 => "#FF000000", 1 => "#FFFFFFFF", 2 => "#FFFF0000", 3 => "#FF00FF00",
                4 => "#FF0000FF", 5 => "#FFFFFF00", 6 => "#FFFF00FF", 7 => "#FF00FFFF",
                8 => "#FF800000", 9 => "#FF008000", 10 => "#FF000080", 11 => "#FF808000",
                12 => "#FF800080", 13 => "#FF008080", 14 => "#FFC0C0C0", 15 => "#FF808080",
                16 => "#FF9999FF", 17 => "#FF993366", 18 => "#FFFFFFCC", 19 => "#FFCCFFFF",
                20 => "#FF660066", 21 => "#FFFF8080", 22 => "#FF0066CC", 23 => "#FFCCCCFF",
                24 => "#FF000080", 25 => "#FFFF00FF", 26 => "#FFFFFF00", 27 => "#FF00FFFF",
                28 => "#FF800080", 29 => "#FF800000", 30 => "#FF008080", 31 => "#FF0000FF",
                32 => "#FF00CCFF", 33 => "#FFCCFFFF", 34 => "#FFCCFFCC", 35 => "#FFFFFF99",
                36 => "#FF99CCFF", 37 => "#FFFF99CC", 38 => "#FFCC99FF", 39 => "#FFFFCC99",
                40 => "#FF3366FF", 41 => "#FF33CCCC", 42 => "#FF99CC00", 43 => "#FFFFCC00",
                44 => "#FFFF9900", 45 => "#FFFF6600", 46 => "#FF666699", 47 => "#FF969696",
                48 => "#FF003366", 49 => "#FF339966", 50 => "#FF003300", 51 => "#FF333300",
                52 => "#FF993300", 53 => "#FF993366", 54 => "#FF333399", 55 => "#FF333333",
                _ => null
            };
        }
    }
}
