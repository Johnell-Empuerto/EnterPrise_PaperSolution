using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Reads theme1.xml from the workbook and resolves theme colors.
    /// Supports all 12 theme colors (Accent1-6, Dark1/2, Light1/2, Hyperlink, FollowedHyperlink)
    /// with tint/shade calculations using ECMA-376 color transformation.
    ///
    /// This was a remaining limitation from Phase 11C — theme colors now resolve to #AARRGGBB.
    /// </summary>
    public class ThemeResolver
    {
        private static readonly string[] ThemeColorNames =
        {
            "Dark1", "Light1", "Dark2", "Light2",
            "Accent1", "Accent2", "Accent3", "Accent4",
            "Accent5", "Accent6",
            "Hyperlink", "FollowedHyperlink"
        };

        private readonly Dictionary<int, string> _themeColors = new();
        private bool _loaded;

        /// <summary>
        /// Load theme colors from the workbook part.
        /// Must be called before ResolveThemeColor.
        /// </summary>
        public void LoadFromWorkbook(WorkbookPart wbPart)
        {
            if (_loaded) return;

            var themePart = wbPart.ThemePart;
            if (themePart?.Theme == null)
            {
                LoadDefaults();
                _loaded = true;
                return;
            }

            var theme = themePart.Theme;
            var themeElements = theme.ThemeElements;
            if (themeElements?.ColorScheme == null)
            {
                LoadDefaults();
                _loaded = true;
                return;
            }

            var colorScheme = themeElements.ColorScheme;
            var colorValues = new Dictionary<string, string>();

            // Read color scheme children using LINQ to XML
            foreach (var child in colorScheme.ChildElements)
            {
                string elementName = child.LocalName ?? "";
                if (string.IsNullOrEmpty(elementName)) continue;

                // Try to read the color value from child elements
                string? hexColor = null;

                // Read color value from child element:
                //   <srgbClr val="44546A"/>  — RGB hex value
                //   <sysClr val="windowText"/> — system color name
                foreach (var subChild in child.ChildElements)
                {
                    string subName = subChild.LocalName ?? "";
                    if (subName is "srgbClr" or "srgbclr")
                    {
                        // Use GetAttributes() to avoid KeyNotFoundException from GetAttribute()
                        hexColor = subChild.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "val")
                            .Value;
                        break;
                    }
                    if (subName is "sysClr" or "sysclr")
                    {
                        hexColor = subChild.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "val")
                            .Value;
                        // sysClr returns a system color name like "windowText"
                        if (!string.IsNullOrEmpty(hexColor) && !hexColor.StartsWith("#"))
                        {
                            hexColor = MapSystemColor(hexColor);
                        }
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(hexColor))
                {
                    colorValues[elementName] = ColorResolver.RgbToArgb("#" + hexColor.TrimStart('#'));
                }
            }

            // Map to theme color indices
            for (int i = 0; i < ThemeColorNames.Length; i++)
            {
                if (colorValues.TryGetValue(ThemeColorNames[i], out var hex))
                {
                    _themeColors[i] = hex;
                }
            }

            _loaded = true;
        }

        /// <summary>
        /// Resolve a theme color by index with optional tint.
        /// </summary>
        public string? ResolveThemeColor(int themeIndex, double? tint)
        {
            if (!_loaded) return null;
            if (!_themeColors.TryGetValue(themeIndex, out var baseColor))
                return null;
            if (tint == null || tint == 0)
                return baseColor;
            return ApplyTint(baseColor, tint.Value);
        }

        /// <summary>
        /// Apply tint/shade using ECMA-376 algorithm.
        /// </summary>
        private static string ApplyTint(string argb, double tint)
        {
            string hex = argb.TrimStart('#');
            if (hex.Length != 8) return argb;

            byte a = Convert.ToByte(hex[..2], 16);
            byte r = Convert.ToByte(hex[2..4], 16);
            byte g = Convert.ToByte(hex[4..6], 16);
            byte b = Convert.ToByte(hex[6..8], 16);

            double t = Math.Clamp(tint, -1.0, 1.0);

            if (t > 0)
            {
                r = (byte)(r + (255 - r) * t);
                g = (byte)(g + (255 - g) * t);
                b = (byte)(b + (255 - b) * t);
            }
            else if (t < 0)
            {
                double factor = 1.0 + t;
                r = (byte)(r * factor);
                g = (byte)(g * factor);
                b = (byte)(b * factor);
            }

            return $"#{a:X2}{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>
        /// Map a system color name (from sysClr) to a usable hex ARGB value.
        /// </summary>
        private static string? MapSystemColor(string sysColor)
        {
            // Common Windows system colors mapped to approximate hex values
            return sysColor.ToLowerInvariant() switch
            {
                "windowtext" => "#FF000000",
                "window" => "#FFFFFFFF",
                "buttontext" => "#FF000000",
                "buttonface" => "#FFF0F0F0",
                "highlight" => "#FF0078D7",
                "highlighttext" => "#FFFFFFFF",
                "captiontext" => "#FF000000",
                "activecaption" => "#FF0055A4",
                "graytext" => "#FF6D6D6D",
                "menutext" => "#FF000000",
                "inactivecaptiontext" => "#FF000000",
                "inactivecaption" => "#FFE0E0E0",
                "infotext" => "#FF000000",
                "infobk" => "#FFFFFFE0",
                _ => "#FF000000"  // default to black
            };
        }

        private void LoadDefaults()
        {
            _themeColors[0] = "#FF000000";   // Dark1
            _themeColors[1] = "#FFFFFFFF";   // Light1
            _themeColors[2] = "#FF44546A";   // Dark2
            _themeColors[3] = "#FFF2F2F2";   // Light2
            _themeColors[4] = "#FF4472C4";   // Accent1
            _themeColors[5] = "#FFED7D31";   // Accent2
            _themeColors[6] = "#FFA5A5A5";   // Accent3
            _themeColors[7] = "#FFFFC000";   // Accent4
            _themeColors[8] = "#FF5B9BD5";   // Accent5
            _themeColors[9] = "#FF70AD47";   // Accent6
            _themeColors[10] = "#FF0563C1";  // Hyperlink
            _themeColors[11] = "#FF954F72";  // FollowedHyperlink
        }
    }
}
