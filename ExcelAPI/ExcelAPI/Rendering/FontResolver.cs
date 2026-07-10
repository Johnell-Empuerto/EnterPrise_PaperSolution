using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Resolves SkiaSharp typefaces from font names with a fallback chain.
    /// Never crashes if a font is unavailable — always returns a usable typeface.
    ///
    /// Fallback chain (ordered by preference):
    ///   1. Requested font (e.g., "Aptos Narrow", "Calibri")
    ///   2. Font family fallback (e.g., "Arial" for sans-serif)
    ///   3. "Segoe UI"
    ///   4. "Arial"
    ///   5. SKTypeface.Default (system default)
    /// </summary>
    public class FontResolver
    {
        /// <summary>
        /// Resolve a typeface for the given font name, bold, and italic settings.
        /// </summary>
        public SKTypeface ResolveTypeface(string? fontName, bool bold, bool italic)
        {
            string name = (fontName ?? "Calibri").Trim();

            // Try the requested font first
            var tf = TryCreateTypeface(name, bold, italic);
            if (tf != null) return tf;

            // Try known fallbacks in order
            foreach (var fallback in GetFallbackNames(name))
            {
                tf = TryCreateTypeface(fallback, bold, italic);
                if (tf != null) return tf;
            }

            // Last resort: system default
            return SKTypeface.Default;
        }

        /// <summary>
        /// Try to create a typeface for the given name. Returns null if not available.
        /// SKTypeface.FromFamilyName returns a non-null typeface even for unknown fonts
        /// (SkiaSharp falls back internally), so exceptions are not expected in practice.
        /// </summary>
        private static SKTypeface? TryCreateTypeface(string familyName, bool bold, bool italic)
        {
            var weight = bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
            var slant = italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
            var width = SKFontStyleWidth.Normal;

            return SKTypeface.FromFamilyName(familyName, weight, width, slant);
        }

        /// <summary>
        /// Get the fallback font name chain for a given font.
        /// Maps common Excel fonts to available system fonts.
        /// </summary>
        private static string[] GetFallbackNames(string fontName)
        {
            string lower = fontName.ToLowerInvariant();

            // Known font mapping
            if (lower.Contains("aptos"))      // Aptos Narrow — new Office default
                return new[] { "Calibri", "Segoe UI", "Arial" };
            if (lower.Contains("calibri"))
                return new[] { "Segoe UI", "Arial", "Calibri" };
            if (lower.Contains("meiryo"))
                return new[] { "Yu Gothic", "Segoe UI", "Arial" };

            // Japanese fonts
            if (lower.Contains("yu gothic") || lower.Contains("yugothic"))
                return new[] { "Meiryo", "Segoe UI", "Arial" };
            if (lower.Contains("ms gothic") || lower.Contains("msgothic"))
                return new[] { "Yu Gothic", "Meiryo", "Segoe UI" };
            if (lower.Contains("ms mincho") || lower.Contains("msmincho"))
                return new[] { "Yu Mincho", "Meiryo", "Segoe UI" };

            // Default fallback for sans-serif fonts
            if (lower.Contains("arial") || lower.Contains("helvetica") ||
                lower.Contains("verdana") || lower.Contains("tahoma") ||
                lower.Contains("trebuchet"))
                return new[] { "Arial", "Segoe UI", "Calibri" };

            // Serif fonts
            if (lower.Contains("times") || lower.Contains("georgia") ||
                lower.Contains("garamond") || lower.Contains("palatino"))
                return new[] { "Times New Roman", "Georgia", "Segoe UI" };

            // Monospace fonts
            if (lower.Contains("courier") || lower.Contains("consolas") ||
                lower.Contains("monaco") || lower.Contains("monospace"))
                return new[] { "Courier New", "Consolas", "Segoe UI" };

            // Generic fallback
            return new[] { "Segoe UI", "Arial", "Calibri" };
        }

        /// <summary>
        /// Check if a font name represents a CJK (Chinese/Japanese/Korean) font.
        /// </summary>
        public static bool IsCjkFont(string? fontName)
        {
            if (string.IsNullOrEmpty(fontName)) return false;
            string lower = fontName.ToLowerInvariant();
            return lower.Contains("yu gothic") || lower.Contains("yugothic") ||
                   lower.Contains("meiryo") || lower.Contains("ms gothic") ||
                   lower.Contains("msgothic") || lower.Contains("ms mincho") ||
                   lower.Contains("msmincho") || lower.Contains("yu mincho") ||
                   lower.Contains("yumincho") || lower.Contains("ipa") ||
                   lower.Contains("microsoft jhenghei") || lower.Contains("microsoft yahei");
        }
    }
}
