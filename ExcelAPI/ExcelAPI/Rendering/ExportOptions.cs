namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Configurable options for the Production Export Engine (Phase 11G).
    /// Controls DPI, background, font embedding, multi-page, and metadata settings.
    /// Both PNG and PDF exports consume this options object.
    /// </summary>
    public class ExportOptions
    {
        /// <summary>Rendering DPI: 96, 150, 300(default), or 600.</summary>
        public int Dpi { get; set; } = 300;

        /// <summary>Whether to use a transparent background instead of white.</summary>
        public bool TransparentBackground { get; set; }

        /// <summary>Whether to embed fonts in PDF output (increases file size).</summary>
        public bool EmbedFonts { get; set; }

        /// <summary>Whether to rasterize images only in PDF (true = vector everything else).</summary>
        public bool RasterImagesOnly { get; set; } = true;

        /// <summary>Whether to compress PDF output.</summary>
        public bool CompressPdf { get; set; } = true;

        /// <summary>Whether to include PDF metadata (title, author, date, etc.).</summary>
        public bool IncludeMetadata { get; set; } = true;

        /// <summary>PDF title (overrides workbook title).</summary>
        public string? Title { get; set; }

        /// <summary>PDF author.</summary>
        public string? Author { get; set; }

        /// <summary>PDF subject.</summary>
        public string? Subject { get; set; }

        /// <summary>PDF keywords.</summary>
        public string? Keywords { get; set; }

        /// <summary>Enable high-quality antialiasing.</summary>
        public bool HighQualityAntialiasing { get; set; } = true;

        /// <summary>Enable LCD-optimized text rendering.</summary>
        public bool LcdTextRendering { get; set; }

        /// <summary>Enable subpixel text positioning.</summary>
        public bool SubpixelText { get; set; } = true;

        // ── Multi-page ────────────────────────────────────────────────

        /// <summary>Maximum pages to render (0 = unlimited).</summary>
        public int MaxPages { get; set; }

        /// <summary>Include page numbers on each page.</summary>
        public bool IncludePageNumbers { get; set; }

        /// <summary>Page number format (e.g., "Page {0} of {1}").</summary>
        public string PageNumberFormat { get; set; } = "Page {0} of {1}";

        /// <summary>Page number font size in points.</summary>
        public double PageNumberFontSizePt { get; set; } = 8;

        // ── PNG specifics ─────────────────────────────────────────────

        /// <summary>PNG compression level (0-100, higher = better quality).</summary>
        public int PngCompressionLevel { get; set; } = 100;

        /// <summary>Whether to crop PNG output to the exact PrintArea bounds.</summary>
        public bool CropToPrintArea { get; set; }

        // ── Preview helpers ───────────────────────────────────────────

        /// <summary>Output file prefix (e.g., \"page\" → \"page1.png\", \"page2.png\").</summary>
        public string FilePrefix { get; set; } = "page";

        /// <summary>Output directory for multi-page exports.</summary>
        public string? OutputDirectory { get; set; }

        /// <summary>Computed points-to-pixels multiplier.</summary>
        public double PointsToPixels => Dpi / 72.0;
    }
}
