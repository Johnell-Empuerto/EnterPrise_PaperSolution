namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Complete runtime form definition produced by the FormRuntimeBuilder.
    /// Consumed by the Next.js frontend via GET /api/runtime/{templateId}.
    /// No rendering logic — pure runtime metadata.
    /// </summary>
    public class RuntimeForm
    {
        /// <summary>Workbook file name (without extension).</summary>
        public string WorkbookName { get; set; } = "";

        /// <summary>Human-readable title from workbook metadata.</summary>
        public string Title { get; set; } = "";

        /// <summary>List of sheets in this runtime form.</summary>
        public List<RuntimeSheet> Sheets { get; set; } = new();

        // ── Page Dimensions ───────────────────────────────────────────

        /// <summary>Page width in pixels (at rendering DPI).</summary>
        public int PageWidth { get; set; }

        /// <summary>Page height in pixels (at rendering DPI).</summary>
        public int PageHeight { get; set; }

        /// <summary>Rendering scale factor (1.0 = 100%).</summary>
        public double Scale { get; set; } = 1.0;

        /// <summary>Rendering DPI (typically 300).</summary>
        public int Dpi { get; set; } = 300;

        // ── Metadata ──────────────────────────────────────────────────

        /// <summary>Runtime engine version.</summary>
        public string Version { get; set; } = "1.0";
    }
}
