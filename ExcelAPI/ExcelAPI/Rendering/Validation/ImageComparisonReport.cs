using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Comprehensive validation report for a single workbook comparison.
    /// Serialized to both JSON (for programmatic consumption) and HTML (for visual review).
    /// </summary>
    public class ImageComparisonReport
    {
        // ── Workbook Info ─────────────────────────────────────────────

        /// <summary>Workbook file name (e.g., "Invoice.xlsx").</summary>
        public string WorkbookName { get; set; } = "";

        /// <summary>Full path to the workbook template.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? WorkbookPath { get; set; }

        /// <summary>Number of pages compared.</summary>
        public int Pages { get; set; }

        /// <summary>Overall pixel difference percentage.</summary>
        public double PixelDifference { get; set; }

        /// <summary>Whether the overall validation passed.</summary>
        public bool Passed { get; set; }

        // ── Per-Category Results ──────────────────────────────────────

        /// <summary>Whether cell fills matched within tolerance.</summary>
        public bool FillsMatched { get; set; } = true;

        /// <summary>Whether borders matched.</summary>
        public bool BordersMatched { get; set; } = true;

        /// <summary>Whether text rendered correctly.</summary>
        public bool TextMatched { get; set; } = true;

        /// <summary>Whether images matched.</summary>
        public bool ImagesMatched { get; set; } = true;

        /// <summary>Whether shapes matched.</summary>
        public bool ShapesMatched { get; set; } = true;

        /// <summary>Whether gridlines matched.</summary>
        public bool GridlinesMatched { get; set; } = true;

        /// <summary>Whether scaling/alignment matched.</summary>
        public bool LayoutMatched { get; set; } = true;

        // ── Performance ───────────────────────────────────────────────

        /// <summary>Total rendering time in milliseconds.</summary>
        public long RenderingTimeMs { get; set; }

        /// <summary>Workbook parse time in milliseconds.</summary>
        public long ParseTimeMs { get; set; }

        /// <summary>Memory used during rendering (bytes).</summary>
        public long MemoryUsageBytes { get; set; }

        /// <summary>Peak bitmap pixel dimensions.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? PeakBitmapSize { get; set; }

        // ── Per-Page Results ──────────────────────────────────────────

        /// <summary>Detailed results for each page.</summary>
        public List<PageComparisonResult> PageResults { get; set; } = new();

        /// <summary>Timestamp of the validation run.</summary>
        public string Timestamp { get; set; } = DateTime.UtcNow.ToString("o");

        /// <summary>FormLess rendering engine version.</summary>
        public string EngineVersion { get; set; } = "1.3";
    }

    /// <summary>
    /// Comparison result for a single page.
    /// </summary>
    public class PageComparisonResult
    {
        /// <summary>Page number (1-based).</summary>
        public int PageNumber { get; set; }

        /// <summary>Pixel difference percentage for this page.</summary>
        public double PixelDifference { get; set; }

        /// <summary>Number of different pixels.</summary>
        public int DifferentPixels { get; set; }

        /// <summary>Total pixel count.</summary>
        public int TotalPixels { get; set; }

        /// <summary>Whether this page passed validation.</summary>
        public bool Passed { get; set; }

        /// <summary>Maximum per-channel error (0-255).</summary>
        public int MaxError { get; set; }

        /// <summary>Average error across different pixels.</summary>
        public double AverageError { get; set; }

        /// <summary>Path to the expected (baseline) image.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ExpectedImagePath { get; set; }

        /// <summary>Path to the actual (FormLess) image.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ActualImagePath { get; set; }

        /// <summary>Path to the diff image.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? DiffImagePath { get; set; }

        /// <summary>Page render time in milliseconds.</summary>
        public long RenderTimeMs { get; set; }

        /// <summary>Page dimensions (width x height).</summary>
        public string? Dimensions { get; set; }
    }

    /// <summary>
    /// Summary statistics across all tested workbooks.
    /// </summary>
    public class ValidationSummary
    {
        /// <summary>Total workbooks tested.</summary>
        public int TotalWorkbooks { get; set; }

        /// <summary>Workbooks that passed.</summary>
        public int Passed { get; set; }

        /// <summary>Workbooks that failed.</summary>
        public int Failed { get; set; }

        /// <summary>Total pages compared.</summary>
        public int TotalPages { get; set; }

        /// <summary>Average pixel difference across all workbooks.</summary>
        public double AveragePixelDifference { get; set; }

        /// <summary>Total rendering time.</summary>
        public long TotalRenderingTimeMs { get; set; }

        /// <summary>Average rendering time per workbook.</summary>
        public double AverageRenderingTimeMs { get; set; }

        /// <summary>Timestamp of the validation run.</summary>
        public string Timestamp { get; set; } = DateTime.UtcNow.ToString("o");

        /// <summary>Individual workbook reports.</summary>
        public List<ImageComparisonReport> Reports { get; set; } = new();

        /// <summary>Serialize to JSON.</summary>
        public string ToJson()
        {
            return JsonSerializer.Serialize(this, new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
    }
}
