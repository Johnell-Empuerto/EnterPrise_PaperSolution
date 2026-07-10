namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Configuration for the FormLess Regression &amp; Pixel-Diff Validation framework (Phase 11H).
    /// Controls tolerance, test paths, comparison modes, and report generation.
    /// </summary>
    public class RegressionConfiguration
    {
        // ── Paths ─────────────────────────────────────────────────────

        /// <summary>Root directory containing test workbook templates (.xlsx).</summary>
        public string TemplatesDirectory { get; set; } = "RegressionTests/Templates";

        /// <summary>Directory for baseline (reference) images from Excel.</summary>
        public string BaselineDirectory { get; set; } = "RegressionTests/Baseline";

        /// <summary>Directory for current (FormLess-generated) images.</summary>
        public string CurrentDirectory { get; set; } = "RegressionTests/Current";

        /// <summary>Directory for diff (comparison) images.</summary>
        public string DiffDirectory { get; set; } = "RegressionTests/Diff";

        /// <summary>Directory for validation reports.</summary>
        public string ReportsDirectory { get; set; } = "RegressionTests/Reports";

        /// <summary>Directory for logs.</summary>
        public string LogsDirectory { get; set; } = "RegressionTests/Logs";

        // ── Comparison Settings ───────────────────────────────────────

        /// <summary>Pixel tolerance (0 = exact match, 1 = ±1 per channel, etc.).</summary>
        public int PixelTolerance { get; set; } = 1;

        /// <summary>Maximum acceptable pixel difference percentage (0.0-100.0).</summary>
        public double MaxDifferencePercent { get; set; } = 0.5;

        /// <summary>Target pixel difference percentage for ideal output.</summary>
        public double TargetDifferencePercent { get; set; } = 0.1;

        /// <summary>Whether to skip alpha channel comparison.</summary>
        public bool SkipAlphaChannel { get; set; }

        /// <summary>Whether to generate diff images.</summary>
        public bool GenerateDiffImages { get; set; } = true;

        /// <summary>Whether to generate HTML reports.</summary>
        public bool GenerateHtmlReport { get; set; } = true;

        /// <summary>Whether to generate JSON reports.</summary>
        public bool GenerateJsonReport { get; set; } = true;

        // ── Rendering Settings ────────────────────────────────────────

        /// <summary>DPI for test renders (must match baseline).</summary>
        public int Dpi { get; set; } = 300;

        /// <summary>Whether to stop on first failure.</summary>
        public bool StopOnFirstFailure { get; set; }

        /// <summary>Whether to save current renders even if baselines exist.</summary>
        public bool AlwaysRegenerateCurrent { get; set; } = true;

        // ── Baseline Management ───────────────────────────────────────

        /// <summary>Whether to create baselines if they don't exist.</summary>
        public bool AutoCreateBaselines { get; set; }

        /// <summary>Whether to update baselines that have approved differences.</summary>
        public bool AutoUpdateApprovedBaselines { get; set; }

        /// <summary>File containing approved differences (JSON list of 'workbook:page' keys).</summary>
        public string ApprovedDifferencesFile { get; set; } = "RegressionTests/approved_differences.json";

        // ── Test Selection ────────────────────────────────────────────

        /// <summary>Specific workbooks to test (empty = test all).</summary>
        public List<string> WorkbookFilter { get; set; } = new();

        /// <summary>Maximum workbooks to test (0 = unlimited).</summary>
        public int MaxWorkbooks { get; set; }

        /// <summary>Maximum pages per workbook to test (0 = unlimited).</summary>
        public int MaxPagesPerWorkbook { get; set; } = 1;

        // ── Performance ───────────────────────────────────────────────

        /// <summary>Whether to record rendering performance metrics.</summary>
        public bool RecordPerformance { get; set; } = true;

        /// <summary>Whether to log memory usage.</summary>
        public bool LogMemoryUsage { get; set; } = true;

        /// <summary>Timeout in seconds per workbook (0 = no timeout).</summary>
        public int TimeoutSeconds { get; set; } = 60;
    }
}
