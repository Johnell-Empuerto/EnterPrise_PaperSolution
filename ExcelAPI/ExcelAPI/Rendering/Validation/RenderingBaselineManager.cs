using System.Text.Json;

namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Manages baseline (reference) images for regression testing.
    /// Supports creating, updating, and approving differences.
    ///
    /// Folder structure:
    ///   RegressionTests/
    ///     Templates/          — XLSX workbooks to test
    ///     Baseline/           — Reference images (from Excel)
    ///     Current/            — FormLess-generated images
    ///     Diff/               — Difference images
    ///     Reports/            — Validation reports
    ///     approved_differences.json  — Pre-approved differences
    /// </summary>
    public class RenderingBaselineManager
    {
        private readonly RegressionConfiguration _config;
        private readonly HashSet<string> _approvedDifferences;

        public RenderingBaselineManager(RegressionConfiguration config)
        {
            _config = config;
            _approvedDifferences = LoadApprovedDifferences();
        }

        /// <summary>
        /// Get the expected path for a baseline image.
        /// </summary>
        public string GetBaselinePath(string workbookName, int pageNumber)
        {
            string name = Path.GetFileNameWithoutExtension(workbookName);
            return Path.Combine(_config.BaselineDirectory, $"{name}_page{pageNumber}.png");
        }

        /// <summary>
        /// Get the path for a current (FormLess) render.
        /// </summary>
        public string GetCurrentPath(string workbookName, int pageNumber)
        {
            string name = Path.GetFileNameWithoutExtension(workbookName);
            return Path.Combine(_config.CurrentDirectory, $"{name}_page{pageNumber}.png");
        }

        /// <summary>
        /// Get the path for a diff image.
        /// </summary>
        public string GetDiffPath(string workbookName, int pageNumber)
        {
            string name = Path.GetFileNameWithoutExtension(workbookName);
            return Path.Combine(_config.DiffDirectory, $"{name}_page{pageNumber}_diff.png");
        }

        /// <summary>
        /// Check whether a baseline image exists for the given workbook and page.
        /// </summary>
        public bool BaselineExists(string workbookName, int pageNumber)
        {
            return System.IO.File.Exists(GetBaselinePath(workbookName, pageNumber));
        }

        /// <summary>
        /// Check whether a difference has been pre-approved.
        /// A difference is approved if the workbook:page key exists in the approved list.
        /// </summary>
        public bool IsDifferenceApproved(string workbookName, int pageNumber)
        {
            string key = $"{Path.GetFileNameWithoutExtension(workbookName)}:{pageNumber}";
            return _approvedDifferences.Contains(key);
        }

        /// <summary>
        /// Approve a difference (suppresses failure in validation results).
        /// </summary>
        public void ApproveDifference(string workbookName, int pageNumber)
        {
            string key = $"{Path.GetFileNameWithoutExtension(workbookName)}:{pageNumber}";
            _approvedDifferences.Add(key);
            SaveApprovedDifferences();
        }

        /// <summary>
        /// Create all required directories for the regression test suite.
        /// </summary>
        public void EnsureDirectoriesExist()
        {
            string[] dirs =
            [
                _config.TemplatesDirectory,
                _config.BaselineDirectory,
                _config.CurrentDirectory,
                _config.DiffDirectory,
                _config.ReportsDirectory,
                _config.LogsDirectory
            ];

            foreach (var dir in dirs)
            {
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
        }

        /// <summary>
        /// Ensure a baseline image exists for the given workbook and page.
        /// If auto-create is enabled and the baseline doesn't exist, the current
        /// rendering is saved as the baseline (marking it as the reference).
        /// </summary>
        public string? EnsureBaseline(string workbookName, int pageNumber,
            SkiaSharp.SKBitmap? renderBitmap = null)
        {
            string baselinePath = GetBaselinePath(workbookName, pageNumber);

            if (System.IO.File.Exists(baselinePath))
                return baselinePath;

            if (!_config.AutoCreateBaselines || renderBitmap == null)
                return null;

            // Save the current render as the baseline
            var dir = Path.GetDirectoryName(baselinePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var image = SkiaSharp.SKImage.FromBitmap(renderBitmap);
            using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(baselinePath, data.ToArray());

            return baselinePath;
        }

        // ── Private Helpers ───────────────────────────────────────────

        private HashSet<string> LoadApprovedDifferences()
        {
            try
            {
                if (System.IO.File.Exists(_config.ApprovedDifferencesFile))
                {
                    var json = System.IO.File.ReadAllText(_config.ApprovedDifferencesFile);
                    var list = JsonSerializer.Deserialize<List<string>>(json);
                    return new HashSet<string>(list ?? new List<string>());
                }
            }
            catch
            {
                // If the file is corrupt, start fresh
            }

            return new HashSet<string>();
        }

        private void SaveApprovedDifferences()
        {
            var dir = Path.GetDirectoryName(_config.ApprovedDifferencesFile);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(_approvedDifferences.ToList(), new JsonSerializerOptions
            {
                WriteIndented = true
            });

            System.IO.File.WriteAllText(_config.ApprovedDifferencesFile, json);
        }
    }
}
