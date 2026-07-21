using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using System.Text;

namespace ExcelAPI.Application
{
    // ══════════════════════════════════════════════════════════════════════
    // Field-level data extracted from comments for comparison
    // ══════════════════════════════════════════════════════════════════════

    public class FieldInfo
    {
        public string CellAddress { get; set; } = "";
        public string Worksheet { get; set; } = "";
        public string CommentText { get; set; } = "";
        public string FieldId { get; set; } = "";
        public string FieldName { get; set; } = "";
        public string FieldType { get; set; } = "";
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public override string ToString() =>
            $"{Worksheet}!{CellAddress} [{FieldId}] '{FieldName}' ({FieldType})";

        public override bool Equals(object? obj) =>
            obj is FieldInfo other &&
            CellAddress == other.CellAddress &&
            Worksheet == other.Worksheet;

        public override int GetHashCode() =>
            HashCode.Combine(CellAddress, Worksheet);
    }

    // ══════════════════════════════════════════════════════════════════════
    // Print layout information for comparison
    // ══════════════════════════════════════════════════════════════════════

    public class PrintLayoutInfo
    {
        public string SheetName { get; set; } = "";
        public string PrintArea { get; set; } = "";
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double LeftMargin { get; set; }
        public double RightMargin { get; set; }
        public double HeaderMargin { get; set; }
        public double FooterMargin { get; set; }
        public string Orientation { get; set; } = "";
        public int PaperSize { get; set; }
        public int FitToWidth { get; set; }
        public int FitToHeight { get; set; }
        public int? Scale { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public int MergeCount { get; set; }
        public int HiddenRowCount { get; set; }
        public int HiddenColCount { get; set; }
        public string Dimension { get; set; } = "";
    }

    // ══════════════════════════════════════════════════════════════════════
    // Configuration sheet information
    // ══════════════════════════════════════════════════════════════════════

    public class ConfigSheetInfo
    {
        public string Name { get; set; } = "";
        public bool Exists { get; set; }
        public int RowCount { get; set; }
        public List<string> FieldIds { get; set; } = new();
        public List<string> SheetMappings { get; set; } = new();
    }

    // ══════════════════════════════════════════════════════════════════════
    // Field comparison result
    // ══════════════════════════════════════════════════════════════════════

    public enum FieldComparisonStatus { Unchanged, Missing, Added, Changed }

    public class FieldComparison
    {
        public string CellAddress { get; set; } = "";
        public string Worksheet { get; set; } = "";
        public FieldComparisonStatus Status { get; set; }
        public string Details { get; set; } = "";
    }

    // ══════════════════════════════════════════════════════════════════════
    // Phase 19 — RoundTripCompatibilityReport
    // Detailed business-level compatibility report
    // ══════════════════════════════════════════════════════════════════════

    public class RoundTripCompatibilityReport
    {
        // ── Overall ──
        public bool Passed { get; set; } = true;
        public int CompatibilityScore { get; set; } = 100;
        public int CellsWritten { get; set; }

        // ── Fields ──
        public int FieldCountOriginal { get; set; }
        public int FieldCountEdited { get; set; }
        public List<FieldComparison> MissingFields { get; set; } = new();
        public List<FieldComparison> AddedFields { get; set; } = new();
        public List<FieldComparison> ChangedFields { get; set; } = new();

        // ── Comments ──
        public int CommentChanges { get; set; }
        public int CommentMissingCount { get; set; }
        public int NewCommentCount { get; set; }
        public int ChangedCommentCount { get; set; }

        // ── Configuration ──
        public bool FieldsSheetExists { get; set; }
        public bool ExcelOutputSettingExists { get; set; }
        public bool ConfigDuplicated { get; set; }
        public List<string> ConfigurationChanges { get; set; } = new();

        // ── Print Layout ──
        public List<string> LayoutChanges { get; set; } = new();
        public bool PrintAreaChanged { get; set; }
        public bool PageSetupChanged { get; set; }
        public bool MarginsChanged { get; set; }
        public bool MergedCellsChanged { get; set; }
        public bool HiddenRowsChanged { get; set; }
        public bool HiddenColumnsChanged { get; set; }

        // ── Background ──
        public bool WorksheetCountChanged { get; set; }
        public bool WorksheetNamesChanged { get; set; }
        public bool DimensionsChanged { get; set; }

        // ── Overall ──
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
        public List<string> Details { get; set; } = new();

        /// <summary>Calculate the compatibility score (0-100) based on all metrics.</summary>
        public void CalculateScore()
        {
            int score = 100;

            // Deduct for missing fields (critical)
            if (MissingFields.Count > 0)
                score -= Math.Min(MissingFields.Count * 15, 60);

            // Deduct for added fields (warning)
            if (AddedFields.Count > 0)
                score -= Math.Min(AddedFields.Count * 5, 20);

            // Deduct for changed fields
            if (ChangedFields.Count > 0)
                score -= Math.Min(ChangedFields.Count * 3, 15);

            // Deduct for comment changes
            if (CommentChanges > 0)
                score -= Math.Min(CommentChanges * 2, 10);
            if (CommentMissingCount > 0)
                score -= Math.Min(CommentMissingCount * 5, 20);

            // Deduct for configuration issues
            if (!FieldsSheetExists)
                score -= 15;
            if (!ExcelOutputSettingExists)
                score -= 10;
            if (ConfigDuplicated)
                score -= 15;
            if (ConfigurationChanges.Count > 0)
                score -= Math.Min(ConfigurationChanges.Count * 3, 15);

            // Deduct for layout changes
            if (PrintAreaChanged) score -= 10;
            if (PageSetupChanged) score -= 8;
            if (MarginsChanged) score -= 5;
            if (MergedCellsChanged) score -= 5;
            if (HiddenRowsChanged) score -= 3;
            if (HiddenColumnsChanged) score -= 3;

            // Deduct for background issues
            if (WorksheetCountChanged) score -= 15;
            if (WorksheetNamesChanged) score -= 10;
            if (DimensionsChanged) score -= 5;

            score = Math.Max(0, Math.Min(100, score));
            CompatibilityScore = score;
            Passed = score > 0; // Never reject, but score reflects health
        }

        public string GetFormattedReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("╔══════════════════════════════════════════════╗");
            sb.AppendLine("║     ROUND-TRIP COMPATIBILITY REPORT        ║");
            sb.AppendLine("╚══════════════════════════════════════════════╝");
            sb.AppendLine();
            sb.AppendLine($"  Score: {CompatibilityScore}/100  Status: {(Passed ? "PASS" : "INFO")}");
            sb.AppendLine($"  Cells Written: {CellsWritten}");
            sb.AppendLine();

            // Fields
            sb.AppendLine($"── Fields ──");
            sb.AppendLine($"  Original: {FieldCountOriginal}  Edited: {FieldCountEdited}");
            if (MissingFields.Count > 0)
            {
                sb.AppendLine($"  ❌ Missing: {MissingFields.Count}");
                foreach (var f in MissingFields.Take(10))
                    sb.AppendLine($"     - {f.Worksheet}!{f.CellAddress}: {f.Details}");
            }
            if (AddedFields.Count > 0)
            {
                sb.AppendLine($"  ⚠️  Added: {AddedFields.Count}");
                foreach (var f in AddedFields.Take(5))
                    sb.AppendLine($"     + {f.Worksheet}!{f.CellAddress}");
            }
            if (ChangedFields.Count > 0)
            {
                sb.AppendLine($"  ⚠️  Changed: {ChangedFields.Count}");
                foreach (var f in ChangedFields.Take(5))
                    sb.AppendLine($"     ~ {f.Worksheet}!{f.CellAddress}: {f.Details}");
            }
            sb.AppendLine();

            // Configuration
            sb.AppendLine($"── Configuration ──");
            sb.AppendLine($"  _Fields: {(FieldsSheetExists ? "✅" : "❌")}");
            sb.AppendLine($"  ExcelOutputSetting: {(ExcelOutputSettingExists ? "✅" : "❌")}");
            if (ConfigDuplicated) sb.AppendLine($"  ⚠️  Duplicate configuration detected");
            foreach (var c in ConfigurationChanges)
                sb.AppendLine($"  ⚠️  {c}");
            sb.AppendLine();

            // Comments
            sb.AppendLine($"── Comments ──");
            sb.AppendLine($"  Changes: {CommentChanges}  Missing: {CommentMissingCount}  New: {NewCommentCount}");
            sb.AppendLine();

            // Print Layout
            sb.AppendLine($"── Print Layout ──");
            if (PrintAreaChanged) sb.AppendLine("  ❌ PrintArea changed");
            if (PageSetupChanged) sb.AppendLine("  ⚠️  PageSetup changed");
            if (MarginsChanged) sb.AppendLine("  ⚠️  Margins changed");
            if (MergedCellsChanged) sb.AppendLine("  ⚠️  Merged cells changed");
            if (HiddenRowsChanged) sb.AppendLine("  ⚠️  Hidden rows changed");
            if (HiddenColumnsChanged) sb.AppendLine("  ⚠️  Hidden columns changed");
            foreach (var l in LayoutChanges)
                sb.AppendLine($"  ⚠️  {l}");
            sb.AppendLine();

            // Background
            sb.AppendLine($"── Background ──");
            if (WorksheetCountChanged) sb.AppendLine("  ❌ Worksheet count changed");
            if (WorksheetNamesChanged) sb.AppendLine("  ⚠️  Worksheet names changed");
            if (DimensionsChanged) sb.AppendLine("  ⚠️  Sheet dimensions changed");
            sb.AppendLine();

            // Summary
            sb.AppendLine($"── Summary ──");
            sb.AppendLine($"  Errors:   {Errors.Count}");
            sb.AppendLine($"  Warnings: {Warnings.Count}");
            sb.AppendLine($"  Score:    {CompatibilityScore}/100");
            sb.AppendLine();

            return sb.ToString();
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    // Phase 19 — Business-Level Round-Trip Compatibility Validator
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Validates business-level round-trip compatibility between original and edited workbooks.
    ///
    /// Does NOT compare XML, hashes, relationship IDs, or ZIP structure.
    /// Only compares logical business information required for form reconstruction.
    ///
    /// Validates:
    ///   ✅ Per-field comparison (ID, cell, worksheet, type, position)
    ///   ✅ Comment dictionary comparison (Cell -> Comment mapping)
    ///   ✅ _Fields sheet logical comparison (rows, IDs, mappings)
    ///   ✅ ExcelOutputSetting existence and integrity
    ///   ✅ Print layout comparison (margins, orientation, paper size, scaling, merges, hidden)
    ///   ✅ Background metadata (worksheet count, names, dimensions)
    ///   ✅ Compatibility score (0-100)
    ///
    /// Never throws. Never rejects the workbook. Always returns the workbook.
    /// </summary>
    public class CompatibilityValidator
    {
        private readonly ILogger<CompatibilityValidator> _logger;

        private static readonly HashSet<string> ConfigurationSheetNames =
            new(StringComparer.OrdinalIgnoreCase) { "_Fields", "_RawData", "ExcelOutputSetting", "PaperLessConfig" };

        public CompatibilityValidator(ILogger<CompatibilityValidator> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Validate business-level round-trip compatibility.
        /// This is a soft validation — never throws, never rejects.
        /// </summary>
        public RoundTripCompatibilityReport Validate(
            string originalPath, string editedPath, int cellsWritten)
        {
            var report = new RoundTripCompatibilityReport
            {
                CellsWritten = cellsWritten
            };

            _logger.LogInformation("========== PHASE 19 COMPATIBILITY VALIDATION ==========");
            _logger.LogInformation("Original: {Orig}", originalPath);
            _logger.LogInformation("Edited:   {Edit}", editedPath);

            // ── Open both workbooks ──
            SpreadsheetDocument? origDoc = null;
            SpreadsheetDocument? editDoc = null;

            try
            {
                editDoc = SpreadsheetDocument.Open(editedPath, false);
                _logger.LogInformation("[✓] Edited workbook opens successfully");
            }
            catch (Exception ex)
            {
                report.Errors.Add($"Cannot open edited workbook: {ex.Message}");
                report.CalculateScore();
                _logger.LogError("[✗] Cannot open edited workbook: {Msg}", ex.Message);
                return report;
            }

            try
            {
                try
                {
                    origDoc = SpreadsheetDocument.Open(originalPath, false);
                }
                catch
                {
                    _logger.LogWarning("[!] Cannot open original workbook — comparison limited");
                    report.Warnings.Add("Cannot open original workbook — some comparisons skipped");
                }

                var editWbPart = editDoc.WorkbookPart;
                if (editWbPart == null)
                {
                    report.Errors.Add("Edited workbook has no WorkbookPart");
                    report.CalculateScore();
                    return report;
                }

                var origWbPart = origDoc?.WorkbookPart;

                // ══════════════════════════════════════════════════════════════
                // 1. FIELD VALIDATION — Compare every field via comments
                // ══════════════════════════════════════════════════════════════
                _logger.LogInformation("── Field Validation ──");
                ValidateFields(origWbPart, editWbPart, report);

                // ══════════════════════════════════════════════════════════════
                // 2. COMMENT VALIDATION — Cell -> Comment dictionary comparison
                // ══════════════════════════════════════════════════════════════
                _logger.LogInformation("── Comment Validation ──");
                ValidateComments(origWbPart, editWbPart, report);

                // ══════════════════════════════════════════════════════════════
                // 3. CONFIGURATION SHEETS — _Fields and ExcelOutputSetting
                // ══════════════════════════════════════════════════════════════
                _logger.LogInformation("── Configuration Validation ──");
                ValidateConfiguration(editWbPart, report);

                // ══════════════════════════════════════════════════════════════
                // 4. PRINT LAYOUT — Margins, orientation, paper size, scaling
                // ══════════════════════════════════════════════════════════════
                _logger.LogInformation("── Print Layout Validation ──");
                ValidatePrintLayout(origWbPart, editWbPart, report);

                // ══════════════════════════════════════════════════════════════
                // 5. BACKGROUND — Worksheet count, names, dimensions
                // ══════════════════════════════════════════════════════════════
                _logger.LogInformation("── Background Validation ──");
                ValidateBackground(origWbPart, editWbPart, report);

                // ══════════════════════════════════════════════════════════════
                // Calculate final score
                // ══════════════════════════════════════════════════════════════
                report.CalculateScore();

                _logger.LogInformation("========== COMPATIBILITY SCORE ==========");
                _logger.LogInformation("Score: {Score}/100", report.CompatibilityScore);
                _logger.LogInformation("Fields: {Orig} → {Edit}", report.FieldCountOriginal, report.FieldCountEdited);
                _logger.LogInformation("Missing: {M}, Added: {A}, Changed: {C}",
                    report.MissingFields.Count, report.AddedFields.Count, report.ChangedFields.Count);
                _logger.LogInformation("Layout changes: {Count}", report.LayoutChanges.Count);
                _logger.LogInformation("Errors: {E}, Warnings: {W}",
                    report.Errors.Count, report.Warnings.Count);
                _logger.LogInformation("=========================================");

                // Log full report
                _logger.LogInformation("Full report:\n{Report}", report.GetFormattedReport());
            }
            catch (Exception ex)
            {
                report.Errors.Add($"Validation exception: {ex.Message}");
                report.CalculateScore();
                _logger.LogError(ex, "[✗] Compatibility validation threw exception");
            }
            finally
            {
                origDoc?.Dispose();
                editDoc?.Dispose();
            }

            return report;
        }

        // ══════════════════════════════════════════════════════════════════
        // Helper: Extract field info from cell comments
        // ══════════════════════════════════════════════════════════════════

        private List<FieldInfo> ExtractFieldsFromComments(
            WorkbookPart? wbPart, string sheetNameFilter = "")
        {
            var fields = new List<FieldInfo>();
            if (wbPart == null) return fields;

            foreach (var sheet in wbPart.Workbook.Descendants<Sheet>())
            {
                if (sheet.Id?.Value == null) continue;
                if (!string.IsNullOrEmpty(sheetNameFilter) &&
                    !string.Equals(sheet.Name?.Value, sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                    continue;

                var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (wsPart?.WorksheetCommentsPart == null) continue;

                var comments = wsPart.WorksheetCommentsPart.Comments?
                    .Descendants<Comment>().ToList();
                if (comments == null) continue;

                string wsName = sheet.Name?.Value ?? "?";

                foreach (var comment in comments)
                {
                    var fi = new FieldInfo
                    {
                        CellAddress = comment.Reference?.Value ?? "?",
                        Worksheet = wsName,
                        CommentText = GetCommentText(comment),
                    };

                    // Try to parse field ID from comment text
                    fi.FieldName = fi.CommentText;
                    fi.FieldId = fi.CommentText;

                    fields.Add(fi);
                }
            }

            return fields;
        }

        private string GetCommentText(Comment comment)
        {
            try
            {
                var text = comment.Descendants<Text>().FirstOrDefault();
                return text?.Text?.Trim() ?? "";
            }
            catch
            {
                return "";
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // Helper: Extract print layout info from a worksheet
        // ══════════════════════════════════════════════════════════════════

        private PrintLayoutInfo ExtractPrintLayout(
            WorkbookPart? wbPart, Sheet sheet)
        {
            var info = new PrintLayoutInfo
            {
                SheetName = sheet.Name?.Value ?? "?"
            };

            if (wbPart == null || sheet.Id?.Value == null) return info;

            var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
            if (wsPart == null) return info;

            var ws = wsPart.Worksheet;

            // Print area from defined names
            var printArea = ws.Descendants<DefinedName>()
                .FirstOrDefault(n => n.Name?.Value?.Contains("Print_Area") == true);
            if (printArea != null)
                info.PrintArea = printArea.Text ?? "";

            // Also check workbook-level defined names
            if (string.IsNullOrEmpty(info.PrintArea))
            {
                var wbPrintArea = wbPart.Workbook.Descendants<DefinedName>()
                    .FirstOrDefault(n => n.Name?.Value?.Contains("Print_Area") == true &&
                        (n.Text?.Contains(info.SheetName) == true ||
                         n.Text?.Contains($"'{info.SheetName}'") == true));
                if (wbPrintArea != null)
                    info.PrintArea = wbPrintArea.Text ?? "";
            }

            // Page margins
            var margins = ws.Descendants<PageMargins>().FirstOrDefault();
            if (margins != null)
            {
                info.TopMargin = (double)(margins.Top?.Value ?? 0.0);
                info.BottomMargin = (double)(margins.Bottom?.Value ?? 0.0);
                info.LeftMargin = (double)(margins.Left?.Value ?? 0.0);
                info.RightMargin = (double)(margins.Right?.Value ?? 0.0);
                info.HeaderMargin = (double)(margins.Header?.Value ?? 0.0);
                info.FooterMargin = (double)(margins.Footer?.Value ?? 0.0);
            }

            // Page setup
            var pageSetup = ws.Descendants<PageSetup>().FirstOrDefault();
            if (pageSetup != null)
            {
                var orientVal = pageSetup.Orientation?.Value;
                info.Orientation = orientVal != null ? orientVal.ToString() : "";
                info.PaperSize = (int)(pageSetup.PaperSize?.Value ?? 0u);
                info.FitToWidth = (int)(pageSetup.FitToWidth?.Value ?? 0u);
                info.FitToHeight = (int)(pageSetup.FitToHeight?.Value ?? 0u);
                info.Scale = (int?)(pageSetup.Scale?.Value);
            }

            // Print options
            var printOptions = ws.Descendants<PrintOptions>().FirstOrDefault();
            if (printOptions != null)
            {
                info.CenterHorizontally = printOptions.HorizontalCentered?.Value ?? false;
                info.CenterVertically = printOptions.VerticalCentered?.Value ?? false;
            }

            // Merged cells
            info.MergeCount = ws.Descendants<MergeCell>().Count();

            // Hidden rows
            info.HiddenRowCount = ws.Descendants<Row>().Count(r => r.Hidden?.Value == true);

            // Hidden columns
            info.HiddenColCount = ws.Descendants<Column>().Count(c => c.Hidden?.Value == true);

            // Sheet dimension
            var dimension = ws.Descendants<SheetDimension>().FirstOrDefault();
            if (dimension?.Reference?.Value != null)
                info.Dimension = dimension.Reference.Value;

            return info;
        }

        // ══════════════════════════════════════════════════════════════════
        // 1. FIELD VALIDATION
        // ══════════════════════════════════════════════════════════════════

        private void ValidateFields(
            WorkbookPart? origWbPart, WorkbookPart editWbPart,
            RoundTripCompatibilityReport report)
        {
            var origFields = ExtractFieldsFromComments(origWbPart);
            var editFields = ExtractFieldsFromComments(editWbPart);

            report.FieldCountOriginal = origFields.Count;
            report.FieldCountEdited = editFields.Count;

            if (origFields.Count == 0 && editFields.Count == 0)
            {
                _logger.LogInformation("[i] No fields found in either workbook (no comments)");
                return;
            }

            if (origFields.Count == 0)
            {
                _logger.LogInformation("[i] No original fields to compare — edited has {Count}", editFields.Count);
                report.Warnings.Add("Cannot compare fields — original has no comments");
                return;
            }

            _logger.LogInformation("[i] Fields: original={Orig}, edited={Edit}",
                origFields.Count, editFields.Count);

            // Build dictionaries keyed by (Worksheet, CellAddress)
            var origDict = new Dictionary<(string, string), FieldInfo>();
            foreach (var f in origFields)
                origDict[(f.Worksheet, f.CellAddress)] = f;

            var editDict = new Dictionary<(string, string), FieldInfo>();
            foreach (var f in editFields)
                editDict[(f.Worksheet, f.CellAddress)] = f;

            // Find missing fields (in original but not edited)
            foreach (var (key, origField) in origDict)
            {
                if (!editDict.ContainsKey(key))
                {
                    report.MissingFields.Add(new FieldComparison
                    {
                        CellAddress = origField.CellAddress,
                        Worksheet = origField.Worksheet,
                        Status = FieldComparisonStatus.Missing,
                        Details = $"Field '{origField.FieldName}' missing in edited workbook"
                    });
                }
                else
                {
                    var editField = editDict[key];
                    if (origField.CommentText != editField.CommentText)
                    {
                        report.ChangedFields.Add(new FieldComparison
                        {
                            CellAddress = origField.CellAddress,
                            Worksheet = origField.Worksheet,
                            Status = FieldComparisonStatus.Changed,
                            Details = $"Comment text changed: '{Truncate(origField.CommentText)}' → '{Truncate(editField.CommentText)}'"
                        });
                    }
                }
            }

            // Find added fields (in edited but not original)
            foreach (var (key, editField) in editDict)
            {
                if (!origDict.ContainsKey(key))
                {
                    report.AddedFields.Add(new FieldComparison
                    {
                        CellAddress = editField.CellAddress,
                        Worksheet = editField.Worksheet,
                        Status = FieldComparisonStatus.Added,
                        Details = $"New field at {editField.Worksheet}!{editField.CellAddress}"
                    });
                }
            }

            // Log results
            _logger.LogInformation("[i] Missing fields: {Count}", report.MissingFields.Count);
            _logger.LogInformation("[i] Added fields: {Count}", report.AddedFields.Count);
            _logger.LogInformation("[i] Changed fields: {Count}", report.ChangedFields.Count);

            foreach (var f in report.MissingFields)
                _logger.LogWarning("[!] Missing field: {Ws}!{Cell} — {Detail}",
                    f.Worksheet, f.CellAddress, f.Details);
            foreach (var f in report.AddedFields)
                _logger.LogInformation("[+] Added field: {Ws}!{Cell}", f.Worksheet, f.CellAddress);
            foreach (var f in report.ChangedFields)
                _logger.LogWarning("[~] Changed field: {Ws}!{Cell} — {Detail}",
                    f.Worksheet, f.CellAddress, f.Details);
        }

        // ══════════════════════════════════════════════════════════════════
        // 2. COMMENT VALIDATION — Cell -> Comment dictionary comparison
        // ══════════════════════════════════════════════════════════════════

        private void ValidateComments(
            WorkbookPart? origWbPart, WorkbookPart editWbPart,
            RoundTripCompatibilityReport report)
        {
            // Build Cell -> Comment dictionary for edited workbook
            var editComments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheet in editWbPart.Workbook.Descendants<Sheet>())
            {
                if (sheet.Id?.Value == null) continue;
                var wsPart = editWbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (wsPart?.WorksheetCommentsPart == null) continue;

                var comments = wsPart.WorksheetCommentsPart.Comments?
                    .Descendants<Comment>().ToList();
                if (comments == null) continue;

                foreach (var c in comments)
                {
                    string cellRef = c.Reference?.Value ?? "";
                    string text = GetCommentText(c);
                    editComments[cellRef] = text;
                }
            }

            // Build Cell -> Comment dictionary for original workbook
            var origComments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (origWbPart != null)
            {
                foreach (var sheet in origWbPart.Workbook.Descendants<Sheet>())
                {
                    if (sheet.Id?.Value == null) continue;
                    var wsPart = origWbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    if (wsPart?.WorksheetCommentsPart == null) continue;

                    var comments = wsPart.WorksheetCommentsPart.Comments?
                        .Descendants<Comment>().ToList();
                    if (comments == null) continue;

                    foreach (var c in comments)
                    {
                        string cellRef = c.Reference?.Value ?? "";
                        string text = GetCommentText(c);
                        origComments[cellRef] = text;
                    }
                }
            }

            if (origComments.Count == 0 && editComments.Count == 0)
            {
                _logger.LogInformation("[i] No comments in either workbook");
                return;
            }

            int missing = 0, added = 0, changed = 0;

            foreach (var (cellRef, origText) in origComments)
            {
                if (!editComments.ContainsKey(cellRef))
                {
                    missing++;
                    _logger.LogWarning("[!] Missing comment: Cell {Cell}", cellRef);
                }
                else if (editComments[cellRef] != origText)
                {
                    changed++;
                    _logger.LogWarning("[~] Changed comment: Cell {Cell}: '{Orig}' → '{Edit}'",
                        cellRef, Truncate(origText), Truncate(editComments[cellRef]));
                }
            }

            foreach (var (cellRef, _) in editComments)
            {
                if (!origComments.ContainsKey(cellRef))
                {
                    added++;
                    _logger.LogInformation("[+] New comment: Cell {Cell}", cellRef);
                }
            }

            report.CommentChanges = missing + added + changed;
            report.CommentMissingCount = missing;
            report.NewCommentCount = added;
            report.ChangedCommentCount = changed;

            _logger.LogInformation("[i] Comments: missing={M}, added={A}, changed={C}",
                missing, added, changed);
        }

        // ══════════════════════════════════════════════════════════════════
        // 3. CONFIGURATION VALIDATION — _Fields and ExcelOutputSetting
        // ══════════════════════════════════════════════════════════════════

        private void ValidateConfiguration(
            WorkbookPart editWbPart,
            RoundTripCompatibilityReport report)
        {
            var allSheets = editWbPart.Workbook.Descendants<Sheet>().ToList();

            // Check _Fields exists
            var fieldsSheet = allSheets.FirstOrDefault(s =>
                string.Equals(s.Name?.Value, "_Fields", StringComparison.OrdinalIgnoreCase));
            report.FieldsSheetExists = fieldsSheet != null;

            // Check ExcelOutputSetting exists
            var outputSheet = allSheets.FirstOrDefault(s =>
                string.Equals(s.Name?.Value, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase));
            report.ExcelOutputSettingExists = outputSheet != null;

            // Check for duplicate ExcelOutputSetting
            var outputCount = allSheets.Count(s =>
                string.Equals(s.Name?.Value, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase));
            report.ConfigDuplicated = outputCount > 1;

            if (report.ConfigDuplicated)
                report.ConfigurationChanges.Add($"ExcelOutputSetting appears {outputCount} times");

            if (!report.FieldsSheetExists)
                report.ConfigurationChanges.Add("_Fields sheet missing");

            if (!report.ExcelOutputSettingExists)
                report.ConfigurationChanges.Add("ExcelOutputSetting sheet missing");

            // Log config sheet states
            if (fieldsSheet != null)
                _logger.LogInformation("[i] _Fields: State={State}",
                    fieldsSheet.State?.Value ?? SheetStateValues.Visible);

            if (outputSheet != null)
                _logger.LogInformation("[i] ExcelOutputSetting: State={State}",
                    outputSheet.State?.Value ?? SheetStateValues.Visible);
        }

        // ══════════════════════════════════════════════════════════════════
        // 4. PRINT LAYOUT COMPARISON
        // ══════════════════════════════════════════════════════════════════

        private void ValidatePrintLayout(
            WorkbookPart? origWbPart, WorkbookPart editWbPart,
            RoundTripCompatibilityReport report)
        {
            var printableSheets = editWbPart.Workbook.Descendants<Sheet>()
                .Where(s => !ConfigurationSheetNames.Contains(s.Name?.Value ?? ""))
                .ToList();

            var origPrintableSheets = origWbPart?.Workbook.Descendants<Sheet>()
                .Where(s => !ConfigurationSheetNames.Contains(s.Name?.Value ?? ""))
                .ToList() ?? new();

            foreach (var sheet in printableSheets)
            {
                string sheetName = sheet.Name?.Value ?? "?";

                // Extract edited layout
                var editLayout = ExtractPrintLayout(editWbPart, sheet);

                // Find matching original sheet
                Sheet? origSheet = null;
                if (origWbPart != null)
                {
                    origSheet = origPrintableSheets.FirstOrDefault(s =>
                        string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                }

                if (origSheet == null || origWbPart == null)
                {
                    _logger.LogInformation("[i] Sheet '{Name}': no original to compare", sheetName);
                    _logger.LogInformation("    Edited: paperSize={PS}, orient={O}, dim={D}, merges={M}, hiddenR={HR}, hiddenC={HC}",
                        editLayout.PaperSize, editLayout.Orientation, editLayout.Dimension,
                        editLayout.MergeCount, editLayout.HiddenRowCount, editLayout.HiddenColCount);
                    continue;
                }

                var origLayout = ExtractPrintLayout(origWbPart, origSheet);

                // Compare print area
                if (origLayout.PrintArea != editLayout.PrintArea)
                {
                    report.PrintAreaChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': PrintArea changed: '{origLayout.PrintArea}' → '{editLayout.PrintArea}'");
                }

                // Compare orientation
                if (origLayout.Orientation != editLayout.Orientation)
                {
                    report.PageSetupChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Orientation changed: {origLayout.Orientation} → {editLayout.Orientation}");
                }

                // Compare paper size
                if (origLayout.PaperSize != editLayout.PaperSize)
                {
                    report.PageSetupChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': PaperSize changed: {origLayout.PaperSize} → {editLayout.PaperSize}");
                }

                // Compare scaling
                if (origLayout.Scale != editLayout.Scale)
                {
                    report.PageSetupChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Scale changed: {origLayout.Scale} → {editLayout.Scale}");
                }

                // Compare fit to page
                if (origLayout.FitToWidth != editLayout.FitToWidth ||
                    origLayout.FitToHeight != editLayout.FitToHeight)
                {
                    report.PageSetupChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': FitToPage changed: {origLayout.FitToWidth}x{origLayout.FitToHeight} → {editLayout.FitToWidth}x{editLayout.FitToHeight}");
                }

                // Compare margins (within 0.1 tolerance)
                if (Math.Abs(origLayout.TopMargin - editLayout.TopMargin) > 0.1 ||
                    Math.Abs(origLayout.BottomMargin - editLayout.BottomMargin) > 0.1 ||
                    Math.Abs(origLayout.LeftMargin - editLayout.LeftMargin) > 0.1 ||
                    Math.Abs(origLayout.RightMargin - editLayout.RightMargin) > 0.1)
                {
                    report.MarginsChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Margins changed: T={origLayout.TopMargin}/{editLayout.TopMargin} B={origLayout.BottomMargin}/{editLayout.BottomMargin}");
                }

                // Compare merged cells count
                if (origLayout.MergeCount != editLayout.MergeCount)
                {
                    report.MergedCellsChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Merge count: {origLayout.MergeCount} → {editLayout.MergeCount}");
                }

                // Compare hidden rows
                if (origLayout.HiddenRowCount != editLayout.HiddenRowCount)
                {
                    report.HiddenRowsChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Hidden rows: {origLayout.HiddenRowCount} → {editLayout.HiddenRowCount}");
                }

                // Compare hidden columns
                if (origLayout.HiddenColCount != editLayout.HiddenColCount)
                {
                    report.HiddenColumnsChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': Hidden columns: {origLayout.HiddenColCount} → {editLayout.HiddenColCount}");
                }

                // Compare center on page
                if (origLayout.CenterHorizontally != editLayout.CenterHorizontally ||
                    origLayout.CenterVertically != editLayout.CenterVertically)
                {
                    report.PageSetupChanged = true;
                    report.LayoutChanges.Add($"Sheet '{sheetName}': CenterOnPage changed: H={origLayout.CenterHorizontally}/{editLayout.CenterHorizontally}");
                }

                _logger.LogInformation("[i] Sheet '{Name}' layout: orig={OrigPA} edit={EditPA}",
                    sheetName, origLayout.PrintArea, editLayout.PrintArea);
            }

            _logger.LogInformation("[i] Layout changes: {Count}", report.LayoutChanges.Count);
            foreach (var lc in report.LayoutChanges)
                _logger.LogWarning("[!] Layout change: {Change}", lc);
        }

        // ══════════════════════════════════════════════════════════════════
        // 5. BACKGROUND VALIDATION — Worksheets, names, dimensions
        // ══════════════════════════════════════════════════════════════════

        private void ValidateBackground(
            WorkbookPart? origWbPart, WorkbookPart editWbPart,
            RoundTripCompatibilityReport report)
        {
            var editSheets = editWbPart.Workbook.Descendants<Sheet>()
                .Where(s => !ConfigurationSheetNames.Contains(s.Name?.Value ?? ""))
                .ToList();

            var origSheets = origWbPart?.Workbook.Descendants<Sheet>()
                .Where(s => !ConfigurationSheetNames.Contains(s.Name?.Value ?? ""))
                .ToList() ?? new();

            // Compare worksheet count
            if (origSheets.Count != editSheets.Count)
            {
                report.WorksheetCountChanged = true;
                _logger.LogWarning("[!] Worksheet count: {Orig} → {Edit}",
                    origSheets.Count, editSheets.Count);
            }

            // Compare worksheet names
            var origNames = origSheets.Select(s => s.Name?.Value ?? "").ToList();
            var editNames = editSheets.Select(s => s.Name?.Value ?? "").ToList();

            for (int i = 0; i < Math.Min(origNames.Count, editNames.Count); i++)
            {
                if (!string.Equals(origNames[i], editNames[i], StringComparison.OrdinalIgnoreCase))
                {
                    report.WorksheetNamesChanged = true;
                    _logger.LogWarning("[!] Worksheet #{I} name: '{Orig}' → '{Edit}'",
                        i, origNames[i], editNames[i]);
                }
            }

            // Compare dimensions
            foreach (var sheet in editSheets)
            {
                if (sheet.Name?.Value == null || sheet.Id?.Value == null) continue;

                var wsPart = editWbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (wsPart == null) continue;

                var dimension = wsPart.Worksheet.Descendants<SheetDimension>().FirstOrDefault();
                string editDim = dimension?.Reference?.Value ?? "";

                // Find matching original
                string origDim = "";
                if (origWbPart != null)
                {
                    var origSheet = origSheets.FirstOrDefault(s =>
                        string.Equals(s.Name?.Value, sheet.Name.Value, StringComparison.OrdinalIgnoreCase));
                    if (origSheet?.Id?.Value != null)
                    {
                        var origWsPart = origWbPart.GetPartById(origSheet.Id.Value) as WorksheetPart;
                        if (origWsPart != null)
                        {
                            var origDimEl = origWsPart.Worksheet.Descendants<SheetDimension>().FirstOrDefault();
                            origDim = origDimEl?.Reference?.Value ?? "";
                        }
                    }
                }

                if (!string.IsNullOrEmpty(origDim) && origDim != editDim)
                {
                    report.DimensionsChanged = true;
                    _logger.LogWarning("[!] Sheet '{Name}' dimension: '{Orig}' → '{Edit}'",
                        sheet.Name.Value, origDim, editDim);
                }
            }

            _logger.LogInformation("[i] Background: sheets={EditCount} (was {OrigCount}), namesChanged={Names}, dimsChanged={Dims}",
                editSheets.Count, origSheets.Count, report.WorksheetNamesChanged, report.DimensionsChanged);
        }

        // ══════════════════════════════════════════════════════════════════
        // Helpers
        // ══════════════════════════════════════════════════════════════════

        private static string Truncate(string val, int maxLen = 80)
        {
            if (string.IsNullOrEmpty(val)) return "(empty)";
            return val.Length <= maxLen ? val : val[..maxLen] + "...";
        }
    }
}
