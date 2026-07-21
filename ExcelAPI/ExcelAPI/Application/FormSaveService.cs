using ExcelAPI.Designer.Generation;
using ExcelAPI.Models;
using ExcelAPI.Models.WorkbookDefinition;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Application
{
    public class FormSaveResult
    {
        public string XmlPath { get; set; } = "";
        public string WorkbookPath { get; set; } = "";
        public string PreviewPath { get; set; } = "";
        public string PdfPath { get; set; } = "";
        public DatabaseResult? DatabaseObjects { get; set; }
        /// <summary>Phase 19 business-level round-trip compatibility validation result.
        /// Includes field-level, comment, print layout, and configuration comparisons.
        /// Never rejects the workbook. Always returns it.</summary>
        public RoundTripCompatibilityReport? RoundTripReport { get; set; }
        /// <summary>Number of cells written by WorkbookValueWriter.</summary>
        public int CellsWritten { get; set; }
        /// <summary>Source workbook path that was used.</summary>
        public string SourcePath { get; set; } = "";
    }

    public interface IFormSaveService
    {
        Task<FormSaveResult> SaveAsync(FormDefinition form, string outputDirectory, CancellationToken cancellationToken = default);
        Task<FormSaveResult> SaveFromDefinitionAsync(WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default);
        [Obsolete("Use SaveEditedValuesAsync(WorkbookDefinition) instead. This path uses COM WorkbookGenerator.")]
        Task<FormSaveResult> OutputExcelAsync(FormDefinition form, string outputDirectory, CancellationToken cancellationToken = default);
        Task<FormSaveResult> OutputExcelFromDefinitionAsync(WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default);

        /// <summary>
        /// Save edited values — resolves source workbook from WbDef.SourceFileName or SourcePath.
        /// </summary>
        Task<FormSaveResult> SaveEditedValuesAsync(WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default);

        /// <summary>
        /// Save edited values — use the provided sourcePath directly (Phase 5.2: session-resolved path).
        /// This overload is preferred when the controller has already resolved the workbook path
        /// from the session store, avoiding any SourceFileName-based fallback.
        /// </summary>
        Task<FormSaveResult> SaveEditedValuesAsync(WorkbookDefinition definition, string outputDirectory, string sourcePath, CancellationToken cancellationToken = default);
    }

    public class FormSaveService : IFormSaveService
    {
        private readonly XmlGenerator _xmlGenerator;
        private readonly DatabaseGenerator _databaseGenerator;
        private readonly WorkbookGenerator _workbookGenerator;
        private readonly PreviewGenerator _previewGenerator;
        private readonly PdfGenerator _pdfGenerator;
        private readonly ConMasCompatibleWorkbookWriter _valueWriter;
        private readonly WorkbookStyleWriter _styleWriter;
        private readonly CompatibilityValidator _compatValidator;
        private readonly IPaperLessConfigWriter _paperLessConfigWriter;
        private readonly ILogger<FormSaveService> _logger;

        public FormSaveService(
            XmlGenerator xmlGenerator,
            DatabaseGenerator databaseGenerator,
            WorkbookGenerator workbookGenerator,
            PreviewGenerator previewGenerator,
            PdfGenerator pdfGenerator,
            ConMasCompatibleWorkbookWriter valueWriter,
            WorkbookStyleWriter styleWriter,
            CompatibilityValidator compatValidator,
            IPaperLessConfigWriter paperLessConfigWriter,
            ILogger<FormSaveService> logger)
        {
            _xmlGenerator = xmlGenerator;
            _databaseGenerator = databaseGenerator;
            _workbookGenerator = workbookGenerator;
            _previewGenerator = previewGenerator;
            _pdfGenerator = pdfGenerator;
            _valueWriter = valueWriter;
            _styleWriter = styleWriter;
            _compatValidator = compatValidator;
            _paperLessConfigWriter = paperLessConfigWriter;
            _logger = logger;
        }

        /// <summary>
        /// Save a FormDefinition (existing path, unchanged).
        /// </summary>
        public async Task<FormSaveResult> SaveAsync(FormDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            // 1. Generate XML
            cancellationToken.ThrowIfCancellationRequested();
            string xml = _xmlGenerator.Generate(definition);
            result.XmlPath = Path.Combine(outputDirectory, $"form_{formId}.xml");
            await File.WriteAllTextAsync(result.XmlPath, xml, cancellationToken);
            _logger.LogInformation("XML saved to {Path}", result.XmlPath);

            // 2. Generate Database Objects
            cancellationToken.ThrowIfCancellationRequested();
            result.DatabaseObjects = _databaseGenerator.Generate(definition);
            _logger.LogInformation("Database objects generated: {SheetCount} sheets, {ClusterCount} clusters",
                result.DatabaseObjects.Sheets.Count, result.DatabaseObjects.Clusters.Count);

            // 3. Generate Workbook (XLSX) — now includes cell values, comments, and _Fields sheet
            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"form_{formId}.xlsx");
            result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
            _logger.LogInformation("Workbook saved to {Path}", result.WorkbookPath);

            // 4. Preview for each sheet
            cancellationToken.ThrowIfCancellationRequested();
            var firstSheet = definition.Sheets.FirstOrDefault();
            if (firstSheet != null)
            {
                result.PreviewPath = Path.Combine(outputDirectory, $"preview_{formId}.png");
                result.PreviewPath = _previewGenerator.Generate(definition, firstSheet.Id, result.PreviewPath);
                _logger.LogInformation("Preview saved to {Path}", result.PreviewPath);
            }

            // 5. PDF for each sheet
            cancellationToken.ThrowIfCancellationRequested();
            if (firstSheet != null)
            {
                result.PdfPath = Path.Combine(outputDirectory, $"form_{formId}.pdf");
                result.PdfPath = _pdfGenerator.Generate(definition, firstSheet.Id, result.PdfPath);
                _logger.LogInformation("PDF saved to {Path}", result.PdfPath);
            }

            return result;
        }

        /// <summary>
        /// Save from a WorkbookDefinition (canonical model path).
        /// Converts to FormDefinition via the reverse converter then delegates
        /// to the existing SaveAsync(FormDefinition) path.
        /// </summary>
        public async Task<FormSaveResult> SaveFromDefinitionAsync(
            WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            var form = WorkbookDefinitionConverter.ToFormDefinition(definition);
            _logger.LogInformation(
                "[WBDF] Saving from WorkbookDefinition: {Title} ({Sheets} sheets)",
                definition.Info?.Title, definition.Sheets.Count);
            return await SaveAsync(form, outputDirectory, cancellationToken);
        }

        /// <summary>
        /// Generate Output Excel using the legacy COM-based WorkbookGenerator.
        /// Phase 4.6: DEPRECATED — use OutputExcelFromDefinitionAsync (WbDef path) instead.
        /// The WbDef path calls WorkbookGenerator.Generate(WorkbookDefinition), which
        /// internally converts to FormDefinition as a COM implementation detail.
        /// This overload exists only for backward compatibility and will be removed.
        /// </summary>
        [Obsolete("Use OutputExcelFromDefinitionAsync(WorkbookDefinition) instead. This path uses COM WorkbookGenerator.")]
        public async Task<FormSaveResult> OutputExcelAsync(FormDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            _logger.LogWarning("[DEPRECATED] OutputExcelAsync(FormDefinition) called — use OutputExcelFromDefinitionAsync(WbDef) instead.");
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"output_{formId}.xlsx");
            result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
            _logger.LogInformation("[DEPRECATED] Output Excel saved to {Path}", result.WorkbookPath);

            return result;
        }

        /// <summary>
        /// Generate Output Excel from a WorkbookDefinition (canonical model path).
        /// Calls WorkbookGenerator.Generate(WbDef) directly — no FormDefinition round-trip.
        /// </summary>
        public async Task<FormSaveResult> OutputExcelFromDefinitionAsync(
            WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"output_{formId}.xlsx");
            result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
            _logger.LogInformation(
                "[WBDF] Output Excel generated from WorkbookDefinition: {Path} ({Sheets} sheets)",
                result.WorkbookPath, definition.Sheets.Count);

            return result;
        }



        /// <summary>
        /// Save edited values — resolves source workbook from WbDef.SourceFileName or SourcePath.
        /// Phase 5.2: Delegates to the sourcePath overload after resolving the file.
        /// </summary>
        public async Task<FormSaveResult> SaveEditedValuesAsync(
            WorkbookDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            // Try to find the source workbook
            string? sourcePath = null;

            if (!string.IsNullOrEmpty(definition.SourcePath) && System.IO.File.Exists(definition.SourcePath))
            {
                sourcePath = definition.SourcePath;
            }
            else if (!string.IsNullOrEmpty(definition.SourceFileName))
            {
                string candidate = Path.Combine(outputDirectory, definition.SourceFileName);
                if (System.IO.File.Exists(candidate))
                    sourcePath = candidate;
                else
                {
                    candidate = Path.Combine(outputDirectory, $"{Path.GetFileNameWithoutExtension(definition.SourceFileName)}.xlsx");
                    if (System.IO.File.Exists(candidate))
                        sourcePath = candidate;
                }
            }

            if (string.IsNullOrEmpty(sourcePath))
            {
                _logger.LogError("SAVE PIPELINE FAILED: source workbook not found. SourceFileName={File}",
                    definition.SourceFileName);
                throw new FileNotFoundException(
                    "Cannot save edited values: no source workbook found. " +
                    "Use the SaveEditedValuesAsync(definition, outputDir, sourcePath) overload " +
                    "or set WorkbookDefinition.SourceFileName to the original XLSX filename.");
            }

            return await SaveEditedValuesAsync(definition, outputDirectory, sourcePath, cancellationToken);
        }

        /// <summary>
        /// Check if a CellStyle has any actual style properties to apply.
        /// </summary>
        private static bool HasStyleProperties(CellStyle? style)
        {
            if (style == null) return false;
            bool hasFont = style.Font != null &&
                (!string.IsNullOrEmpty(style.Font.Name) ||
                 style.Font.SizePt > 0 ||
                 style.Font.Bold ||
                 style.Font.Italic ||
                 !string.IsNullOrEmpty(style.Font.ColorArgb));
            bool hasFill = style.Fill != null &&
                !string.IsNullOrEmpty(style.Fill.ColorArgb) &&
                style.Fill.PatternType != "none";
            bool hasAlign = style.Alignment != null &&
                (!string.IsNullOrEmpty(style.Alignment.Horizontal) ||
                 !string.IsNullOrEmpty(style.Alignment.Vertical));
            bool hasWrap = style.WrapText;
            return hasFont || hasFill || hasAlign || hasWrap;
        }

        /// <summary>
        /// Save edited values back into the original workbook (Phase 4.4 / 5.2).
        /// Uses WorkbookValueWriter to surgically modify only editable cells
        /// in the original XLSX, preserving all formatting, layouts, and structure.
        ///
        /// Phase 5.2: Accepts a pre-resolved sourcePath (from session store).
        /// The controller resolves the path via ISessionWorkbookStore and passes it here,
        /// eliminating the need for SourceFileName-based fallback.
        ///
        /// === PIPELINE LOGGING ===
        /// WorkbookGenerator invoked: FALSE
        /// WorkbookValueWriter invoked: TRUE
        /// Auto-validation via WorkbookDiffValidator: ENABLED
        /// </summary>
        /// <param name="definition">WorkbookDefinition with edited field values.</param>
        /// <param name="outputDirectory">Directory to save the edited workbook.</param>
        /// <param name="sourcePath">Server-resolved path to the original workbook (from session store).</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task<FormSaveResult> SaveEditedValuesAsync(
            WorkbookDefinition definition, string outputDirectory, string sourcePath, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            // ═════════════════════════════════════════════════════════
            // PHASE 21.4 — STAGE 7: SAVE SERVICE INPUT
            // ═════════════════════════════════════════════════════════
            int serviceFieldCount = definition.Sheets?.Sum(s => s.Fields?.Count ?? 0) ?? 0;
            int serviceWithValues = definition.Sheets?.Sum(s => s.Fields?.Count(f => !string.IsNullOrWhiteSpace(f.Value)) ?? 0) ?? 0;

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("SAVE SERVICE INPUT");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Workbook: {Title}", definition.Info?.Title ?? "(untitled)");
            _logger.LogInformation("SessionId: {SessionId}", definition.SessionId ?? "(not set)");
            _logger.LogInformation("SourcePath: {Path}", sourcePath);
            _logger.LogInformation("Sheets: {Count}", definition.Sheets.Count);
            _logger.LogInformation("Fields (total): {Count}", serviceFieldCount);
            _logger.LogInformation("Fields (with values): {Count}", serviceWithValues);
            _logger.LogInformation("Fields (empty): {Count}", serviceFieldCount - serviceWithValues);

            // Field loss check: compare with controller stage
            if (serviceFieldCount == 0)
            {
                _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                _logger.LogError("FIELD LOSS DETECTED AT STAGE 7: 0 fields entered FormSaveService");
                _logger.LogError("This means the WorkbookDefinition passed from the Controller");
                _logger.LogError("already has 0 fields. The bug is in the Controller call chain.");
                _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
            }
            else if (serviceWithValues == 0 && serviceFieldCount > 0)
            {
                _logger.LogWarning("STAGE 7: {Total} fields exist but ALL have empty values", serviceFieldCount);
                _logger.LogWarning("Workbook will be copied unchanged (0 cells written).");
            }
            _logger.LogInformation("=========================================================");

            _logger.LogInformation("========== SAVE PIPELINE ==========");
            _logger.LogInformation("WorkbookValueWriter START");
            _logger.LogInformation("Workbook: {Title}", definition.Info?.Title ?? "(untitled)");
            _logger.LogInformation("SessionId: {SessionId}", definition.SessionId ?? "(not set)");
            _logger.LogInformation("SourcePath: {Path}", sourcePath);
            _logger.LogInformation("Sheets: {Count}", definition.Sheets.Count);

            // Log field values being written
            int totalFieldsWithValues = 0;
            foreach (var sheet in definition.Sheets)
            {
                foreach (var field in sheet.Fields)
                {
                    if (!string.IsNullOrWhiteSpace(field.Value))
                    {
                        totalFieldsWithValues++;
                        _logger.LogInformation("  {Cell} -> {Value}", field.Cell?.Address ?? "?", field.Value);
                    }
                }
            }
            if (totalFieldsWithValues == 0)
            {
                _logger.LogWarning("  No field values to write — workbook copied unchanged.");
            }

            _logger.LogInformation("WorkbookGenerator invoked: FALSE");
            _logger.LogInformation("WorkbookValueWriter invoked: TRUE");
            _logger.LogInformation("Original workbook: {Path}", sourcePath);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            // ═════════════════════════════════════════════════════════
            // PHASE 21.3 — STAGE 2: WORKBOOK DEFINITION DIAGNOSTIC
            // ═════════════════════════════════════════════════════════

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("WORKBOOK DEFINITION — Before WorkbookValueWriter");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Workbook: {Title}", definition.Info?.Title ?? "(untitled)");
            _logger.LogInformation("SessionId: {Sid}", definition.SessionId ?? "(none)");
            _logger.LogInformation("Pages (sheets): {Count}", definition.Sheets.Count);

            int wbDefTotalFields = 0;
            int wbDefTotalWithValues = 0;
            foreach (var wbSheet in definition.Sheets)
            {
                foreach (var field in wbSheet.Fields)
                {
                    wbDefTotalFields++;
                    if (!string.IsNullOrWhiteSpace(field.Value))
                        wbDefTotalWithValues++;

                    string fieldId = field.Id ?? $"field_{wbDefTotalFields}";
                    string cellAddr = field.Cell?.Address ?? "?";
                    string fieldValue = field.Value ?? "";
                    string isRequired = field.Required ? "true" : "false";
                    string maxLen = field.MaxLength > 0 ? field.MaxLength.ToString() : "(unlimited)";
                    string placeholder = field.Placeholder ?? "";
                    string defaultValue = field.DefaultValue ?? "";
                    string formula = field.Formula ?? "";
                    string styleFont = "";
                    if (field.Style?.Font != null)
                    {
                        var f = field.Style.Font;
                        styleFont = $"Font='{f.Name ?? "?"}/{f.SizePt}' Bold={f.Bold} Color='{f.ColorArgb ?? ""}'";
                    }

                    _logger.LogInformation("  Field #{Count}:", wbDefTotalFields);
                    _logger.LogInformation("    Id        : {Id}", fieldId);
                    _logger.LogInformation("    Page      : {Sheet}", wbSheet.Name);
                    _logger.LogInformation("    Cell      : {Cell}", cellAddr);
                    _logger.LogInformation("    Type      : {Type}", field.Type.ToString());
                    _logger.LogInformation("    Value     : {Val}", string.IsNullOrEmpty(fieldValue) ? "(empty)" : fieldValue);
                    _logger.LogInformation("    Required  : {Req}", isRequired);
                    _logger.LogInformation("    Locked    : {Locked}", field.Locked);
                    _logger.LogInformation("    MaxLength : {Max}", maxLen);
                    _logger.LogInformation("    Placeholder : {PH}", string.IsNullOrEmpty(placeholder) ? "(none)" : placeholder);
                    _logger.LogInformation("    Default   : {Def}", string.IsNullOrEmpty(defaultValue) ? "(none)" : defaultValue);
                    _logger.LogInformation("    Formula   : {F}", string.IsNullOrEmpty(formula) ? "(none)" : formula);
                    _logger.LogInformation("    Style     : {S}", string.IsNullOrEmpty(styleFont) ? "(default)" : styleFont);
                    _logger.LogInformation("    ---------------------------------------------------------");
                }
            }

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Total fields in WorkbookDefinition: {Total}", wbDefTotalFields);
            _logger.LogInformation("Fields with non-empty values: {WithVal}", wbDefTotalWithValues);
            _logger.LogInformation("=========================================================");

            // Write values into the original workbook using OpenXml
            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"edited_{formId}.xlsx");
            result.CellsWritten = _valueWriter.WriteValues(definition, sourcePath, result.WorkbookPath);
            result.SourcePath = sourcePath;

            _logger.LogInformation(
                "[WBDF] Edited values saved: {Count} cells written to {Path} (source: {Source})",
                result.CellsWritten, result.WorkbookPath, sourcePath);

            // Phase 22: Apply browser-edited styles (font, fill, alignment) to cells
            if (definition.Sheets.Any(s => s.Fields.Any(f => f.Style != null && HasStyleProperties(f.Style))))
            {
                int stylesApplied = _styleWriter.ApplyStyles(definition, result.WorkbookPath);
                _logger.LogInformation(
                    "[PHASE22] Browser styles applied: {Count} cells styled",
                    stylesApplied);
            }
            else
            {
                _logger.LogInformation("[PHASE22] No field style overrides found — skipping style writer");
            }

            // ═════════════════════════════════════════════════════════
            // PAPERLESS DEBUG STAGE 3 — Before PaperLessConfigWriter
            // ═════════════════════════════════════════════════════════
            _logger.LogInformation("========================================================");
            _logger.LogInformation("PAPERLESS DEBUG STAGE 3 — FormSaveService Before Config Writer");
            _logger.LogInformation("========================================================");
            if (definition.Sheets != null)
            {
                foreach (var s in definition.Sheets)
                {
                    _logger.LogInformation("  Sheet: {Name}", s.Name ?? "(unnamed)");
                    if (s.Fields != null)
                    {
                        foreach (var f in s.Fields.Take(3))
                        {
                            _logger.LogInformation("    Field ID: {Id}", f.Id ?? "(empty)");
                            _logger.LogInformation("    Cell: {Cell}", f.Cell?.Address ?? "?");
                            _logger.LogInformation("    FontSize (CellStyle): {Sz}", f.Style?.Font?.SizePt ?? 0);
                            _logger.LogInformation("    FontName: {Fn}", f.Style?.Font?.Name ?? "(null)");
                            _logger.LogInformation("    Bold: {B}", f.Style?.Font?.Bold ?? false);
                        }
                    }
                }
            }
            _logger.LogInformation("========================================================");

            // PaperLess config persistence — embed field identity, style, and config
            // as VeryHidden JSON inside the workbook for re-upload restoration.
            _paperLessConfigWriter.WritePaperLessConfig(definition, result.WorkbookPath);

            // Phase 19: Business-level round-trip compatibility validation.
            // Compares fields, comments, configuration, print layout, and background.
            // Never rejects the workbook — always returns it to the user.
            // Matches ConMas behavior: ConMas never performs structural validation.
            _logger.LogInformation("Running Phase 19 CompatibilityValidator...");
            var compatReport = _compatValidator.Validate(
                sourcePath, result.WorkbookPath, result.CellsWritten);
            result.RoundTripReport = compatReport;

            _logger.LogInformation("========== COMPATIBILITY REPORT ==========");
            _logger.LogInformation("Score: {Score}/100", compatReport.CompatibilityScore);
            _logger.LogInformation("Fields: {Orig} → {Edit}",
                compatReport.FieldCountOriginal, compatReport.FieldCountEdited);
            _logger.LogInformation("Errors: {E}, Warnings: {W}",
                compatReport.Errors.Count, compatReport.Warnings.Count);

            foreach (var warning in compatReport.Warnings)
                _logger.LogWarning("  COMPAT WARNING: {Warn}", warning);
            foreach (var error in compatReport.Errors)
                _logger.LogError("  COMPAT ERROR: {Err}", error);

            _logger.LogInformation("========== SAVE PIPELINE END ==========");

            return result;
        }
    }
}