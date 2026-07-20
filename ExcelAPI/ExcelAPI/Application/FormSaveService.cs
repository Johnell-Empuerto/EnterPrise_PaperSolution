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
        /// <summary>Validation results from WorkbookDiffValidator (Phase 4.4.1).</summary>
        public WorkbookDiffResult? ValidationResult { get; set; }
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
        private readonly WorkbookValueWriter _valueWriter;
        private readonly WorkbookDiffValidator _diffValidator;
        private readonly ILogger<FormSaveService> _logger;

        public FormSaveService(
            XmlGenerator xmlGenerator,
            DatabaseGenerator databaseGenerator,
            WorkbookGenerator workbookGenerator,
            PreviewGenerator previewGenerator,
            PdfGenerator pdfGenerator,
            WorkbookValueWriter valueWriter,
            WorkbookDiffValidator diffValidator,
            ILogger<FormSaveService> logger)
        {
            _xmlGenerator = xmlGenerator;
            _databaseGenerator = databaseGenerator;
            _workbookGenerator = workbookGenerator;
            _previewGenerator = previewGenerator;
            _pdfGenerator = pdfGenerator;
            _valueWriter = valueWriter;
            _diffValidator = diffValidator;
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

            // Write values into the original workbook using OpenXml
            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"edited_{formId}.xlsx");
            result.CellsWritten = _valueWriter.WriteValues(definition, sourcePath, result.WorkbookPath);
            result.SourcePath = sourcePath;

            _logger.LogInformation(
                "[WBDF] Edited values saved: {Count} cells written to {Path} (source: {Source})",
                result.CellsWritten, result.WorkbookPath, sourcePath);

            // Auto-validate by comparing original vs edited — Phase 4.5 ENFORCEMENT
            _logger.LogInformation("Running WorkbookDiffValidator...");
            var diffResult = _diffValidator.Compare(sourcePath, result.WorkbookPath);
            result.ValidationResult = diffResult;

            _logger.LogInformation("========== VALIDATION RESULTS ==========");
            _logger.LogInformation("Sheet Count Changed: {V}", diffResult.SheetCountChanges);
            _logger.LogInformation("Styles Changed: {V}", diffResult.StyleChanges);
            _logger.LogInformation("Fonts Changed: {V}", diffResult.FontChanges);
            _logger.LogInformation("Fills Changed: {V}", diffResult.FillChanges);
            _logger.LogInformation("Borders Changed: {V}", diffResult.BorderChanges);
            _logger.LogInformation("Alignment Changed: {V}", diffResult.AlignmentChanges);
            _logger.LogInformation("NumberFormat Changed: {V}", diffResult.NumberFormatChanges);
            _logger.LogInformation("Merged Cells Changed: {V}", diffResult.MergeChanges);
            _logger.LogInformation("Print Setup Changed: {V}", diffResult.PageSetupChanges);
            _logger.LogInformation("Page Margins Changed: {V}", diffResult.PageMarginChanges);
            _logger.LogInformation("Freeze Panes Changed: {V}", diffResult.FreezePaneChanges);
            _logger.LogInformation("Row Heights Changed: {V}", diffResult.RowHeightChanges);
            _logger.LogInformation("Column Widths Changed: {V}", diffResult.ColumnWidthChanges);
            _logger.LogInformation("Hidden Rows Changed: {V}", diffResult.HiddenRowChanges);
            _logger.LogInformation("Hidden Columns Changed: {V}", diffResult.HiddenColumnChanges);
            _logger.LogInformation("Drawings Changed: {V}", diffResult.DrawingChanges);
            _logger.LogInformation("Images Changed: {V}", diffResult.ImageChanges);
            _logger.LogInformation("Comments Changed: {V}", diffResult.CommentChanges);
            _logger.LogInformation("Hyperlinks Changed: {V}", diffResult.HyperlinkChanges);
            _logger.LogInformation("Data Validation Changed: {V}", diffResult.DataValidationChanges);
            _logger.LogInformation("Conditional Format Changed: {V}", diffResult.ConditionalFormattingChanges);
            _logger.LogInformation("Tables Changed: {V}", diffResult.TableChanges);
            _logger.LogInformation("Defined Names Changed: {V}", diffResult.DefinedNameChanges);
            _logger.LogInformation("Header/Footer Changed: {V}", diffResult.HeaderFooterChanges);
            _logger.LogInformation("Formula Changed: {V}", diffResult.FormulaChanges);
            _logger.LogInformation("Editable Values Changed: {V}", diffResult.EditableValueChanges);
            _logger.LogInformation("Validation Passed: {V}", diffResult.Passed);

            // Phase 4.5: REJECT on validation failure — never return a corrupted workbook
            if (!diffResult.Passed)
            {
                _logger.LogError("VALIDATION FAILED: {Total} structural differences detected!", diffResult.TotalDifferences);
                foreach (var detail in diffResult.Details)
                {
                    _logger.LogError("  MISMATCH: {Detail}", detail);
                }

                foreach (var cellChange in diffResult.EditableChangedCells)
                {
                    _logger.LogInformation("  CELL EDIT: {Change}", cellChange);
                }

                // Phase 5.5.2 DEBUG: Save a copy before deletion for inspection
                try
                {
                    if (System.IO.File.Exists(result.WorkbookPath))
                    {
                        string debugDir = Path.Combine(Path.GetDirectoryName(result.WorkbookPath) ?? ".", "debug_output");
                        Directory.CreateDirectory(debugDir);
                        string debugPath = Path.Combine(debugDir, $"phase552_debug_{Guid.NewGuid():N}.xlsx");
                        System.IO.File.Copy(result.WorkbookPath, debugPath, overwrite: true);
                        _logger.LogInformation("[PHASE5.5.2 DEBUG] Saved debug output to {Path}", debugPath);
                    }
                }
                catch (Exception exDbg)
                {
                    _logger.LogWarning("[PHASE5.5.2 DEBUG] Failed to save debug copy: {Msg}", exDbg.Message);
                }

                try { if (System.IO.File.Exists(result.WorkbookPath)) System.IO.File.Delete(result.WorkbookPath); } catch { }

                throw new WorkbookFidelityException(diffResult);
            }

            _logger.LogInformation("Validation PASSED — {VC} editable values changed correctly.",
                diffResult.EditableValueChanges);
            foreach (var cellChange in diffResult.EditableChangedCells)
            {
                _logger.LogInformation("  CELL CHANGE: {Change}", cellChange);
            }

            _logger.LogInformation("========== SAVE PIPELINE END ==========");

            return result;
        }
    }
}