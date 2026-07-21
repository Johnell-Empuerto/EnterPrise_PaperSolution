using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelAPI.Models;
using Microsoft.Extensions.Logging;
using System.Text;
using System.Text.Json;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Reads an Excel workbook and reconstructs the complete DesignerModel directly
    /// from the OOXML package. This is the deserialization half of the Phase 20/21
    /// architecture: the workbook is the project file, and DesignerModelReader
    /// deserializes it back into the designer state.
    ///
    /// Does NOT use COM Interop. Does NOT depend on original template.
    /// Reads everything directly from the workbook using OpenXML SDK.
    /// </summary>
    public interface IDesignerModelReader
    {
        /// <summary>
        /// Read a workbook and reconstruct the DesignerModel.
        /// Returns null if the workbook cannot be read.
        /// </summary>
        DesignerModel? Read(string filePath, string fileName, string sessionId);
    }

    /// <summary>
    /// Phase 21/21.1: Reads an Excel workbook and reconstructs the DesignerModel.
    /// Includes comprehensive debug logging (Phase 21.1) for round-trip verification.
    ///
    /// Every log line is prefixed with [DesignerModelReader].
    /// Sections are clearly labeled for easy console reading.
    /// </summary>
    public class DesignerModelReader : IDesignerModelReader
    {
        private readonly ILogger<DesignerModelReader> _logger;

        private static readonly HashSet<string> SkipSheetNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "_RawData"
        };

        private static readonly HashSet<string> ConfigSheetNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "_Fields", "ExcelOutputSetting", "PaperLessConfig"
        };

        private static readonly JsonSerializerOptions PaperLessJsonOptions = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        public DesignerModelReader(ILogger<DesignerModelReader> logger)
        {
            _logger = logger;
        }

        public DesignerModel? Read(string filePath, string fileName, string sessionId)
        {
            var log = new StringBuilder();

            // ═══════════════════════════════════════════════════════════════
            // PHASE 21.1: LOG START
            // ═══════════════════════════════════════════════════════════════
            log.AppendLine("========================================================");
            log.AppendLine("[DesignerModelReader] Started");
            log.AppendLine($"[DesignerModelReader] Workbook: {fileName}");
            log.AppendLine($"[DesignerModelReader] Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            log.AppendLine("========================================================");

            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var wbPart = doc.WorkbookPart;
                if (wbPart == null)
                {
                    _logger.LogError("[DesignerModelReader] Workbook has no WorkbookPart");
                    return null;
                }

                var model = new DesignerModel
                {
                    SessionId = sessionId
                };

                // ═════════════════════════════════════════════════════════
                // 1. WORKBOOK INFORMATION
                // ═════════════════════════════════════════════════════════
                ReadWorkbookInfo(doc, wbPart, model, log);

                // ═════════════════════════════════════════════════════════
                // 2. WORKSHEETS — enumerate all
                // ═════════════════════════════════════════════════════════
                var allSheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                log.AppendLine();
                log.AppendLine("[DesignerModelReader] Worksheets Found");
                log.AppendLine("----------------------------------------");
                foreach (var sheet in allSheets)
                {
                    string skipReason = "";
                    if (SkipSheetNames.Contains(sheet.Name?.Value ?? ""))
                        skipReason = " (SKIPPED — reserved)";
                    log.AppendLine($"  [{allSheets.IndexOf(sheet) + 1}] {sheet.Name?.Value ?? "?"}{skipReason}");
                }

                var printableSheets = allSheets
                    .Where(s => !SkipSheetNames.Contains(s.Name?.Value ?? ""))
                    .Where(s => !ConfigSheetNames.Contains(s.Name?.Value ?? ""))
                    .ToList();
                var configSheets = allSheets
                    .Where(s => ConfigSheetNames.Contains(s.Name?.Value ?? ""))
                    .ToList();

                log.AppendLine();
                log.AppendLine($"[DesignerModelReader] Printable Sheets: {printableSheets.Count}");
                log.AppendLine($"[DesignerModelReader] Config Sheets: {configSheets.Count}");

                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 7 — Re-upload XLSX Detected
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 7 — Re-upload XLSX Detected");
                _logger.LogInformation("========================================================");
                _logger.LogInformation("Total worksheets: {Count}", allSheets.Count);
                foreach (var s in allSheets)
                {
                    _logger.LogInformation("  Sheet: {Name}", s.Name?.Value ?? "?");
                }
                bool hasPaperLessConfig = configSheets.Any(s =>
                    string.Equals(s.Name?.Value, "PaperLessConfig", StringComparison.OrdinalIgnoreCase));
                _logger.LogInformation("PaperLessConfig sheet present: {Has}", hasPaperLessConfig);
                _logger.LogInformation("Printable sheets: {Count}", printableSheets.Count);
                _logger.LogInformation("Config sheets: {Count}", configSheets.Count);
                _logger.LogInformation("========================================================");

                // ═════════════════════════════════════════════════════════
                // 10. _FIELDS SHEET — authoritative metadata source
                // ═════════════════════════════════════════════════════════
                var fieldsData = new List<(string cellAddr, string fieldId, string name, string type,
                    string required, string readOnly, string defaultValue, string maxLength,
                    string placeholder, string options, string validation, string sheetName)>();

                var fieldsSheet = configSheets.FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, "_Fields", StringComparison.OrdinalIgnoreCase));

                if (fieldsSheet != null && fieldsSheet.Id?.Value != null)
                {
                    model.Configuration.HasFieldsSheet = true;
                    fieldsData = ReadFieldsSheet(wbPart, fieldsSheet, log);
                    model.Configuration.FieldsSheetRowCount = fieldsData.Count;
                    model.Configuration.FieldsSheetFieldIds = fieldsData
                        .Select(f => f.fieldId)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .ToList();
                }
                else
                {
                    log.AppendLine();
                    log.AppendLine("[DesignerModelReader] WARNING: _Fields sheet NOT FOUND");
                    _logger.LogWarning("[DesignerModelReader] _Fields sheet not found");
                }

                // ═════════════════════════════════════════════════════════
                // 11. EXCELOUTPUTSETTING
                // ═════════════════════════════════════════════════════════
                var outputSheet = configSheets.FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase));
                if (outputSheet != null)
                {
                    model.Configuration.HasExcelOutputSetting = true;
                    log.AppendLine();
                    log.AppendLine("[DesignerModelReader] Reading ExcelOutputSetting...");
                    log.AppendLine("----------------------------------------");
                    ReadExcelOutputSetting(wbPart, outputSheet, model, log);
                }
                else
                {
                    log.AppendLine();
                    log.AppendLine("[DesignerModelReader] WARNING: ExcelOutputSetting NOT FOUND");
                    _logger.LogWarning("[DesignerModelReader] ExcelOutputSetting not found");
                }

                // Check for duplicate ExcelOutputSetting
                int outputCount = configSheets.Count(s =>
                    string.Equals(s.Name?.Value, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase));
                model.Configuration.IsDuplicated = outputCount > 1;
                if (model.Configuration.IsDuplicated)
                {
                    log.AppendLine("[DesignerModelReader] WARNING: ExcelOutputSetting appears multiple times");
                }

                // ═════════════════════════════════════════════════════════
                // 12. PAPERLESSCONFIG — embedded PaperLess JSON configuration
                // ═════════════════════════════════════════════════════════
                WbDef.PaperLessConfig? paperLessConfig = null;
                var paperLessConfigSheet = configSheets.FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, "PaperLessConfig", StringComparison.OrdinalIgnoreCase));
                if (paperLessConfigSheet != null && paperLessConfigSheet.Id?.Value != null)
                {
                    log.AppendLine();
                    log.AppendLine("[PAPERLESS CONFIG] Configuration sheet found");
                    log.AppendLine("----------------------------------------");

                    try
                    {
                        var plWsPart = wbPart.GetPartById(paperLessConfigSheet.Id.Value) as WorksheetPart;
                        if (plWsPart != null)
                        {
                            var plSheetData = plWsPart.Worksheet
                                .Descendants<Row>()
                                .FirstOrDefault()? // Row 1
                                .Descendants<Cell>()
                                .ToList();

                            // Config JSON is in B1 (second cell)
                            var jsonCell = plSheetData?.FirstOrDefault(c =>
                                string.Equals(c.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                            if (jsonCell != null)
                            {
                                string? jsonText = GetCellInlineText(jsonCell);
                                if (!string.IsNullOrWhiteSpace(jsonText))
                                {
                                    // ═════════════════════════════════════════════════════════
                                    // PAPERLESS DEBUG STAGE 8 — Raw PaperLessConfig JSON
                                    // ═════════════════════════════════════════════════════════
                                    _logger.LogInformation("========================================================");
                                    _logger.LogInformation("PAPERLESS DEBUG STAGE 8 — Raw PaperLessConfig JSON");
                                    _logger.LogInformation("========================================================");
                                    int p1f1Raw = jsonText.IndexOf("\"id\":\"p1f1\"", StringComparison.OrdinalIgnoreCase);
                                    if (p1f1Raw >= 0)
                                    {
                                        int start = Math.Max(0, p1f1Raw - 80);
                                        int len = Math.Min(jsonText.Length - start, 400);
                                        _logger.LogInformation("p1f1 in raw JSON: {Ctx}", jsonText.Substring(start, len));
                                    }
                                    else
                                    {
                                        _logger.LogInformation("p1f1 NOT FOUND in raw JSON — first 400 chars:");
                                        _logger.LogInformation("{Ctx}", jsonText.Length <= 400 ? jsonText : jsonText[..400] + "...");
                                    }
                                    _logger.LogInformation("Raw JSON length: {Len}", jsonText.Length);
                                    _logger.LogInformation("========================================================");

                                    paperLessConfig = JsonSerializer.Deserialize<WbDef.PaperLessConfig>(
                                        jsonText, PaperLessJsonOptions);

                                    if (paperLessConfig != null)
                                    {
                                        // ═════════════════════════════════════════════════════════
                                        // PAPERLESS DEBUG STAGE 9 — Deserialized PaperLessConfig
                                        // ═════════════════════════════════════════════════════════
                                        _logger.LogInformation("========================================================");
                                        _logger.LogInformation("PAPERLESS DEBUG STAGE 9 — Deserialized PaperLessConfig");
                                        _logger.LogInformation("========================================================");
                                        if (paperLessConfig.Sheets != null)
                                        {
                                            foreach (var ps9 in paperLessConfig.Sheets)
                                            {
                                                _logger.LogInformation("  Sheet: {Name}", ps9.Name);
                                                if (ps9.Fields != null)
                                                {
                                                    foreach (var pf9 in ps9.Fields.Take(3))
                                                    {
                                                        _logger.LogInformation("    Field ID: {Id}", pf9.Id);
                                                        _logger.LogInformation("    Cell: {Cell}", pf9.Cell);
                                                        _logger.LogInformation("    FontSize: {Sz}", pf9.Style?.Font?.SizePt ?? 0);
                                                        _logger.LogInformation("    FontName: {Fn}", pf9.Style?.Font?.Name ?? "(null)");
                                                        _logger.LogInformation("    Bold: {B}", pf9.Style?.Font?.Bold ?? false);
                                                    }
                                                }
                                            }
                                        }
                                        _logger.LogInformation("========================================================");

                                        log.AppendLine($"[PAPERLESS CONFIG] JSON parsed successfully");
                                        log.AppendLine($"[PAPERLESS CONFIG] Schema version: {paperLessConfig.SchemaVersion}");
                                        log.AppendLine($"[PAPERLESS CONFIG] Sheets in config: {paperLessConfig.Sheets?.Count ?? 0}");

                                        int configFieldCount = paperLessConfig.Sheets?.Sum(s => s.Fields?.Count ?? 0) ?? 0;
                                        log.AppendLine($"[PAPERLESS CONFIG] Fields in config: {configFieldCount}");

                                        if (paperLessConfig.SchemaVersion != 1)
                                        {
                                            log.AppendLine($"[PAPERLESS CONFIG] WARNING: Unsupported schema version {paperLessConfig.SchemaVersion} — ignoring config");
                                            _logger.LogWarning("[PAPERLESS CONFIG] Unsupported schema version {Version} — ignoring config", paperLessConfig.SchemaVersion);
                                            paperLessConfig = null;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.AppendLine($"[PAPERLESS CONFIG] WARNING: Failed to parse configuration JSON: {ex.Message}");
                        _logger.LogWarning(ex, "[PAPERLESS CONFIG] Configuration JSON invalid — falling back to Excel field detection");
                        paperLessConfig = null;
                    }
                }
                else
                {
                    log.AppendLine();
                    log.AppendLine("[PAPERLESS CONFIG] No PaperLessConfig sheet found — legacy workbook");
                    _logger.LogInformation("[PAPERLESS CONFIG] No configuration sheet found — legacy workbook");
                }

                // ═════════════════════════════════════════════════════════
                // 3-9. BUILD PAGES from printable sheets
                // ═════════════════════════════════════════════════════════
                int pageIndex = 0;
                foreach (var sheet in printableSheets)
                {
                    string sheetName = sheet.Name?.Value ?? "?";
                    string pageId = $"page_{pageIndex}";

                    log.AppendLine();
                    log.AppendLine("========================================");
                    log.AppendLine($"[DesignerModelReader] Processing Sheet: {sheetName}");
                    log.AppendLine("========================================");

                    var page = new DesignerPage
                    {
                        Id = pageId,
                        Name = sheetName,
                        Index = pageIndex
                    };

                    if (sheet.Id?.Value == null) continue;
                    var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    if (wsPart == null) continue;

                    var ws = wsPart.Worksheet;

                    // ═══════════════════════════════════════════════════
                    // 3. PAGE LAYOUT
                    // ═══════════════════════════════════════════════════
                    ReadPageLayout(ws, wbPart, sheet, page, log);

                    // ═══════════════════════════════════════════════════
                    // 4. ROWS
                    // ═══════════════════════════════════════════════════
                    ReadRows(ws, page, log);

                    // ═══════════════════════════════════════════════════
                    // 5. COLUMNS
                    // ═══════════════════════════════════════════════════
                    ReadColumns(ws, page, log);

                    // ═══════════════════════════════════════════════════
                    // 6. MERGED CELLS
                    // ═══════════════════════════════════════════════════
                    ReadMergedCells(ws, page, log);

                    // ═══════════════════════════════════════════════════
                    // 9. COMMENTS
                    // ═══════════════════════════════════════════════════
                    var sheetComments = ReadComments(wsPart, sheetName, log);

                    // ═══════════════════════════════════════════════════
                    // FIELD CREATION from _Fields data
                    // ═══════════════════════════════════════════════════
                    var sheetFieldsData = fieldsData
                        .Where(f => string.Equals(f.sheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    // Also match fields where sheetName is empty (legacy data)
                    if (sheetFieldsData.Count == 0)
                    {
                        sheetFieldsData = fieldsData
                            .Where(f => string.IsNullOrEmpty(f.sheetName))
                            .ToList();
                    }

                    // Build style lookup for this sheet
                    var cellStyles = BuildCellStyleLookup(wsPart, wbPart);
                    var allCellFormats = ReadAllCellFormats(wbPart);

                    foreach (var fd in sheetFieldsData)
                    {
                        string cellAddr = fd.cellAddr.Split(':')[0];
                        var field = new DesignerField
                        {
                            Id = fd.fieldId,
                            Name = fd.name,
                            CellAddress = cellAddr,
                            Type = MapFieldType(fd.type),
                            Behavior = new DesignerFieldBehavior
                            {
                                Required = fd.required == "1" || string.Equals(fd.required, "true", StringComparison.OrdinalIgnoreCase),
                                ReadOnly = fd.readOnly == "1" || string.Equals(fd.readOnly, "true", StringComparison.OrdinalIgnoreCase)
                            },
                            MaxLength = int.TryParse(fd.maxLength, out int ml) ? ml : 0,
                            Placeholder = fd.placeholder,
                            DefaultValue = fd.defaultValue,
                            Label = fd.name,
                            Description = GetCommentForCell(sheetComments, cellAddr)
                        };

                        if (!string.IsNullOrEmpty(fd.options))
                        {
                            field.Options = fd.options
                                .Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(o => o.Trim())
                                .Where(o => !string.IsNullOrEmpty(o))
                                .ToList();
                        }

                        if (!string.IsNullOrEmpty(fd.validation))
                        {
                            field.Validation = new DesignerFieldValidation
                            {
                                Type = "custom",
                                Formula1 = fd.validation
                            };
                        }

                        // ═══════════════════════════════════════════════
                        // 8. CELL POSITION (bounds)
                        // ═══════════════════════════════════════════════
                        ReadCellBounds(ws, wbPart, sheet, cellAddr, field, log);

                        // ═══════════════════════════════════════════════
                        // 7. CELL FORMATTING
                        // ═══════════════════════════════════════════════
                        ReadCellFormatting(cellAddr, cellStyles, allCellFormats, field, log);

                        // ═══════════════════════════════════════════════
                        // 12. DATA VALIDATION
                        // ═══════════════════════════════════════════════
                        ReadDataValidation(ws, cellAddr, field, log);

                        // Log field summary
                        log.AppendLine();
                        log.AppendLine("----------------------------------------");
                        log.AppendLine($"[DesignerModelReader] Field: {fd.fieldId}");
                        log.AppendLine($"  Cell: {cellAddr}");
                        log.AppendLine($"  Type: {field.Type}");
                        log.AppendLine($"  Required: {field.Behavior.Required}");
                        log.AppendLine($"  ReadOnly: {field.Behavior.ReadOnly}");
                        log.AppendLine("----------------------------------------");

                        page.Fields.Add(field);
                    }

                    // Also create fields from comments (fallback if no _Fields data)
                    if (sheetFieldsData.Count == 0 && sheetComments.Count > 0)
                    {
                        log.AppendLine();
                        log.AppendLine("[DesignerModelReader] No _Fields data — creating fields from comments (fallback)");
                        int fi = 0;
                        foreach (var (cellAddr, text, _) in sheetComments)
                        {
                            string fieldId = $"field_{pageIndex}_{fi}";
                            string fieldName = text.Split('\n')[0].Trim();
                            if (string.IsNullOrWhiteSpace(fieldName)) fieldName = fieldId;

                            var field = new DesignerField
                            {
                                Id = fieldId,
                                Name = fieldName,
                                CellAddress = cellAddr,
                                Type = "text",
                                Description = text
                            };

                            ReadCellBounds(ws, wbPart, sheet, cellAddr, field, log);
                            ReadCellFormatting(cellAddr, cellStyles, allCellFormats, field, log);
                            ReadDataValidation(ws, cellAddr, field, log);

                            page.Fields.Add(field);
                            fi++;
                        }
                    }

                    // ═══════════════════════════════════════════════════
                    // 13. IMAGES
                    // ═══════════════════════════════════════════════════
                    ReadImages(wsPart, sheetName, log);

                    // ═══════════════════════════════════════════════════
                    // 14. SHAPES
                    // ═══════════════════════════════════════════════════
                    ReadShapes(wsPart, sheetName, log);

                    model.Pages.Add(page);
                    pageIndex++;
                }

                // ═════════════════════════════════════════════════════════
                // 12b. OVERLAY PAPERLESSCONFIG — restore PaperLess field identity,
                //      style, and configuration on top of Excel-detected fields.
                // ═════════════════════════════════════════════════════════

                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 10 — Excel Field Detection Before Overlay
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 10 — Field Detection Before Overlay");
                _logger.LogInformation("========================================================");
                foreach (var p10 in model.Pages)
                {
                    _logger.LogInformation("  Page: {Name}", p10.Name);
                    foreach (var f10 in p10.Fields)
                    {
                        _logger.LogInformation("    Field ID: {Id}", f10.Id);
                        _logger.LogInformation("    Cell: {Cell}", f10.CellAddress ?? "?");
                        _logger.LogInformation("    FontSize: {Sz}", f10.Style?.FontSize ?? 0);
                        _logger.LogInformation("    FontFamily: {Ff}", f10.Style?.FontFamily ?? "(default)");
                        _logger.LogInformation("    Bold: {B}", f10.Style?.Bold ?? false);
                    }
                }
                _logger.LogInformation("========================================================");

                if (paperLessConfig?.Sheets != null)
                {
                    int overlaidFields = 0;
                    foreach (var configSheet in paperLessConfig.Sheets)
                    {
                        if (configSheet.Fields == null) continue;

                        // Find the matching page by sheet name
                        var page = model.Pages.FirstOrDefault(p =>
                            string.Equals(p.Name, configSheet.Name, StringComparison.OrdinalIgnoreCase));
                        if (page == null) continue;

                        foreach (var configField in configSheet.Fields)
                        {
                            // Match by cell address — extract first cell from range
                            string configCell = NormalizeCellAddress(configField.Cell);
                            if (string.IsNullOrWhiteSpace(configCell)) continue;

                            // ═════════════════════════════════════════════════════════
                            // PAPERLESS DEBUG STAGE 11 — Field Matching
                            // ═════════════════════════════════════════════════════════
                            _logger.LogInformation("--------------------------------------------------------");
                            _logger.LogInformation("PAPERLESS DEBUG STAGE 11 — Field Matching");
                            _logger.LogInformation("  Config Field:");
                            _logger.LogInformation("    ID = {Id}", configField.Id);
                            _logger.LogInformation("    Cell = {Cell}", configField.Cell);
                            _logger.LogInformation("    Normalized = {NCell}", configCell);
                            _logger.LogInformation("  Searching in page '{Page}' with {FieldCount} fields",
                                page.Name, page.Fields.Count);

                            var matchingField = page.Fields.FirstOrDefault(f =>
                                string.Equals(NormalizeCellAddress(f.CellAddress ?? ""), configCell, StringComparison.OrdinalIgnoreCase));

                            if (matchingField == null)
                            {
                                _logger.LogInformation("  MATCH = FALSE — no field found with normalized cell '{NCell}'", configCell);
                                _logger.LogInformation("--------------------------------------------------------");
                                continue;
                            }

                            _logger.LogInformation("  Detected Field:");
                            _logger.LogInformation("    ID = {Id}", matchingField.Id);
                            _logger.LogInformation("    Cell = {Cell}", matchingField.CellAddress ?? "?");
                            _logger.LogInformation("  MATCH = TRUE");
                            _logger.LogInformation("--------------------------------------------------------");

                            // ═════════════════════════════════════════════════════════
                            // PAPERLESS DEBUG STAGE 12 — Before Overlay
                            // ═════════════════════════════════════════════════════════
                            _logger.LogInformation("PAPERLESS DEBUG STAGE 12 — Before Overlay");
                            _logger.LogInformation("  Detected field state BEFORE overlay:");
                            _logger.LogInformation("    Field ID: {Id}", matchingField.Id);
                            _logger.LogInformation("    Cell: {Cell}", matchingField.CellAddress ?? "?");
                            _logger.LogInformation("    FontSize: {Sz}", matchingField.Style?.FontSize ?? 0);
                            _logger.LogInformation("    FontFamily: {Ff}", matchingField.Style?.FontFamily ?? "(default)");
                            _logger.LogInformation("    Bold: {B}", matchingField.Style?.Bold ?? false);
                            _logger.LogInformation("  Config value to apply:");
                            _logger.LogInformation("    Config FontSize: {Sz}", configField.Style?.Font?.SizePt ?? 0);
                            _logger.LogInformation("--------------------------------------------------------");

                            // Overlay PaperLess field ID (stable identity)
                            if (!string.IsNullOrWhiteSpace(configField.Id))
                            {
                                matchingField.Id = configField.Id;
                            }

                            // Restore PaperLess style properties from config (takes precedence)
                            if (configField.Style != null)
                            {
                                matchingField.Style ??= new DesignerFieldStyle();

                                if (configField.Style.Font != null)
                                {
                                    var cf = configField.Style.Font;
                                    if (!string.IsNullOrEmpty(cf.Name))
                                        matchingField.Style.FontFamily = cf.Name;
                                    if (cf.SizePt > 0)
                                        matchingField.Style.FontSize = cf.SizePt;
                                    matchingField.Style.Bold = cf.Bold;
                                    matchingField.Style.Italic = cf.Italic;
                                    matchingField.Style.Underline = cf.Underline;
                                    if (!string.IsNullOrEmpty(cf.ColorArgb))
                                        matchingField.Style.FontColor = cf.ColorArgb;
                                }

                                if (configField.Style.Fill != null)
                                {
                                    matchingField.Style.FillColor = configField.Style.Fill.ColorArgb;
                                }

                                if (configField.Style.Alignment != null)
                                {
                                    if (!string.IsNullOrEmpty(configField.Style.Alignment.Horizontal))
                                        matchingField.Style.HorizontalAlignment = MapHorizontalAlignment(configField.Style.Alignment.Horizontal);
                                    if (!string.IsNullOrEmpty(configField.Style.Alignment.Vertical))
                                        matchingField.Style.VerticalAlignment = MapVerticalAlignment(configField.Style.Alignment.Vertical);
                                }

                                matchingField.Style.WrapText = configField.Style.WrapText;
                            }

                            // Restore field type from config
                            if (!string.IsNullOrWhiteSpace(configField.Type))
                            {
                                string mappedType = MapFieldType(configField.Type);
                                if (!string.IsNullOrEmpty(mappedType))
                                    matchingField.Type = mappedType;
                            }

                            // Restore input configuration
                            if (configField.Config != null)
                            {
                                matchingField.Behavior.Required = configField.Config.Required;
                                if (configField.Config.MaxLength > 0)
                                    matchingField.MaxLength = configField.Config.MaxLength;
                            }

                            // ═════════════════════════════════════════════════════════
                            // PAPERLESS DEBUG STAGE 13 — After Overlay
                            // ═════════════════════════════════════════════════════════
                            _logger.LogInformation("PAPERLESS DEBUG STAGE 13 — After Overlay");
                            _logger.LogInformation("  Field state AFTER overlay:");
                            _logger.LogInformation("    Field ID: {Id}", matchingField.Id);
                            _logger.LogInformation("    Cell: {Cell}", matchingField.CellAddress ?? "?");
                            _logger.LogInformation("    FontSize: {Sz}", matchingField.Style?.FontSize ?? 0);
                            _logger.LogInformation("    FontFamily: {Ff}", matchingField.Style?.FontFamily ?? "(default)");
                            _logger.LogInformation("    Bold: {B}", matchingField.Style?.Bold ?? false);
                            _logger.LogInformation("--------------------------------------------------------");

                            log.AppendLine($"[PAPERLESS CONFIG] Field '{configField.Id}' matched by cell {configField.Cell}");
                            if (configField.Style?.Font != null)
                                log.AppendLine($"[PAPERLESS CONFIG]   Restored style: fontSize={configField.Style.Font.SizePt}, bold={configField.Style.Font.Bold}, font='{configField.Style.Font.Name}'");
                            overlaidFields++;
                        }
                    }

                    if (overlaidFields > 0)
                    {
                        log.AppendLine();
                        log.AppendLine($"[PAPERLESS CONFIG] Fields restored from configuration: {overlaidFields}");
                        _logger.LogInformation("[PAPERLESS CONFIG] Fields restored from configuration: {Count}", overlaidFields);
                    }
                }

                // ═════════════════════════════════════════════════════════
                // Add comments to model
                // ═════════════════════════════════════════════════════════
                foreach (var sheet in allSheets)
                {
                    if (sheet.Id?.Value == null) continue;
                    var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    if (wsPart == null) continue;

                    var comments = wsPart.WorksheetCommentsPart?.Comments?
                        .Descendants<Comment>().ToList() ?? new();

                    foreach (var c in comments)
                    {
                        string cellRef = c.Reference?.Value ?? "";
                        string text = GetCommentText(c);
                        string author = "";

                        model.Comments.Add(new DesignerComment
                        {
                            CellAddress = cellRef,
                            Worksheet = sheet.Name?.Value ?? "",
                            Text = text,
                            Author = author
                        });
                    }
                }

                log.AppendLine();
                log.AppendLine("========================================================");
                log.AppendLine("[DesignerModelReader] DesignerModel Created Successfully");
                log.AppendLine($"  Pages: {model.Pages.Count}");
                log.AppendLine($"  Fields: {model.Pages.Sum(p => p.Fields.Count)}");
                log.AppendLine($"  Comments: {model.Comments.Count}");
                log.AppendLine($"  _Fields Sheet: {(model.Configuration.HasFieldsSheet ? "YES" : "NO")}");
                log.AppendLine($"  ExcelOutputSetting: {(model.Configuration.HasExcelOutputSetting ? "YES" : "NO")}");
                log.AppendLine("========================================================");

                // ═══════════════════════════════════════════════════════════════
                // PHASE 21.2: ROUND TRIP STATE — per-field PASS/FAIL comparison
                // ═══════════════════════════════════════════════════════════════
                log.AppendLine();
                log.AppendLine("========================================================");
                log.AppendLine("  ROUND TRIP STATE — Workbook vs DesignerModel");
                log.AppendLine("========================================================");

                // Re-read the workbook to get raw cell styles for comparison
                var wbStyles = BuildAllCellStylesFromWorkbook(wbPart);
                int passCount = 0;
                int failCount = 0;

                foreach (var page in model.Pages)
                {
                    foreach (var field in page.Fields)
                    {
                        string cellAddr = field.CellAddress ?? "";
                        if (string.IsNullOrEmpty(cellAddr)) continue;

                        log.AppendLine();
                        log.AppendLine($"  Field: {field.Id}");
                        log.AppendLine($"  ------------------------------");

                        // Get workbook raw style for this cell
                        wbStyles.TryGetValue(cellAddr.ToUpperInvariant(), out var wbStyle);

                        // Compare Font Family
                        string wbFont = wbStyle?.FontName ?? "(not set)";
                        string dmFont = field.Style?.FontFamily ?? "(not set)";
                        bool fontPass = string.Equals(wbFont, dmFont, StringComparison.OrdinalIgnoreCase)
                            || (string.IsNullOrEmpty(wbFont) && string.IsNullOrEmpty(dmFont));
                        log.AppendLine($"    Font Name: Workbook='{wbFont}' DesignerModel='{dmFont}' {(fontPass ? "PASS" : "FAIL")}");
                        if (fontPass) passCount++; else failCount++;

                        // Compare Font Size
                        double wbSize = wbStyle?.FontSize ?? 0;
                        double? dmSize = field.Style?.FontSize;
                        bool sizePass = Math.Abs(wbSize - (dmSize ?? 0)) < 0.1
                            || (wbSize == 0 && (dmSize == null || dmSize == 0));
                        log.AppendLine($"    Font Size:  Workbook='{wbSize}' DesignerModel='{dmSize}' {(sizePass ? "PASS" : "FAIL")}");
                        if (sizePass) passCount++; else failCount++;

                        // Compare Bold
                        bool wbBold = wbStyle?.Bold ?? false;
                        bool? dmBold = field.Style?.Bold;
                        bool boldPass = wbBold == (dmBold ?? false);
                        log.AppendLine($"    Bold:       Workbook='{wbBold}' DesignerModel='{dmBold}' {(boldPass ? "PASS" : "FAIL")}");
                        if (boldPass) passCount++; else failCount++;

                        // Compare Italic
                        bool wbItalic = wbStyle?.Italic ?? false;
                        bool? dmItalic = field.Style?.Italic;
                        bool italicPass = wbItalic == (dmItalic ?? false);
                        log.AppendLine($"    Italic:     Workbook='{wbItalic}' DesignerModel='{dmItalic}' {(italicPass ? "PASS" : "FAIL")}");
                        if (italicPass) passCount++; else failCount++;

                        // Compare Font Color
                        string wbColor = string.IsNullOrEmpty(wbStyle?.FontColor) ? "(default)" : wbStyle.FontColor;
                        string dmColor = string.IsNullOrEmpty(field.Style?.FontColor) ? "(default)" : field.Style.FontColor;
                        bool colorPass = string.Equals(wbColor, dmColor, StringComparison.OrdinalIgnoreCase);
                        log.AppendLine($"    Font Color: Workbook='{wbColor}' DesignerModel='{dmColor}' {(colorPass ? "PASS" : "FAIL")}");
                        if (colorPass) passCount++; else failCount++;

                        // Compare Fill Color
                        string wbFill = string.IsNullOrEmpty(wbStyle?.FillColor) ? "(none)" : wbStyle.FillColor;
                        string dmFill = string.IsNullOrEmpty(field.Style?.FillColor) ? "(none)" : field.Style.FillColor;
                        bool fillPass = string.Equals(wbFill, dmFill, StringComparison.OrdinalIgnoreCase);
                        log.AppendLine($"    Fill Color: Workbook='{wbFill}' DesignerModel='{dmFill}' {(fillPass ? "PASS" : "FAIL")}");
                        if (fillPass) passCount++; else failCount++;

                        // Compare Horizontal Alignment
                        string wbHAlign = string.IsNullOrEmpty(wbStyle?.HorizontalAlignment) ? "(default)" : wbStyle.HorizontalAlignment;
                        string dmHAlign = string.IsNullOrEmpty(field.Style?.HorizontalAlignment) ? "(default)" : field.Style.HorizontalAlignment;
                        bool hAlignPass = string.Equals(wbHAlign, dmHAlign, StringComparison.OrdinalIgnoreCase);
                        log.AppendLine($"    H-Align:    Workbook='{wbHAlign}' DesignerModel='{dmHAlign}' {(hAlignPass ? "PASS" : "FAIL")}");
                        if (hAlignPass) passCount++; else failCount++;

                        // Compare Vertical Alignment
                        string wbVAlign = string.IsNullOrEmpty(wbStyle?.VerticalAlignment) ? "(default)" : wbStyle.VerticalAlignment;
                        string dmVAlign = string.IsNullOrEmpty(field.Style?.VerticalAlignment) ? "(default)" : field.Style.VerticalAlignment;
                        bool vAlignPass = string.Equals(wbVAlign, dmVAlign, StringComparison.OrdinalIgnoreCase);
                        log.AppendLine($"    V-Align:    Workbook='{wbVAlign}' DesignerModel='{dmVAlign}' {(vAlignPass ? "PASS" : "FAIL")}");
                        if (vAlignPass) passCount++; else failCount++;

                        // Compare Wrap Text
                        bool wbWrap = wbStyle?.WrapText ?? false;
                        bool? dmWrap = field.Style?.WrapText;
                        bool wrapPass = wbWrap == (dmWrap ?? false);
                        log.AppendLine($"    Wrap Text:  Workbook='{wbWrap}' DesignerModel='{dmWrap}' {(wrapPass ? "PASS" : "FAIL")}");
                        if (wrapPass) passCount++; else failCount++;
                    }
                }

                log.AppendLine();
                log.AppendLine("========================================================");
                log.AppendLine($"  ROUND TRIP RESULT: {passCount} PASS / {failCount} FAIL / {(passCount + failCount)} TOTAL");
                if (failCount == 0)
                    log.AppendLine("  ALL PROPERTIES MATCH — Pipeline is correct");
                else
                {
                    log.AppendLine($"  {failCount} PROPERTIES DO NOT MATCH — Investigate above");
                    _logger.LogWarning("[DesignerModelReader] Round-trip check: {FailCount} properties mismatched", failCount);
                }
                log.AppendLine("========================================================");

                // Log the complete output
                _logger.LogInformation("{Log}", log.ToString());

                return model;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[DesignerModelReader] Failed to read workbook: {Message}", ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Phase 21.2: Build a comprehensive per-cell style dictionary for round-trip comparison.
        /// Reads ALL cells from ALL worksheets and returns their raw OpenXML style.
        /// </summary>
        private Dictionary<string, OpenXmlCellStyleModel> BuildAllCellStylesFromWorkbook(WorkbookPart wbPart)
        {
            var allStyles = new Dictionary<string, OpenXmlCellStyleModel>(StringComparer.OrdinalIgnoreCase);

            var stylesPart = wbPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet == null) return allStyles;

            var fonts = stylesPart.Stylesheet.Fonts?.Cast<Font>().ToList() ?? new();
            var fills = stylesPart.Stylesheet.Fills?.Cast<Fill>().ToList() ?? new();
            var borders = stylesPart.Stylesheet.Borders?.Cast<Border>().ToList() ?? new();
            var cellFormats = stylesPart.Stylesheet.CellFormats?.Cast<CellFormat>().ToList() ?? new();

            var sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
            foreach (var sheet in sheets)
            {
                if (sheet.Id?.Value == null) continue;
                var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (wsPart == null) continue;

                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                foreach (var cell in cells)
                {
                    string addr = cell.CellReference?.Value ?? "";
                    if (string.IsNullOrEmpty(addr)) continue;

                    uint styleIdx = cell.StyleIndex?.Value ?? 0u;
                    if (styleIdx >= cellFormats.Count) continue;

                    var xf = cellFormats[(int)styleIdx];
                    var style = new OpenXmlCellStyleModel();

                    // Font
                    if (xf.FontId?.Value != null)
                    {
                        int fontId = (int)xf.FontId.Value;
                        if (fontId < fonts.Count)
                        {
                            var font = fonts[fontId];
                            style.FontName = font.FontName?.Val?.Value ?? "";
                            if (font.FontSize?.Val?.Value != null)
                                style.FontSize = (double)font.FontSize.Val.Value;
                            style.Bold = font.Bold?.Val?.Value ?? false;
                            style.Italic = font.Italic?.Val?.Value ?? false;
                            style.Underline = font.Underline?.Val?.Value != null;
                            var fontColor = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                            style.FontColor = GetColorHex(fontColor);
                        }
                    }

                    // Fill
                    if (xf.FillId?.Value != null)
                    {
                        int fillId = (int)xf.FillId.Value;
                        if (fillId < fills.Count && fillId >= 0)
                        {
                            var fill = fills[fillId];
                            var patternFill = fill.PatternFill;
                            if (patternFill?.PatternType?.Value != null &&
                                patternFill.PatternType.Value != PatternValues.None &&
                                patternFill.PatternType.Value != PatternValues.Gray125)
                            {
                                style.FillColor = GetFillColorHex(patternFill.ForegroundColor);
                                if (string.IsNullOrEmpty(style.FillColor))
                                    style.FillColor = GetFillColorHex(patternFill.BackgroundColor);
                            }
                        }
                    }

                    // Alignment
                    if (xf.Alignment != null)
                    {
                        var align = xf.Alignment;
                        if (align.Horizontal?.Value != null)
                            style.HorizontalAlignment = MapHoriz(align.Horizontal.Value.ToString());
                        if (align.Vertical?.Value != null)
                            style.VerticalAlignment = MapVert(align.Vertical.Value.ToString());
                        style.WrapText = align.WrapText?.Value ?? false;
                    }

                    // Use uppercase address for matching
                    allStyles[addr.ToUpperInvariant()] = style;
                }
            }

            return allStyles;
        }

        // ═══════════════════════════════════════════════════════════════
        // 1. WORKBOOK INFORMATION
        // ═══════════════════════════════════════════════════════════════

        private void ReadWorkbookInfo(SpreadsheetDocument doc, WorkbookPart wbPart, DesignerModel model, StringBuilder log)
        {
            log.AppendLine();
            log.AppendLine("[DesignerModelReader] Workbook Info");
            log.AppendLine("---------------------------------------");

            // Read from document properties
            string title = "", author = "", description = "", company = "";
            DateTime created = default, modified = default;

            var props = doc.PackageProperties;
            if (props != null)
            {
                title = props.Title ?? "";
                author = props.Creator ?? "";
                description = props.Description ?? "";
                // IPackageProperties does not have Company in all SDK versions
                // company = props.Company ?? "";
                if (props.Created != null) created = props.Created.Value;
                if (props.Modified != null) modified = props.Modified.Value;
            }

            model.Info = new DesignerWorkbookInfo
            {
                Title = title,
                Author = author,
                Description = description,
                Created = created,
                Modified = modified,
                Version = "1.0"
            };

            log.AppendLine($"  Title: {title}");
            log.AppendLine($"  Author: {author}");
            log.AppendLine($"  Company: {(company)}");
            log.AppendLine($"  Created: {created}");
            log.AppendLine($"  Modified: {modified}");

            _logger.LogInformation("[DesignerModelReader] Workbook Info: Title={Title}, Author={Author}",
                title, author);
        }

        // ═══════════════════════════════════════════════════════════════
        // 3. PAGE LAYOUT
        // ═══════════════════════════════════════════════════════════════

        private void ReadPageLayout(Worksheet ws, WorkbookPart wbPart, Sheet sheet,
            DesignerPage page, StringBuilder log)
        {
            var layout = page.Layout;

            // Print area from workbook-level defined names
            var printAreaName = wbPart.Workbook.Descendants<DefinedName>()
                .FirstOrDefault(n => n.Name?.Value != null &&
                    n.Name.Value.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (n.Text?.Contains(page.Name) == true || n.Text?.Contains($"'{page.Name}'") == true));
            if (printAreaName != null)
                layout.PrintArea = printAreaName.Text ?? "";

            // Page margins
            var margins = ws.Descendants<PageMargins>().FirstOrDefault();
            if (margins != null)
            {
                layout.Margins.Top = (double)(margins.Top?.Value ?? 0.0) * 72;
                layout.Margins.Bottom = (double)(margins.Bottom?.Value ?? 0.0) * 72;
                layout.Margins.Left = (double)(margins.Left?.Value ?? 0.0) * 72;
                layout.Margins.Right = (double)(margins.Right?.Value ?? 0.0) * 72;
                layout.Margins.Header = (double)(margins.Header?.Value ?? 0.0) * 72;
                layout.Margins.Footer = (double)(margins.Footer?.Value ?? 0.0) * 72;
            }

            // Page setup
            var pageSetup = ws.Descendants<PageSetup>().FirstOrDefault();
            if (pageSetup != null)
            {
                var orientVal = pageSetup.Orientation?.Value;
                layout.Orientation = orientVal != null
                    ? orientVal.ToString().ToLower()
                    : "portrait";
                layout.Zoom = 100; // OpenXML PageSetup does not have Scale/Zoom — use default
                layout.FitToPagesWide = (int)(pageSetup.FitToWidth?.Value ?? 0u);
                layout.FitToPagesTall = (int)(pageSetup.FitToHeight?.Value ?? 0u);

                // Map paper size to name
                var paperSize = (int)(pageSetup.PaperSize?.Value ?? 0u);
                layout.PaperSize = paperSize switch
                {
                    1 => "Letter",
                    5 => "Legal",
                    9 => "A4",
                    13 => "B5",
                    8 => "A3",
                    11 => "A5",
                    _ => $"Custom({paperSize})"
                };
            }

            // Print options (centering)
            var printOptions = ws.Descendants<PrintOptions>().FirstOrDefault();
            if (printOptions != null)
            {
                layout.CenterHorizontally = printOptions.HorizontalCentered?.Value ?? false;
                layout.CenterVertically = printOptions.VerticalCentered?.Value ?? false;
            }

            // Sheet dimension
            var dimension = ws.Descendants<SheetDimension>().FirstOrDefault();
            if (dimension?.Reference?.Value != null)
            {
                log.AppendLine($"[DesignerModelReader]   Page Layout — Dimension: {dimension.Reference.Value}");
            }

            log.AppendLine($"[DesignerModelReader]   Print Area: {layout.PrintArea}");
            log.AppendLine($"[DesignerModelReader]   Orientation: {layout.Orientation}");
            log.AppendLine($"[DesignerModelReader]   Paper: {layout.PaperSize}");
            log.AppendLine($"[DesignerModelReader]   Zoom: {layout.Zoom}");
            log.AppendLine($"[DesignerModelReader]   FitToPage: {layout.FitToPagesWide}x{layout.FitToPagesTall}");
            log.AppendLine($"[DesignerModelReader]   Margins: L={layout.Margins.Left:F2} R={layout.Margins.Right:F2} T={layout.Margins.Top:F2} B={layout.Margins.Bottom:F2}");

            _logger.LogInformation("[DesignerModelReader] Layout for '{Name}': PA={PA}, O={O}, Paper={Paper}, Zoom={Z}",
                page.Name, layout.PrintArea, layout.Orientation, layout.PaperSize, layout.Zoom);
        }

        // ═══════════════════════════════════════════════════════════════
        // 4. ROWS
        // ═══════════════════════════════════════════════════════════════

        private void ReadRows(Worksheet ws, DesignerPage page, StringBuilder log)
        {
            var rows = ws.Descendants<Row>().ToList();
            int hiddenCount = 0;

            foreach (var row in rows)
            {
                var ri = new DesignerRowInfo
                {
                    Index = (int)(row.RowIndex?.Value ?? 0u),
                    HeightPt = (double)(row.Height?.Value ?? 0.0),
                    Hidden = row.Hidden?.Value ?? false
                };
                if (ri.Hidden) hiddenCount++;
                page.Layout.Rows.Add(ri);
            }

            log.AppendLine($"[DesignerModelReader]   Rows: {rows.Count} (hidden: {hiddenCount})");
        }

        // ═══════════════════════════════════════════════════════════════
        // 5. COLUMNS
        // ═══════════════════════════════════════════════════════════════

        private void ReadColumns(Worksheet ws, DesignerPage page, StringBuilder log)
        {
            var cols = ws.Descendants<Column>().ToList();
            int hiddenCount = 0;

            foreach (var col in cols)
            {
                uint min = col.Min?.Value ?? 1u;
                uint max = col.Max?.Value ?? min;
                double width = (double)(col.Width?.Value ?? 0.0);
                bool hidden = col.Hidden?.Value ?? false;

                for (uint i = min; i <= max; i++)
                {
                    page.Layout.Columns.Add(new DesignerColumnInfo
                    {
                        Index = (int)i,
                        WidthChars = width,
                        Hidden = hidden
                    });
                }
                if (hidden) hiddenCount++;
            }

            log.AppendLine($"[DesignerModelReader]   Columns: {page.Layout.Columns.Count} (hidden: {hiddenCount})");
        }

        // ═══════════════════════════════════════════════════════════════
        // 6. MERGED CELLS
        // ═══════════════════════════════════════════════════════════════

        private void ReadMergedCells(Worksheet ws, DesignerPage page, StringBuilder log)
        {
            var merges = ws.Descendants<MergeCell>().ToList();
            foreach (var m in merges)
            {
                if (m.Reference?.Value != null)
                    page.Layout.MergedCells.Add(m.Reference.Value);
            }

            log.AppendLine($"[DesignerModelReader]   Merged Cells: {merges.Count}");
        }

        // ═══════════════════════════════════════════════════════════════
        // 7. CELL FORMATTING
        // ═══════════════════════════════════════════════════════════════

        private Dictionary<string, OpenXmlCellStyleModel> BuildCellStyleLookup(
            WorksheetPart wsPart, WorkbookPart wbPart)
        {
            var result = new Dictionary<string, OpenXmlCellStyleModel>();

            var stylesPart = wbPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet == null) return result;

            var fonts = stylesPart.Stylesheet.Fonts?.Cast<Font>().ToList() ?? new();
            var fills = stylesPart.Stylesheet.Fills?.Cast<Fill>().ToList() ?? new();
            var borders = stylesPart.Stylesheet.Borders?.Cast<Border>().ToList() ?? new();
            var cellFormats = stylesPart.Stylesheet.CellFormats?.Cast<CellFormat>().ToList() ?? new();

            var sheetData = wsPart.Worksheet.Descendants<Cell>().ToList();

            foreach (var cell in sheetData)
            {
                string addr = cell.CellReference?.Value ?? "";
                if (string.IsNullOrEmpty(addr)) continue;

                uint styleIdx = cell.StyleIndex?.Value ?? 0u;
                if (styleIdx >= cellFormats.Count) continue;

                var xf = cellFormats[(int)styleIdx];
                var style = new OpenXmlCellStyleModel();

                // Font
                if (xf.FontId?.Value != null)
                {
                    int fontId = (int)xf.FontId.Value;
                    if (fontId >= 0 && fontId < fonts.Count)
                    {
                        var font = fonts[fontId];
                        style.FontName = font.FontName?.Val?.Value ?? "";
                        if (font.FontSize?.Val?.Value != null)
                            style.FontSize = (double)font.FontSize.Val.Value;
                        style.Bold = font.Bold?.Val?.Value ?? false;
                        style.Italic = font.Italic?.Val?.Value ?? false;
                        style.Underline = font.Underline?.Val?.Value != null;

                        var fontColor = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                        style.FontColor = GetColorHex(fontColor);
                    }
                }

                // Fill
                if (xf.FillId?.Value != null)
                {
                    int fillId = (int)xf.FillId.Value;
                    if (fillId >= 0 && fillId < fills.Count)
                    {
                        var fill = fills[fillId];
                        var patternFill = fill.PatternFill;
                        if (patternFill?.PatternType?.Value != null &&
                            patternFill.PatternType.Value != PatternValues.None &&
                            patternFill.PatternType.Value != PatternValues.Gray125)
                        {
                            style.FillColor = GetFillColorHex(patternFill.ForegroundColor);
                            if (string.IsNullOrEmpty(style.FillColor))
                                style.FillColor = GetFillColorHex(patternFill.BackgroundColor);
                        }
                    }
                }

                // Border
                if (xf.BorderId?.Value != null)
                {
                    int borderId = (int)xf.BorderId.Value;
                    if (borderId >= 0 && borderId < borders.Count)
                    {
                        var border = borders[borderId];
                        style.BorderTop = GetBorderCss(border.TopBorder);
                        style.BorderBottom = GetBorderCss(border.BottomBorder);
                        style.BorderLeft = GetBorderCss(border.LeftBorder);
                        style.BorderRight = GetBorderCss(border.RightBorder);
                    }
                }

                // Alignment
                if (xf.Alignment != null)
                {
                    var align = xf.Alignment;
                    if (align.Horizontal?.Value != null)
                        style.HorizontalAlignment = MapHoriz(align.Horizontal.Value.ToString());
                    if (align.Vertical?.Value != null)
                        style.VerticalAlignment = MapVert(align.Vertical.Value.ToString());
                    style.WrapText = align.WrapText?.Value ?? false;
                    style.Rotation = (int)(align.TextRotation?.Value ?? 0u);
                }

                // Number format
                if (xf.NumberFormatId?.Value != null)
                {
                    int nfId = (int)xf.NumberFormatId.Value;
                    if (nfId > 0)
                    {
                        // Try custom number format from NumberingFormats
                        var numFmts = stylesPart.Stylesheet.NumberingFormats;
                        if (numFmts != null)
                        {
                            var nf = numFmts.Cast<NumberingFormat>()
                                .FirstOrDefault(n => n.NumberFormatId?.Value == nfId);
                            if (nf?.FormatCode?.Value != null)
                                style.NumberFormat = nf.FormatCode.Value;
                        }
                        if (string.IsNullOrEmpty(style.NumberFormat))
                        {
                            style.NumberFormat = nfId switch
                            {
                                1 => "0",
                                2 => "0.00",
                                3 => "#,##0",
                                4 => "#,##0.00",
                                9 => "0%",
                                10 => "0.00%",
                                11 => "0.00E+00",
                                12 => "# ?/?",
                                13 => "# ??/??",
                                14 => "mm-dd-yy",
                                15 => "d-mmm-yy",
                                16 => "d-mmm",
                                17 => "mmm-yy",
                                18 => "h:mm AM/PM",
                                19 => "h:mm:ss AM/PM",
                                20 => "h:mm",
                                21 => "h:mm:ss",
                                22 => "m/d/yy h:mm",
                                _ => ""
                            };
                        }
                    }
                }

                result[addr] = style;
            }

            return result;
        }

        private List<OpenXmlCellStyleModel> ReadAllCellFormats(WorkbookPart wbPart)
        {
            var result = new List<OpenXmlCellStyleModel>();
            var stylesPart = wbPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet?.CellFormats == null) return result;

            var fonts = stylesPart.Stylesheet.Fonts?.Cast<Font>().ToList() ?? new();
            var fills = stylesPart.Stylesheet.Fills?.Cast<Fill>().ToList() ?? new();
            var borders = stylesPart.Stylesheet.Borders?.Cast<Border>().ToList() ?? new();
            var cellFormats = stylesPart.Stylesheet.CellFormats.Cast<CellFormat>().ToList();

            for (int i = 0; i < cellFormats.Count; i++)
            {
                var xf = cellFormats[i];
                var style = new OpenXmlCellStyleModel();

                if (xf.FontId?.Value != null)
                {
                    int fontId = (int)xf.FontId.Value;
                    if (fontId < fonts.Count)
                    {
                        var font = fonts[fontId];
                        style.FontName = font.FontName?.Val?.Value ?? "";
                        if (font.FontSize?.Val?.Value != null)
                            style.FontSize = (double)font.FontSize.Val.Value;
                        style.Bold = font.Bold?.Val?.Value ?? false;
                        style.Italic = font.Italic?.Val?.Value ?? false;
                    }
                }

                if (xf.Alignment != null)
                {
                    var align = xf.Alignment;
                    if (align.Horizontal?.Value != null)
                        style.HorizontalAlignment = MapHoriz(align.Horizontal.Value.ToString());
                    if (align.Vertical?.Value != null)
                        style.VerticalAlignment = MapVert(align.Vertical.Value.ToString());
                    style.WrapText = align.WrapText?.Value ?? false;
                    style.Rotation = (int)(align.TextRotation?.Value ?? 0u);
                }

                result.Add(style);
            }

            return result;
        }

        private void ReadCellFormatting(string cellAddr,
            Dictionary<string, OpenXmlCellStyleModel> cellStyles,
            List<OpenXmlCellStyleModel> allCellFormats,
            DesignerField field, StringBuilder log)
        {
            if (!cellStyles.TryGetValue(cellAddr, out var style))
            {
                // Try alternate casing or lookup by first part of address
                string addr = cellAddr;
                foreach (var kvp in cellStyles)
                {
                    if (string.Equals(kvp.Key, addr, StringComparison.OrdinalIgnoreCase))
                    {
                        style = kvp.Value;
                        break;
                    }
                }
            }

            if (style == null)
            {
                log.AppendLine($"[DesignerModelReader]   Style for {cellAddr}: NOT FOUND (no cell data)");
                return;
            }

            field.Style = new DesignerFieldStyle
            {
                FontFamily = string.IsNullOrEmpty(style.FontName) ? null : style.FontName,
                FontSize = style.FontSize > 0 ? style.FontSize : null,
                Bold = style.Bold ? true : null,
                Italic = style.Italic ? true : null,
                Underline = style.Underline ? true : null,
                FontColor = string.IsNullOrEmpty(style.FontColor) ? null : style.FontColor,
                FillColor = string.IsNullOrEmpty(style.FillColor) ? null : style.FillColor,
                HorizontalAlignment = string.IsNullOrEmpty(style.HorizontalAlignment) ? null : style.HorizontalAlignment,
                VerticalAlignment = string.IsNullOrEmpty(style.VerticalAlignment) ? null : style.VerticalAlignment,
                WrapText = style.WrapText ? true : null,
                Rotation = style.Rotation,
                NumberFormat = string.IsNullOrEmpty(style.NumberFormat) ? null : style.NumberFormat,
                BorderTop = string.IsNullOrEmpty(style.BorderTop) ? null : style.BorderTop,
                BorderBottom = string.IsNullOrEmpty(style.BorderBottom) ? null : style.BorderBottom,
                BorderLeft = string.IsNullOrEmpty(style.BorderLeft) ? null : style.BorderLeft,
                BorderRight = string.IsNullOrEmpty(style.BorderRight) ? null : style.BorderRight
            };

            // ═══════════════════════════════════════════════════════════
            // PHASE 22.1 — STAGE 23: STYLE READ (Workbook → DesignerModel)
            // ═══════════════════════════════════════════════════════════
            log.AppendLine("========================================================");
            log.AppendLine("  STAGE 23 — Style Read");
            log.AppendLine("========================================================");
            log.AppendLine($"  Cell: {cellAddr}");
            log.AppendLine($"  ✓ Font Name:     {style.FontName}");
            log.AppendLine($"  ✓ Font Size:     {style.FontSize}");
            log.AppendLine($"  ✓ Bold:          {style.Bold}");
            log.AppendLine($"  ✓ Italic:        {style.Italic}");
            log.AppendLine($"  ✓ Underline:     {style.Underline}");
            log.AppendLine($"  ✓ Font Color:    {style.FontColor}");
            log.AppendLine($"  ✓ Fill Color:    {style.FillColor}");
            log.AppendLine($"  ✓ H-Align:       {style.HorizontalAlignment}");
            log.AppendLine($"  ✓ V-Align:       {style.VerticalAlignment}");
            log.AppendLine($"  ✓ Wrap Text:     {style.WrapText}");
            log.AppendLine($"  ✓ Number Format: {style.NumberFormat}");
            log.AppendLine("========================================================");
        }

        // ═══════════════════════════════════════════════════════════════
        // 8. CELL POSITION (bounds from column widths and row heights)
        // ═══════════════════════════════════════════════════════════════

        private void ReadCellBounds(Worksheet ws, WorkbookPart wbPart, Sheet sheet,
            string cellAddr, DesignerField field, StringBuilder log)
        {
            // Parse cell address to get row and column indices
            var match = System.Text.RegularExpressions.Regex.Match(cellAddr, @"^([A-Z]+)(\d+)$");
            if (!match.Success)
            {
                log.AppendLine($"[DesignerModelReader]   Bounds for {cellAddr}: UNABLE TO PARSE");
                return;
            }

            string colPart = match.Groups[1].Value;
            int row = int.Parse(match.Groups[2].Value);
            int col = 0;
            foreach (char c in colPart)
            {
                col = col * 26 + (c - 'A' + 1);
            }

            // Calculate position from column widths and row heights
            double left = 0, top = 0, width = 0, height = 0;

            // Column widths — accumulate left position and get width
            var cols = ws.Descendants<Column>().ToList();
            bool foundCol = false;
            foreach (var c in cols)
            {
                uint min = c.Min?.Value ?? 1u;
                uint max = c.Max?.Value ?? min;
                double colWidth = (double)(c.Width?.Value ?? 8.43);

                if (col >= min && col <= max)
                {
                    if (!foundCol)
                    {
                        width = colWidth;
                        foundCol = true;
                    }
                }
                if (col > max || (!foundCol && col < min))
                {
                    left += colWidth;
                }
            }

            // Convert column width from characters to points (approx: charWidth * 7)
            left *= 7;
            width *= 7;

            // Row heights — accumulate top position and get height
            var rows = ws.Descendants<Row>().ToList();
            foreach (var r in rows)
            {
                uint rIdx = r.RowIndex?.Value ?? 0u;
                double rHeight = (double)(r.Height?.Value ?? 15.0);

                if (rIdx == row)
                {
                    height = rHeight;
                }
                if (rIdx < row)
                {
                    top += rHeight;
                }
            }

            field.Bounds = new DesignerFieldBounds
            {
                Left = left,
                Top = top,
                Width = width,
                Height = height
            };

            log.AppendLine($"[DesignerModelReader]   Bounds for {cellAddr}: L={left:F1} T={top:F1} W={width:F1} H={height:F1}");
        }

        // ═══════════════════════════════════════════════════════════════
        // 9. COMMENTS
        // ═══════════════════════════════════════════════════════════════

        private List<(string cellAddr, string text, string author)> ReadComments(
            WorksheetPart wsPart, string sheetName, StringBuilder log)
        {
            var result = new List<(string, string, string)>();

            var commentsPart = wsPart.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
            {
                log.AppendLine($"[DesignerModelReader]   Comments: NONE");
                return result;
            }

            var comments = commentsPart.Comments.Descendants<Comment>().ToList();
            log.AppendLine($"[DesignerModelReader]   Comments Found: {comments.Count}");
            log.AppendLine("----------------------------------------");

            foreach (var c in comments)
            {
                string cellRef = c.Reference?.Value ?? "";
                string text = GetCommentText(c);
                string author = "";

                result.Add((cellRef, text, author));

                log.AppendLine($"  Cell: {cellRef}");
                log.AppendLine($"  Comment:");
                log.AppendLine($"  ------------------------");
                foreach (var line in text.Split('\n'))
                    log.AppendLine($"    {line.Trim()}");
                log.AppendLine($"  ------------------------");
            }

            _logger.LogInformation("[DesignerModelReader] Comments on '{Name}': {Count}",
                sheetName, comments.Count);

            return result;
        }

        // ═══════════════════════════════════════════════════════════════
        // 10. _FIELDS SHEET
        // ═══════════════════════════════════════════════════════════════

        private List<(string cellAddr, string fieldId, string name, string type,
            string required, string readOnly, string defaultValue, string maxLength,
            string placeholder, string options, string validation, string sheetName)>
            ReadFieldsSheet(WorkbookPart wbPart, Sheet fieldsSheet, StringBuilder log)
        {
            var result = new List<(string, string, string, string, string, string, string, string,
                string, string, string, string)>();

            log.AppendLine();
            log.AppendLine("[DesignerModelReader] Reading _Fields...");
            log.AppendLine("----------------------------------------");

            var wsPart = wbPart.GetPartById(fieldsSheet.Id!.Value!) as WorksheetPart;
            if (wsPart == null)
            {
                log.AppendLine("[DesignerModelReader] ERROR: Cannot open _Fields worksheet");
                return result;
            }

            var sheetData = wsPart.Worksheet.Descendants<Row>().ToList();
            log.AppendLine($"  Rows Found: {sheetData.Count}");

            // Expected columns: Address, FieldId, FieldName, FieldType, SheetName,
            // Required, ReadOnly, DefaultValue, MaxLength, Placeholder, Options, Validation
            foreach (var row in sheetData.Skip(1)) // Skip header row
            {
                var cells = row.Descendants<Cell>().ToList();
                if (cells.Count < 2) continue;

                string addr = GetCellValue(cells.ElementAtOrDefault(0), wbPart) ?? "";
                if (string.IsNullOrWhiteSpace(addr)) continue;

                string fieldId = GetCellValue(cells.ElementAtOrDefault(1), wbPart) ?? "";
                string name = GetCellValue(cells.ElementAtOrDefault(2), wbPart) ?? "";
                string type = GetCellValue(cells.ElementAtOrDefault(3), wbPart) ?? "";
                string sheetName = GetCellValue(cells.ElementAtOrDefault(4), wbPart) ?? "";
                string required = GetCellValue(cells.ElementAtOrDefault(5), wbPart) ?? "";
                string readOnly = GetCellValue(cells.ElementAtOrDefault(6), wbPart) ?? "";
                string defaultVal = GetCellValue(cells.ElementAtOrDefault(7), wbPart) ?? "";
                string maxLen = GetCellValue(cells.ElementAtOrDefault(8), wbPart) ?? "";
                string placeholder = GetCellValue(cells.ElementAtOrDefault(9), wbPart) ?? "";
                string options = GetCellValue(cells.ElementAtOrDefault(10), wbPart) ?? "";
                string validation = GetCellValue(cells.ElementAtOrDefault(11), wbPart) ?? "";

                result.Add((addr, fieldId, name, type, required, readOnly,
                    defaultVal, maxLen, placeholder, options, validation, sheetName));

                // Log each field
                log.AppendLine();
                log.AppendLine($"  Field: {fieldId}");
                log.AppendLine($"  --------");
                log.AppendLine($"    ID: {fieldId}");
                log.AppendLine($"    Cell: {addr}");
                log.AppendLine($"    Type: {type}");
                log.AppendLine($"    Required: {required}");
                log.AppendLine($"    ReadOnly: {readOnly}");
                log.AppendLine($"    MaxLength: {maxLen}");
                if (!string.IsNullOrEmpty(placeholder))
                    log.AppendLine($"    Placeholder: {placeholder}");
                if (!string.IsNullOrEmpty(defaultVal))
                    log.AppendLine($"    Default: {defaultVal}");
            }

            log.AppendLine();
            log.AppendLine("[DesignerModelReader] _Fields Summary");
            log.AppendLine($"  Fields Loaded: {result.Count}");

            _logger.LogInformation("[DesignerModelReader] _Fields: {Count} fields loaded", result.Count);

            return result;
        }

        // ═══════════════════════════════════════════════════════════════
        // 11. EXCELOUTPUTSETTING
        // ═══════════════════════════════════════════════════════════════

        private void ReadExcelOutputSetting(WorkbookPart wbPart, Sheet sheet,
            DesignerModel model, StringBuilder log)
        {
            if (sheet.Id?.Value == null) return;

            var wsPart = wbPart.GetPartById(sheet.Id.Value) as WorksheetPart;
            if (wsPart == null)
            {
                log.AppendLine("[DesignerModelReader]   ERROR: Cannot open ExcelOutputSetting worksheet");
                return;
            }

            var rows = wsPart.Worksheet.Descendants<Row>().ToList();
            log.AppendLine($"  Rows: {rows.Count}");

            // Read key-value pairs from the sheet
            var settings = new Dictionary<string, string>();
            foreach (var row in rows)
            {
                var cells = row.Descendants<Cell>().ToList();
                if (cells.Count >= 2)
                {
                    string key = GetCellValue(cells[0], wbPart) ?? "";
                    string val = GetCellValue(cells[1], wbPart) ?? "";
                    if (!string.IsNullOrEmpty(key))
                        settings[key] = val;
                }
            }

            // Log what we found
            foreach (var kvp in settings)
            {
                log.AppendLine($"  {kvp.Key}: {kvp.Value}");
            }

            log.AppendLine();
            log.AppendLine("[DesignerModelReader] ExcelOutputSetting Loaded Successfully");

            _logger.LogInformation("[DesignerModelReader] ExcelOutputSetting: {Count} settings",
                settings.Count);
        }

        // ═══════════════════════════════════════════════════════════════
        // 12. DATA VALIDATION
        // ═══════════════════════════════════════════════════════════════

        private void ReadDataValidation(Worksheet ws, string cellAddr,
            DesignerField field, StringBuilder log)
        {
            var dataValidations = ws.Descendants<DataValidations>().FirstOrDefault();
            if (dataValidations == null) return;

            foreach (var dv in dataValidations.Descendants<DataValidation>())
            {
                // Compute sequence of references as a string
                string sqref = "";
                var sqrefList = dv.SequenceOfReferences;
                if (sqrefList != null)
                    sqref = sqrefList.InnerText ?? "";

                if (sqref.Contains(cellAddr) || sqref.Contains(cellAddr.Split(':')[0]))
                {
                    // Read Type and Operator safely
                    string dvTypeStr = "any";
                    if (dv.Type != null)
                        dvTypeStr = dv.Type.Value.ToString() ?? "any";

                    string? dvOpStr = null;
                    if (dv.Operator != null)
                        dvOpStr = dv.Operator.Value.ToString();

                    // Read ErrorMessage/ErrorTitle
                    string? errorMessage = null;
                    // DataValidation error message — StringValue uses .Value not .Text
                    if (dv.Error != null)
                        errorMessage = dv.Error.Value;

                    field.Validation = new DesignerFieldValidation
                    {
                        Type = MapValidationType(dvTypeStr),
                        Operator = dvOpStr,
                        Formula1 = dv.Formula1?.Text,
                        Formula2 = dv.Formula2?.Text,
                        AllowBlank = dv.AllowBlank?.Value ?? true,
                        ErrorMessage = errorMessage
                    };

                    log.AppendLine("[DesignerModelReader]   Validation");
                    log.AppendLine($"    Type: {field.Validation.Type}");
                    log.AppendLine($"    Formula1: {field.Validation.Formula1}");
                    log.AppendLine($"    Operator: {field.Validation.Operator}");
                    break;
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // 13. IMAGES
        // ═══════════════════════════════════════════════════════════════

        private void ReadImages(WorksheetPart wsPart, string sheetName, StringBuilder log)
        {
            var drawPart = wsPart.DrawingsPart;
            if (drawPart == null)
            {
                log.AppendLine($"[DesignerModelReader]   Images: NONE");
                return;
            }

            int imageCount = drawPart.ImageParts.Count();
            log.AppendLine($"[DesignerModelReader]   Images Found: {imageCount}");
        }

        // ═══════════════════════════════════════════════════════════════
        // 14. SHAPES
        // ═══════════════════════════════════════════════════════════════

        private void ReadShapes(WorksheetPart wsPart, string sheetName, StringBuilder log)
        {
            var drawPart = wsPart.DrawingsPart;
            if (drawPart == null)
            {
                log.AppendLine($"[DesignerModelReader]   Shapes: NONE");
                return;
            }

            var drawing = drawPart.WorksheetDrawing;
            if (drawing == null)
            {
                log.AppendLine($"[DesignerModelReader]   Shapes: NONE");
                return;
            }

            int shapeCount = 0;
            foreach (var child in drawing.ChildElements)
            {
                if (child.LocalName == "shape" || child.LocalName == "connector" ||
                    child.LocalName == "groupShape" || child.LocalName == "picture")
                {
                    shapeCount++;
                }
            }

            log.AppendLine($"[DesignerModelReader]   Shapes Found: {shapeCount}");
        }

        // ═══════════════════════════════════════════════════════════════
        // HELPERS
        // ═══════════════════════════════════════════════════════════════

        private string GetCommentText(Comment comment)
        {
            try
            {
                var text = comment.Descendants<Text>().FirstOrDefault();
                return text?.Text?.Trim() ?? "";
            }
            catch { return ""; }
        }

        private string GetCommentForCell(List<(string addr, string text, string author)> comments,
            string cellAddr)
        {
            foreach (var (addr, text, _) in comments)
            {
                if (string.Equals(addr, cellAddr, StringComparison.OrdinalIgnoreCase) ||
                    addr.Contains(cellAddr))
                    return text;
            }
            return "";
        }

        private string? GetCellValue(Cell? cell, WorkbookPart wbPart)
        {
            if (cell == null) return null;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sstPart = wbPart.SharedStringTablePart;
                if (sstPart?.SharedStringTable == null) return null;
                int idx = int.Parse(cell.CellValue?.Text ?? "0");
                var items = sstPart.SharedStringTable.Elements<SharedStringItem>().ToList();
                if (idx >= 0 && idx < items.Count)
                    return items[idx].Text?.Text ?? items[idx].InnerText;
                return null;
            }

            return cell.CellValue?.Text;
        }

        private string GetColorHex(DocumentFormat.OpenXml.Spreadsheet.Color? color)
        {
            if (color == null) return "";
            if (color.Rgb?.Value != null)
            {
                string rgb = color.Rgb.Value;
                if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase)))
                    return "#" + rgb[2..];
                if (rgb.Length >= 6) return "#" + rgb;
            }
            if (color.Auto?.Value == true) return "#000000";
            return "";
        }

        private string? GetFillColorHex(ForegroundColor? fgColor)
        {
            if (fgColor == null) return null;
            if (fgColor.Rgb?.Value != null)
            {
                string rgb = fgColor.Rgb.Value;
                if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase)))
                    return "#" + rgb[2..];
                if (rgb.Length >= 6) return "#" + rgb;
            }
            if (fgColor.Auto?.Value == true) return "#000000";
            return null;
        }

        private string? GetFillColorHex(BackgroundColor? bgColor)
        {
            if (bgColor == null) return null;
            if (bgColor.Rgb?.Value != null)
            {
                string rgb = bgColor.Rgb.Value;
                if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase)))
                    return "#" + rgb[2..];
                if (rgb.Length >= 6) return "#" + rgb;
            }
            if (bgColor.Auto?.Value == true) return "#000000";
            return null;
        }

        private string GetBorderCss(BorderPropertiesType? border)
        {
            if (border?.Style == null) return "";
            var borderStyle = border.Style.Value;
            if (borderStyle.ToString() == "None") return "";

            string width = "1px";
            string lineStyle = "solid";

            // BorderStyleValues has CS9135 issues in switch; use string comparison instead
            string bStyle = borderStyle.ToString();
            switch (bStyle)
            {
                case "Thin": width = "1px"; break;
                case "Medium": width = "2px"; break;
                case "Thick": width = "3px"; break;
                case "Double": width = "3px"; lineStyle = "double"; break;
                case "Dotted": lineStyle = "dotted"; break;
                case "Dashed":
                case "DashDot":
                case "DashDotDot": lineStyle = "dashed"; break;
                case "MediumDashed":
                case "MediumDashDot":
                case "MediumDashDotDot": width = "2px"; lineStyle = "dashed"; break;
                case "SlantDashDot": width = "2px"; lineStyle = "dashed"; break;
            }

            string color = GetColorHex(border.Color);
            return $"{width} {lineStyle} {(string.IsNullOrEmpty(color) ? "#000000" : color)}";
        }

        private string MapHoriz(string? align)
        {
            return align switch
            {
                "left" => "left",
                "center" => "center",
                "right" => "right",
                "centerContinuous" => "center",
                "fill" => "left",
                "justify" => "justify",
                "distributed" => "justify",
                _ => ""
            };
        }

        private string MapVert(string? align)
        {
            return align switch
            {
                "top" => "top",
                "center" => "middle",
                "bottom" => "bottom",
                "justify" => "middle",
                "distributed" => "middle",
                _ => ""
            };
        }

        private string MapFieldType(string type)
        {
            return (type ?? "").ToLower() switch
            {
                "text" or "textbox" or "txtbox" => "text",
                "number" or "numeric" or "num" => "number",
                "date" or "datetime" => "date",
                "checkbox" or "check" => "checkbox",
                "signature" or "sig" => "signature",
                "dropdown" or "list" or "combo" => "dropdown",
                "calculated" or "calc" or "formula" => "calculated",
                _ => "text"
            };
        }

        private string MapValidationType(string type)
        {
            return type.ToLower() switch
            {
                "whole" => "whole",
                "decimal" => "decimal",
                "list" => "list",
                "date" => "date",
                "time" => "time",
                "textlength" => "textLength",
                "custom" => "custom",
                _ => "any"
            };
        }

        /// <summary>
        /// Extract text from a cell that may store inline strings (t="inlineStr") or shared strings (t="s").
        /// </summary>
        private static string? GetCellInlineText(Cell cell)
        {
            if (cell == null) return null;

            // Inline string
            if (cell.InlineString != null && cell.InlineString.Text != null)
                return cell.InlineString.Text.Text;

            // Shared string
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && cell.CellValue != null)
            {
                if (int.TryParse(cell.CellValue.Text, out int ssIndex))
                {
                    // Cannot resolve shared string without the SharedStringTablePart
                    // This is a best-effort; the config uses inline strings
                    return $"(shared string index {ssIndex})";
                }
            }

            // Plain text value
            if (cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text))
                return cell.CellValue.Text;

            return null;
        }

        /// <summary>
        /// Normalize a cell address or range to a canonical cell reference (e.g., "$A$1:$B$2" → "A1").
        /// Removes dollar signs, splits ranges, and returns the first cell.
        /// </summary>
        private static string NormalizeCellAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return "";
            string trimmed = address.Trim();
            // Handle ranges: "A1:B2" → "A1", "$A$1:$B$2" → "A1"
            if (trimmed.Contains(':'))
                trimmed = trimmed.Split(':')[0];
            // Remove dollar signs: "$A$1" → "A1"
            trimmed = trimmed.Replace("$", "");
            return trimmed.ToUpperInvariant();
        }

        /// <summary>
        /// Map CellStyle alignment horizontal value to DesignerFieldStyle horizontal alignment.
        /// </summary>
        private static string MapHorizontalAlignment(string? horiz)
        {
            return (horiz ?? "").ToLowerInvariant() switch
            {
                "left" => "left",
                "center" or "centre" => "center",
                "right" => "right",
                "general" => "general",
                "fill" => "fill",
                "justify" => "justify",
                _ => "" // let the default from Excel detection stand
            };
        }

        /// <summary>
        /// Map CellStyle alignment vertical value to DesignerFieldStyle vertical alignment.
        /// </summary>
        private static string MapVerticalAlignment(string? vert)
        {
            return (vert ?? "").ToLowerInvariant() switch
            {
                "top" => "top",
                "center" or "middle" => "middle",
                "bottom" => "bottom",
                "justify" => "justify",
                "distributed" => "distributed",
                _ => "" // let the default from Excel detection stand
            };
        }

        /// <summary>
        /// Lightweight model for cell formatting read from OpenXML styles.
        /// </summary>
        public class OpenXmlCellStyleModel
        {
            public string FontName { get; set; } = "";
            public double FontSize { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public bool Underline { get; set; }
            public string FontColor { get; set; } = "";
            public string FillColor { get; set; } = "";
            public string HorizontalAlignment { get; set; } = "";
            public string VerticalAlignment { get; set; } = "";
            public bool WrapText { get; set; }
            public int Rotation { get; set; }
            public string NumberFormat { get; set; } = "";
            public string BorderTop { get; set; } = "";
            public string BorderBottom { get; set; } = "";
            public string BorderLeft { get; set; } = "";
            public string BorderRight { get; set; } = "";
        }
    }
}
