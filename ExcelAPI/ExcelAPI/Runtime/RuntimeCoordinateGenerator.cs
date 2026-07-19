using ExcelAPI.Models;
using ExcelAPI.Runtime;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Production service that builds a RuntimeForm from Excel COM capture results.
    /// This is the single source of truth for runtime field coordinates.
    ///
    /// Called once during upload to persist runtime metadata.
    /// Called on every GET request to serve runtime data to the frontend.
    ///
    /// Phase 14: Added ratio-based coordinate storage (left_ratio, top_ratio, etc.)
    /// alongside pixel coordinates for legacy ConMas compatibility.
    /// Ratios are relative to page dimensions: ratio = pixel / pageDimension.
    ///
    /// Phase 4.0: Legacy SaveMetadata removed — SaveFromDefinition is the only path.
    ///
    /// No OpenXML parsing. No GeometryBuilder. No CoordinateEngine.
    /// Coordinates come directly from Excel COM measurements or WbDef.
    /// </summary>
    public class RuntimeCoordinateGenerator
    {
        private readonly ILogger<RuntimeCoordinateGenerator> _logger;

        public RuntimeCoordinateGenerator(
            ILogger<RuntimeCoordinateGenerator> logger)
        {
            _logger = logger;
        }

        // ── Primary path: WorkbookDefinition ───────────────────────────

        /// <summary>
        /// Save workbook definitions as persistent runtime metadata.
        /// Requires InternalWorkbookDefinition to be populated on the CaptureResult.
        /// Called during upload — persists field coordinates and background image URL.
        /// </summary>
        public void SaveFromWbDef(
            CaptureResult capture,
            string templateId,
            string formsDir,
            string workbookName,
            int pageWidthPx,
            int pageHeightPx,
            string? backgroundImageUrl = null)
        {
            if (capture.InternalWorkbookDefinition == null)
            {
                _logger.LogWarning(
                    "[WBDF] No InternalWorkbookDefinition available for '{Template}' — cannot persist runtime metadata",
                    templateId);
                return;
            }

            SaveFromDefinition(capture.InternalWorkbookDefinition, templateId, formsDir,
                workbookName, pageWidthPx, pageHeightPx, backgroundImageUrl);
        }

        /// <summary>
        /// Save WorkbookDefinition directly as persistent runtime metadata.
        /// Preferred when a WbDef is available (new COM capture path).
        /// </summary>
        public void SaveFromDefinition(
            WbDef.WorkbookDefinition wbDef,
            string templateId,
            string formsDir,
            string workbookName,
            int pageWidthPx,
            int pageHeightPx,
            string? backgroundImageUrl = null)
        {
            try
            {
                Directory.CreateDirectory(formsDir);
                string path = Path.Combine(formsDir, templateId + ".runtime.json");

                double dpi = 300.0;

                var sheetsList = new List<object>();
                foreach (var sheet in wbDef.Sheets)
                {
                    var fieldList = new List<object>();
                    int tabIndex = 0;

                    foreach (var field in sheet.Fields)
                    {
                        double leftPx = WbDef.Rectangle.PtToPx(field.BoundsPt.Left, dpi);
                        double topPx = WbDef.Rectangle.PtToPx(field.BoundsPt.Top, dpi);
                        double widthPx = WbDef.Rectangle.PtToPx(field.BoundsPt.Width, dpi);
                        double heightPx = WbDef.Rectangle.PtToPx(field.BoundsPt.Height, dpi);

                        double leftRatio = pageWidthPx > 0 ? Math.Round(leftPx / pageWidthPx, 7) : 0;
                        double topRatio = pageHeightPx > 0 ? Math.Round(topPx / pageHeightPx, 7) : 0;
                        double widthRatio = pageWidthPx > 0 ? Math.Round(widthPx / pageWidthPx, 7) : 0;
                        double heightRatio = pageHeightPx > 0 ? Math.Round(heightPx / pageHeightPx, 7) : 0;

                        fieldList.Add(new
                        {
                            id = field.Id,
                            name = field.Name,
                            cellReference = field.Cell.Address,
                            leftPx = Math.Round(leftPx, 1),
                            topPx = Math.Round(topPx, 1),
                            widthPx = Math.Round(widthPx, 1),
                            heightPx = Math.Round(heightPx, 1),
                            leftRatio,
                            topRatio,
                            widthRatio,
                            heightRatio,
                            dataType = field.Type.ToString().ToLowerInvariant(),
                            isMerged = field.MergeInfo?.IsMerged ?? false,
                            mergeRange = field.MergeInfo?.MergeAddress,
                            fontSize = field.Style?.Font?.SizePt ?? 11,
                            bold = field.Style?.Font?.Bold ?? false,
                            readOnly = field.Locked,
                            required = field.Required
                        });
                        tabIndex++;
                    }

                    sheetsList.Add(new
                    {
                        name = sheet.Name,
                        index = sheet.Index,
                        pageWidthPx,
                        pageHeightPx,
                        fields = fieldList,
                        backgroundImage = backgroundImageUrl
                    });
                }

                var metadata = new
                {
                    version = "1.0",
                    capturedAt = DateTime.UtcNow.ToString("o"),
                    workbookName,
                    dpi,
                    scaleX = Math.Round(WbDef.Rectangle.PtToPx(1, dpi), 6),
                    scaleY = Math.Round(WbDef.Rectangle.PtToPx(1, dpi), 6),
                    pageWidthPx,
                    pageHeightPx,
                    sheets = sheetsList
                };

                string json = System.Text.Json.JsonSerializer.Serialize(metadata,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
                System.IO.File.WriteAllText(path, json);

                _logger.LogInformation(
                    "[WBDF] Metadata saved: {Path} ({Fields} field(s) across {Sheets} sheet(s), {W}x{H}px)",
                    path, wbDef.Sheets.Sum(s => s.Fields.Count), wbDef.Sheets.Count, pageWidthPx, pageHeightPx);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[WBDF] Failed to save runtime metadata, falling back to legacy path");
                // Fallback intentionally not called here — caller should handle
            }
        }

        // ── Metadata readback (unchanged) ────────────────────────────────

        /// <summary>
        /// Load COM runtime metadata as a RuntimeForm.
        /// Returns null if no .runtime.json exists (caller falls back to OpenXML).
        /// </summary>
        public RuntimeForm? LoadMetadata(string templateId, string formsDir)
        {
            try
            {
                string path = Path.Combine(formsDir, templateId + ".runtime.json");
                if (!System.IO.File.Exists(path))
                    return null;

                string json = System.IO.File.ReadAllText(path);
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var root = doc.RootElement;

                string workbookName = root.TryGetProperty("workbookName", out var wn)
                    ? wn.GetString() ?? "" : "";
                int dpi = root.TryGetProperty("dpi", out var d) ? d.GetInt32() : 300;
                int pageWidthPx = root.TryGetProperty("pageWidthPx", out var pw) ? pw.GetInt32() : 0;
                int pageHeightPx = root.TryGetProperty("pageHeightPx", out var ph) ? ph.GetInt32() : 0;

                var runtimeForm = new RuntimeForm
                {
                    WorkbookName = workbookName,
                    Title = workbookName,
                    Dpi = dpi,
                    Scale = 1.0,
                    Version = "1.0",
                    PageWidth = pageWidthPx,
                    PageHeight = pageHeightPx
                };

                if (root.TryGetProperty("sheets", out var sheetsElement))
                {
                    int sheetIndex = 0;
                    foreach (var sheetElem in sheetsElement.EnumerateArray())
                    {
                        string sheetName = "Sheet" + (sheetIndex + 1);
                        if (sheetElem.TryGetProperty("name", out var sn))
                        {
                            string? n = sn.GetString();
                            if (!string.IsNullOrEmpty(n))
                                sheetName = n;
                        }

                        int sPageWidth = sheetElem.TryGetProperty("pageWidthPx", out var spw)
                            ? spw.GetInt32() : pageWidthPx;
                        int sPageHeight = sheetElem.TryGetProperty("pageHeightPx", out var sph)
                            ? sph.GetInt32() : pageHeightPx;                                var runtimeSheet = new RuntimeSheet
                                {
                                    Name = sheetName,
                                    Index = sheetIndex,
                                    PageWidthPx = sPageWidth,
                                    PageHeightPx = sPageHeight
                                };

                                // Read backgroundImage from sheet metadata
                                if (sheetElem.TryGetProperty("backgroundImage", out var bg))
                                {
                                    string? bgUrl = bg.GetString();
                                    if (!string.IsNullOrEmpty(bgUrl))
                                        runtimeSheet.BackgroundImage = bgUrl;
                                }

                        if (sheetElem.TryGetProperty("fields", out var fieldsElement))
                        {
                            int tabIndex = 0;
                            foreach (var fieldElem in fieldsElement.EnumerateArray())
                            {
                                string fieldId = "field_" + tabIndex;
                                if (fieldElem.TryGetProperty("id", out var fid))
                                {
                                    string? fi = fid.GetString();
                                    if (!string.IsNullOrEmpty(fi))
                                        fieldId = fi;
                                }

                                string cellRef = "";
                                if (fieldElem.TryGetProperty("cellReference", out var cr))
                                {
                                    string? crs = cr.GetString();
                                    if (crs != null)
                                        cellRef = crs;
                                }

                                string fieldName = fieldElem.TryGetProperty("name", out var nm)
                                    ? nm.GetString() ?? "" : "";

                                runtimeSheet.Fields.Add(new RuntimeField
                                {
                                    Id = fieldId,
                                    Name = fieldName,
                                    CellReference = cellRef,
                                    LeftPx = fieldElem.TryGetProperty("leftPx", out var lx)
                                        ? lx.GetDouble() : 0,
                                    TopPx = fieldElem.TryGetProperty("topPx", out var ty)
                                        ? ty.GetDouble() : 0,
                                    WidthPx = fieldElem.TryGetProperty("widthPx", out var wx)
                                        ? wx.GetDouble() : 0,
                                    HeightPx = fieldElem.TryGetProperty("heightPx", out var hx)
                                        ? hx.GetDouble() : 0,
                                    LeftRatio = fieldElem.TryGetProperty("leftRatio", out var lr)
                                        ? lr.GetDouble() : 0,
                                    TopRatio = fieldElem.TryGetProperty("topRatio", out var tr)
                                        ? tr.GetDouble() : 0,
                                    WidthRatio = fieldElem.TryGetProperty("widthRatio", out var wr)
                                        ? wr.GetDouble() : 0,
                                    HeightRatio = fieldElem.TryGetProperty("heightRatio", out var hr)
                                        ? hr.GetDouble() : 0,
                                    DataType = fieldElem.TryGetProperty("dataType", out var dt)
                                        ? (dt.GetString() ?? "text") : "text",
                                    IsMerged = fieldElem.TryGetProperty("isMerged", out var im)
                                        ? im.GetBoolean() : false,
                                    MergeRange = fieldElem.TryGetProperty("mergeRange", out var mr)
                                        ? mr.GetString() : null,
                                    FontSize = fieldElem.TryGetProperty("fontSize", out var fs)
                                        ? fs.GetDouble() : 11,
                                    Bold = fieldElem.TryGetProperty("bold", out var bd)
                                        ? bd.GetBoolean() : false,
                                    ReadOnly = fieldElem.TryGetProperty("readOnly", out var ro)
                                        ? ro.GetBoolean() : false,
                                    Required = fieldElem.TryGetProperty("required", out var rq)
                                        ? rq.GetBoolean() : false,
                                    TabIndex = tabIndex
                                });
                                tabIndex++;
                            }
                        }

                        runtimeForm.Sheets.Add(runtimeSheet);
                        sheetIndex++;
                    }
                }

                return runtimeForm;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to load runtime metadata, falling back to OpenXML");
                return null;
            }
        }


    }
}
