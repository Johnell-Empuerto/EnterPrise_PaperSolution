using ExcelAPI.Models;
using ExcelAPI.Runtime;

namespace ExcelAPI.Services
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
    /// No OpenXML parsing. No GeometryBuilder. No CoordinateEngine.
    /// Coordinates come directly from Excel COM measurements.
    /// </summary>
    public class RuntimeCoordinateGenerator
    {
        private readonly ILogger<RuntimeCoordinateGenerator> _logger;

        public RuntimeCoordinateGenerator(
            ILogger<RuntimeCoordinateGenerator> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Save COM capture result as persistent runtime metadata.
        /// Called during upload — persists field rectangles in final pixel coordinates
        /// and ratio-based coordinates for legacy ConMas compatibility.
        /// Also saves the background image URL so the frontend can load the exact
        /// image that was used during coordinate computation.
        /// </summary>
        public void SaveMetadata(CaptureResult capture, string templateId, string formsDir, string workbookName, string? backgroundImageUrl = null)
        {
            try
            {
                Directory.CreateDirectory(formsDir);
                string path = Path.Combine(formsDir, templateId + ".runtime.json");

                int pngWidth = capture.Page?.Width ?? 0;
                int pngHeight = capture.Page?.Height ?? 0;
                double dpi = 300.0;
                double scaleX = dpi / 72.0;
                double scaleY = dpi / 72.0;

                if (capture.PageSetup != null)
                {
                    if (capture.PageSetup.ActualScaleX > 0)
                        scaleX = capture.PageSetup.ActualScaleX;
                    if (capture.PageSetup.ActualScaleY > 0)
                        scaleY = capture.PageSetup.ActualScaleY;
                }

                var fieldList = new List<object>();
                int tabIndex = 0;
                foreach (var f in capture.Fields)
                {
                    double leftRatio = pngWidth > 0 ? Math.Round(f.Left / pngWidth, 7) : 0;
                    double topRatio = pngHeight > 0 ? Math.Round(f.Top / pngHeight, 7) : 0;
                    double widthRatio = pngWidth > 0 ? Math.Round(f.Width / pngWidth, 7) : 0;
                    double heightRatio = pngHeight > 0 ? Math.Round(f.Height / pngHeight, 7) : 0;

                    string fieldName = !string.IsNullOrWhiteSpace(f.Name)
                        ? f.Name
                        : $"p1f{tabIndex + 1}";

                    fieldList.Add(new
                    {
                        id = $"page1field{tabIndex + 1}",
                        name = fieldName,
                        cellReference = f.Cell,
                        leftPx = Math.Round(f.Left, 1),
                        topPx = Math.Round(f.Top, 1),
                        widthPx = Math.Round(f.Width, 1),
                        heightPx = Math.Round(f.Height, 1),
                        leftRatio = leftRatio,
                        topRatio = topRatio,
                        widthRatio = widthRatio,
                        heightRatio = heightRatio,
                        dataType = f.Type.ToLowerInvariant(),
                        isMerged = f.IsMerged,
                        mergeRange = f.MergeAddress,
                        fontSize = 11,
                        bold = false,
                        readOnly = false,
                        required = false
                    });

                    tabIndex++;
                }

                var sheets = new List<object>
                {
                    new
                    {
                        name = "Sheet1",
                        index = 0,
                        pageWidthPx = pngWidth,
                        pageHeightPx = pngHeight,
                        fields = fieldList,
                        backgroundImage = backgroundImageUrl
                    }
                };

                var metadata = new
                {
                    version = "1.0",
                    capturedAt = DateTime.UtcNow.ToString("o"),
                    workbookName,
                    dpi = 300,
                    scaleX = Math.Round(scaleX, 6),
                    scaleY = Math.Round(scaleY, 6),
                    pageWidthPx = pngWidth,
                    pageHeightPx = pngHeight,
                    sheets
                };

                string json = System.Text.Json.JsonSerializer.Serialize(metadata,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
                System.IO.File.WriteAllText(path, json);

                _logger.LogInformation(
                    "[RUNTIME] Metadata saved: {Path} ({Fields} field(s), {W}x{H}px)",
                    path, capture.Fields.Count, pngWidth, pngHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to save runtime metadata");
            }
        }

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
