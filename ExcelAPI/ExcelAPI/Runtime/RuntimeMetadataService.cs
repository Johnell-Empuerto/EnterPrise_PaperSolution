using ExcelAPI.Models;
using System.Text.Json;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Persists and loads runtime metadata using Excel COM field rectangles
    /// as the single source of truth. Field coordinates are already in final
    /// pixel space (the COM pipeline applied margin/centering/origin offsets).
    ///
    /// This service eliminates the need for OpenXmlParser, GeometryBuilder,
    /// CoordinateEngine, and FieldDetector in the Runtime GET path.
    ///
    /// File: Forms/{templateId}.runtime.json
    /// </summary>
    public class RuntimeMetadataService
    {
        private readonly ILogger<RuntimeMetadataService> _logger;

        public RuntimeMetadataService(ILogger<RuntimeMetadataService> logger)
        {
            _logger = logger;
        }

        // ── Save ───────────────────────────────────────────────────────

        /// <summary>
        /// Save COM capture result as persistent runtime metadata.
        /// Called once during upload — eliminates OpenXML recalculation on every GET.
        /// </summary>
        public void Save(CaptureResult capture, string templateId, string formsDir, string workbookName)
        {
            try
            {
                Directory.CreateDirectory(formsDir);
                string path = Path.Combine(formsDir, $"{templateId}.runtime.json");

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

                var sheets = new List<object>();
                var fieldList = new List<object>();

                foreach (var f in capture.Fields)
                {
                    fieldList.Add(new
                    {
                        id = f.Id,
                        cellReference = f.Cell,
                        leftPx = Math.Round(f.Left, 1),
                        topPx = Math.Round(f.Top, 1),
                        widthPx = Math.Round(f.Width, 1),
                        heightPx = Math.Round(f.Height, 1),
                        dataType = f.Type.ToLowerInvariant(),
                        isMerged = f.IsMerged,
                        mergeRange = f.MergeAddress,
                        // Default style — frontend handles these gracefully
                        fontSize = 11,
                        bold = false,
                        readOnly = false,
                        required = false
                    });
                }

                sheets.Add(new
                {
                    name = "Sheet1",
                    index = 0,
                    pageWidthPx = pngWidth,
                    pageHeightPx = pngHeight,
                    fields = fieldList
                });

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

                string json = JsonSerializer.Serialize(metadata,
                    new JsonSerializerOptions { WriteIndented = false });
                System.IO.File.WriteAllText(path, json);

                _logger.LogInformation(
                    "[RUNTIME] COM metadata saved: {Path} ({Fields} field(s), {W}x{H}px)",
                    path, capture.Fields.Count, pngWidth, pngHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to save COM runtime metadata");
            }
        }

        // ── Load ───────────────────────────────────────────────────────

        /// <summary>
        /// Load COM runtime metadata as a RuntimeForm.
        /// Returns null if no .runtime.json file exists (caller falls back to OpenXML).
        /// </summary>
        public RuntimeForm? Load(string templateId, string formsDir)
        {
            try
            {
                string path = Path.Combine(formsDir, $"{templateId}.runtime.json");
                if (!System.IO.File.Exists(path))
                {
                    _logger.LogInformation("[RUNTIME] No COM metadata at {Path}, falling back to OpenXML", path);
                    return null;
                }

                string json = System.IO.File.ReadAllText(path);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                string workbookName = root.TryGetProperty("workbookName", out var wn) ? wn.GetString() ?? "" : "";
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

                // Parse sheets
                if (root.TryGetProperty("sheets", out var sheetsElement))
                {
                    int sheetIndex = 0;
                    foreach (var sheetElem in sheetsElement.EnumerateArray())
                    {
                        string sheetName = sheetElem.TryGetProperty("name", out var sn)
                            ? sn.GetString() ?? $"Sheet{sheetIndex + 1}"
                            : $"Sheet{sheetIndex + 1}";
                        int sPageWidth = sheetElem.TryGetProperty("pageWidthPx", out var spw)
                            ? spw.GetInt32() : pageWidthPx;
                        int sPageHeight = sheetElem.TryGetProperty("pageHeightPx", out var sph)
                            ? sph.GetInt32() : pageHeightPx;

                        var runtimeSheet = new RuntimeSheet
                        {
                            Name = sheetName,
                            Index = sheetIndex,
                            PageWidthPx = sPageWidth,
                            PageHeightPx = sPageHeight
                        };

                        // Parse fields
                        if (sheetElem.TryGetProperty("fields", out var fieldsElement))
                        {
                            int tabIndex = 0;
                            foreach (var fieldElem in fieldsElement.EnumerateArray())
                            {
                                var rf = new RuntimeField
                                {
                                    Id = fieldElem.TryGetProperty("id", out var fid)
                                        ? fid.GetString() ?? $"field_{tabIndex}" : $"field_{tabIndex}",
                                    CellReference = fieldElem.TryGetProperty("cellReference", out var cr)
                                        ? cr.GetString() ?? "" : "",
                                    LeftPx = fieldElem.TryGetProperty("leftPx", out var lx)
                                        ? lx.GetDouble() : 0,
                                    TopPx = fieldElem.TryGetProperty("topPx", out var ty)
                                        ? ty.GetDouble() : 0,
                                    WidthPx = fieldElem.TryGetProperty("widthPx", out var wx)
                                        ? wx.GetDouble() : 0,
                                    HeightPx = fieldElem.TryGetProperty("heightPx", out var hx)
                                        ? hx.GetDouble() : 0,
                                    DataType = fieldElem.TryGetProperty("dataType", out var dt)
                                        ? dt.GetString() ?? "text" : "text",
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
                                };

                                runtimeSheet.Fields.Add(rf);
                                tabIndex++;
                            }
                        }

                        runtimeForm.Sheets.Add(runtimeSheet);
                        sheetIndex++;
                    }
                }

                _logger.LogInformation(
                    "[RUNTIME] COM metadata loaded: {Path} ({Sheets} sheet(s), {Fields} field(s))",
                    path, runtimeForm.Sheets.Count,
                    runtimeForm.Sheets.Sum(s => s.Fields.Count));

                return runtimeForm;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to load COM runtime metadata, falling back to OpenXML");
                return null;
            }
        }
    }
}
