using System.Net.Http.Json;
using System.Text.Json.Serialization;
using System.Text.Json;

namespace ExcelAPI.Services;

public class PythonPreviewResponse
{
    [JsonPropertyName("success")] public bool Success { get; set; }
    /// <summary>Multi-page result — one entry per visible worksheet.</summary>
    [JsonPropertyName("pages")] public List<PythonPreviewPageResult>? Pages { get; set; }
}

public class PythonPreviewPageResult
{
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("backgroundImage")] public string BackgroundImage { get; set; } = "";
    [JsonPropertyName("page")] public PythonPreviewPage? Page { get; set; }
    [JsonPropertyName("fields")] public List<PythonPreviewField>? Fields { get; set; }
}

public class PythonPreviewPage
{
    [JsonPropertyName("width")] public int Width { get; set; }
    [JsonPropertyName("height")] public int Height { get; set; }
}

public class PythonPreviewField
{
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("type")] public string Type { get; set; } = "";
    [JsonPropertyName("cellAddr")] public string CellAddr { get; set; } = "";
    [JsonPropertyName("input_parameter")] public string InputParameter { get; set; } = "";
    [JsonPropertyName("left_ratio")] public double LeftRatio { get; set; }
    [JsonPropertyName("top_ratio")] public double TopRatio { get; set; }
    [JsonPropertyName("right_ratio")] public double RightRatio { get; set; }
    [JsonPropertyName("bottom_ratio")] public double BottomRatio { get; set; }
}

public class PythonRenderRequest
{
    [JsonPropertyName("template_id")]
    public int? TemplateId { get; set; }
    [JsonPropertyName("xlsx_path")]
    public string? XlsxPath { get; set; }
    [JsonPropertyName("output_dir")]
    public string? OutputDir { get; set; }
}

public class PythonRenderField
{
    [JsonPropertyName("id")] public string Id { get; set; } = "";
    [JsonPropertyName("label")] public string Label { get; set; } = "";
    [JsonPropertyName("left_px")] public double LeftPx { get; set; }
    [JsonPropertyName("top_px")] public double TopPx { get; set; }
    [JsonPropertyName("width_px")] public double WidthPx { get; set; }
    [JsonPropertyName("height_px")] public double HeightPx { get; set; }
    [JsonPropertyName("type")] public string Type { get; set; } = "text";
    [JsonPropertyName("required")] public bool Required { get; set; }
}

public class PythonRenderResponse
{
    [JsonPropertyName("page_width")] public int PageWidth { get; set; }
    [JsonPropertyName("page_height")] public int PageHeight { get; set; }
    [JsonPropertyName("background_image")] public string BackgroundImage { get; set; } = "";
    [JsonPropertyName("fields")] public List<PythonRenderField> Fields { get; set; } = new();
}

public class PythonRenderService
{
    private readonly HttpClient _http;
    private readonly ILogger<PythonRenderService> _logger;
    private readonly IWebHostEnvironment _env;

    public PythonRenderService(HttpClient http, ILogger<PythonRenderService> logger, IWebHostEnvironment env)
    {
        _http = http;
        _logger = logger;
        _env = env;
    }

    private const string PythonBaseUrl = "http://127.0.0.1:5091";

    public async Task<PythonRenderResponse?> RenderAsync(int templateId)
    {
        var request = new PythonRenderRequest
        {
            TemplateId = templateId,
            OutputDir = Path.Combine(_env.ContentRootPath, "Preview")
        };
        try
        {
            var response = await _http.PostAsJsonAsync($"{PythonBaseUrl}/render/runtime", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<PythonRenderResponse>();
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Python render service call failed for template {TemplateId}", templateId);
            return null;
        }
    }

    /// <summary>
    /// Upload an Excel file to the Python preview endpoint.
    /// Python runs MakeCluster → ExportPdf → pixel scan to generate field ratios,
    /// and renders the background PNG directly to the given output directory.
    /// </summary>
    /// <param name="xlsxPath">Path to the uploaded Excel file on disk.</param>
    /// <param name="outputDir">Directory where Python saves the background PNG.</param>
    /// <returns>Preview response with background image filename, page size, and field data.</returns>
    public async Task<PythonPreviewResponse?> UploadPreviewAsync(string xlsxPath, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        using var formData = new MultipartFormDataContent();
        var fileBytes = await System.IO.File.ReadAllBytesAsync(xlsxPath);
        var fileContent = new ByteArrayContent(fileBytes);
        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        formData.Add(fileContent, "file", Path.GetFileName(xlsxPath));

        string url = $"{PythonBaseUrl}/upload/preview?output_dir={Uri.EscapeDataString(outputDir)}";

        try
        {
            var response = await _http.PostAsync(url, formData);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<PythonPreviewResponse>();
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Python preview call failed");
            return null;
        }
    }
}
