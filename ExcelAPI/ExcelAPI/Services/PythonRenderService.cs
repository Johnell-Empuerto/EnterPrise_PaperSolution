using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace ExcelAPI.Services;

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

    public async Task<PythonRenderResponse?> RenderAsync(int templateId)
    {
        string pythonUrl = "http://127.0.0.1:5091";
        var request = new PythonRenderRequest
        {
            TemplateId = templateId,
            OutputDir = Path.Combine(_env.ContentRootPath, "Preview")
        };
        try
        {
            var response = await _http.PostAsJsonAsync($"{pythonUrl}/render/runtime", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<PythonRenderResponse>();
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Python render service call failed for template {TemplateId}", templateId);
            return null;
        }
    }
}
