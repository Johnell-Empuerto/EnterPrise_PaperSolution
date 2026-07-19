using ExcelAPI.Application;
using ExcelAPI.Designer.Legacy;
using ExcelAPI.Models;

var builder = WebApplication.CreateBuilder(args);

// --- Add services to the container ---

builder.Services.AddControllers();

// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

// Register configuration via Options pattern
builder.Services.Configure<ExcelCaptureOptions>(
    builder.Configuration.GetSection(ExcelCaptureOptions.SectionName));

// ─────────────────────────────────────────────────────
// Module registrations (organized by architectural layer)
// ─────────────────────────────────────────────────────

// Designer — Excel COM-dependent services (Capture + Generation)
builder.Services.AddDesigner();

// Legacy Engine — reverse-engineered ConMas publishing pipeline (COM-bound)
builder.Services.AddLegacyEngine(debugEnabled: true);

// Rendering — OpenXML parsing, SkiaSharp rendering pipeline (no COM)
builder.Services.AddRendering();

// Runtime — form building, field detection, serialization (no COM)
builder.Services.AddRuntime();

// Application — shared services, generators, orchestration (no COM)
builder.Services.AddApplication();

// Python Rendering Service — HTTP client to external renderer
// NOTE: Use ONLY AddHttpClient, NOT AddSingleton.
// AddSingleton would override the typed HttpClient registration from
// AddHttpClient, causing the client to use the default 100-second timeout
// instead of the configured 5-minute timeout.
builder.Services.AddHttpClient<PythonRenderService>(client =>
{
    client.Timeout = TimeSpan.FromMinutes(5);
});

// Configure CORS to allow requests from the Next.js frontend
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowFrontend", policy =>
    {
        policy.WithOrigins("http://localhost:3002")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

// --- Auto-create required directories on startup ---
var contentRoot = app.Environment.ContentRootPath;
var excelCaptureOptions = app.Services.GetRequiredService<
    Microsoft.Extensions.Options.IOptions<ExcelCaptureOptions>>().Value;

string[] directoriesToCreate = [
    Path.Combine(contentRoot, excelCaptureOptions.UploadDirectory),
    Path.Combine(contentRoot, excelCaptureOptions.PreviewDirectory),
    Path.Combine(contentRoot, "Logs")
];

foreach (var dir in directoriesToCreate)
{
    Directory.CreateDirectory(dir);
    app.Logger.LogInformation("Ensured directory exists: {Directory}", dir);
}

Directory.CreateDirectory(Path.Combine(contentRoot, "Forms"));

// --- Configure the HTTP request pipeline ---

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

// Enable CORS
app.UseCors("AllowFrontend");

// Enable serving static files from the wwwroot folder (if any)
app.UseStaticFiles();

// Serve the Preview folder as static files so images are accessible via URL.
// Example: http://localhost:5090/preview/page_abc.png
string previewPath = Path.Combine(contentRoot, excelCaptureOptions.PreviewDirectory);

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(previewPath),
    RequestPath = "/preview"
});

string formsPath = Path.Combine(contentRoot, "Forms");
app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(formsPath),
    RequestPath = "/forms"
});

// Enable authorization
app.UseAuthorization();

// Map controllers
app.MapControllers();

// Log DI configuration for PythonRenderService (verify HttpClient timeout fix)
app.Logger.LogInformation(
    "PaperLess Excel Capture API started. Version: 1.3, " +
    "Preview directory: {PreviewDir}, Max upload: {MaxUpload}MB",
    previewPath,
    excelCaptureOptions.MaxUploadSizeMB);

try
{
    var pythonService = app.Services.GetRequiredService<PythonRenderService>();
    var httpClientField = typeof(PythonRenderService).GetField("_http",
        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
    if (httpClientField != null)
    {
        var httpClient = httpClientField.GetValue(pythonService) as HttpClient;
        if (httpClient != null)
        {
            app.Logger.LogInformation(
                "PythonRenderService: DI registration = AddHttpClient (typed), " +
                "HttpClient.Timeout = {TimeoutSeconds}s, " +
                "Service URL = http://127.0.0.1:5091",
                httpClient.Timeout.TotalSeconds);
        }
    }
}
catch (Exception ex)
{
    app.Logger.LogWarning("PythonRenderService: Could not verify DI config: {Message}", ex.Message);
}

app.Run();


