using ExcelAPI.LegacyEngine;
using ExcelAPI.Models;
using ExcelAPI.Services;
using ExcelAPI.Services.Interfaces;
using ExcelAPI.Generators;
using ExcelAPI.Rendering;
using ExcelAPI.Runtime;

var builder = WebApplication.CreateBuilder(args);

// --- Add services to the container ---

builder.Services.AddControllers();

// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

// Register configuration via Options pattern
builder.Services.Configure<ExcelCaptureOptions>(
    builder.Configuration.GetSection(ExcelCaptureOptions.SectionName));

// Register application services
builder.Services.AddScoped<IExcelCaptureService, ExcelCaptureService>();
builder.Services.AddScoped<ExcelCaptureService>();
builder.Services.AddScoped<IFormSaveService, FormSaveService>();
builder.Services.AddScoped<XmlGenerator>();
builder.Services.AddScoped<DatabaseGenerator>();
builder.Services.AddScoped<WorkbookGenerator>();
builder.Services.AddScoped<PreviewGenerator>();
builder.Services.AddScoped<PdfGenerator>();

// Register Phase 19 Legacy Engine (rebuild of original PaperLess publishing pipeline)
builder.Services.AddLegacyEngine(debugEnabled: true);

// Register FormLess rendering core
builder.Services.AddSingleton<OpenXmlParser>();
builder.Services.AddSingleton<GeometryBuilder>();
builder.Services.AddSingleton<PageGeometryResolver>();
builder.Services.AddSingleton<MarginResolver>();
builder.Services.AddSingleton<ScalingResolver>();
builder.Services.AddSingleton<PrintLayoutEngine>();
builder.Services.AddSingleton<CoordinateEngine>();
builder.Services.AddSingleton<CellGeometryEngine>();
builder.Services.AddSingleton<FillEngine>();
builder.Services.AddSingleton<GridlineLayer>();
builder.Services.AddSingleton<BorderEngine>();

// Register IRenderLayer implementations (order = render order)
builder.Services.AddSingleton<IRenderLayer, FillEngine>();
builder.Services.AddSingleton<IRenderLayer, GridlineLayer>();
builder.Services.AddSingleton<IRenderLayer, BorderEngine>();
builder.Services.AddSingleton<IRenderLayer, TextEngine>();

// Register Text Rendering Engine (Phase 11C / M4)
builder.Services.AddSingleton<FontResolver>();
builder.Services.AddSingleton<TextLayoutEngine>();
builder.Services.AddSingleton<TextEngine>();

// Register Style Resolution Engine (Phase 11E)
builder.Services.AddSingleton<ThemeResolver>();
builder.Services.AddSingleton<ColorResolver>();
builder.Services.AddSingleton<StyleCache>();
builder.Services.AddSingleton<StyleResolver>();

// Register Image & Shape Engine (Phase 11F)
builder.Services.AddSingleton<DrawingParser>();
builder.Services.AddSingleton<ImageResolver>();
builder.Services.AddSingleton<ImageEngine>();
builder.Services.AddSingleton<ShapeResolver>();
builder.Services.AddSingleton<ShapeEngine>();

// Register IRenderLayer for images and shapes (Layer 5 = images, Layer 6 = shapes)
builder.Services.AddSingleton<IRenderLayer, ImageEngine>();
builder.Services.AddSingleton<IRenderLayer, ShapeEngine>();

// Register Production Export Engine (Phase 11G)
builder.Services.AddSingleton<ExportOptions>();
builder.Services.AddSingleton<PageRenderer>();
builder.Services.AddSingleton<ExportCoordinator>();

// Register Form Runtime Engine (Phase 11I)
builder.Services.AddSingleton<FieldTypeResolver>();
builder.Services.AddSingleton<FieldDetector>();
builder.Services.AddSingleton<FormRuntimeBuilder>();
builder.Services.AddSingleton<RuntimeSerializer>();

// Register CoordinateTransformer for printed page origin calculation (Phase 36)
builder.Services.AddSingleton<CoordinateTransformer>();

// Register COM Runtime Coordinate Generator (Phase 12)
// Persists and loads Excel COM field rectangles as the single source of truth.
// Eliminates OpenXML coordinate recalculation on every Runtime GET request.
builder.Services.AddSingleton<RuntimeCoordinateGenerator>();

// Register Python Rendering Service (Phase 47)
builder.Services.AddHttpClient<PythonRenderService>(client =>
{
    client.Timeout = TimeSpan.FromMinutes(5);
});
builder.Services.AddSingleton<PythonRenderService>();

// RendererCoordinator depends on IEnumerable<IRenderLayer>
builder.Services.AddSingleton<RendererCoordinator>();

// Register background cleanup service
builder.Services.AddHostedService<PreviewCleanupService>();

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

app.Logger.LogInformation(
    "PaperLess Excel Capture API started. Version: 1.3, " +
    "Preview directory: {PreviewDir}, Max upload: {MaxUpload}MB",
    previewPath,
    excelCaptureOptions.MaxUploadSizeMB);

app.Run();


