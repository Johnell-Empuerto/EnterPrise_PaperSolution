using ExcelAPI.Models;
using ExcelAPI.Services;
using ExcelAPI.Services.Interfaces;

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
