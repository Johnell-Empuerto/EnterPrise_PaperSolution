using ExcelAPI.Designer.Analysis;
using ExcelAPI.Designer.Capture;
using ExcelAPI.Designer.Generation;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Registers all Designer-layer services (Capture + Analysis + Generation).
/// These services require Excel COM Interop.
/// See also: AddLegacyEngine() for Legacy engine registrations.
/// </summary>
public static class DesignerServiceRegistration
{
    public static IServiceCollection AddDesigner(this IServiceCollection services)
    {
        // Capture services
        services.AddScoped<IExcelCaptureService, ExcelCaptureService>();
        services.AddScoped<ExcelCaptureService>();

        // Analysis services (workbook reading, hybrid COM+OpenXml)
        services.AddScoped<WorkbookReaderService>();

        // Generation services (workbook output, PDF, preview)
        services.AddScoped<WorkbookGenerator>();
        services.AddScoped<PdfGenerator>();
        services.AddScoped<PreviewGenerator>();

        return services;
    }
}
