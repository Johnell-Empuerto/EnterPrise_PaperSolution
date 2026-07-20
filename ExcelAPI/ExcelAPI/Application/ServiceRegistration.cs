using ExcelAPI.Application;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Registers all Application-layer services.
/// These services are COM-free shared utilities used by both Designer and Runtime.
/// </summary>
public static class ApplicationServiceRegistration
{
    public static IServiceCollection AddApplication(this IServiceCollection services)
    {
        // Form save pipeline
        services.AddScoped<IFormSaveService, FormSaveService>();

        // Generators (COM-free)
        services.AddScoped<XmlGenerator>();
        services.AddScoped<DatabaseGenerator>();

        // Coordinate transform for printed page origin (Phase 36)
        services.AddSingleton<CoordinateTransformer>();

        // Phase 4.4: Workbook editing
        services.AddScoped<WorkbookValueWriter>();

        // Phase 17/18: Lightweight functional compatibility validator (replaces WorkbookDiffValidator)
        // Validates only business-critical requirements — never rejects the workbook.
        // Matches ConMas behavior which never performs structural validation.
        services.AddScoped<CompatibilityValidator>();

        // NOTE: WorkbookDiffValidator (Phase 4.4) has been removed from registration.
        // Structural byte-for-byte validation is no longer performed.
        // The source file is kept on disk for reference only.

        // Phase 5.2: Server-side session storage (TempWorkbooks/{sessionId}/original.xlsx)
        // The browser only tracks sessionId — the server owns the workbook.
        services.AddSingleton<ISessionWorkbookStore, SessionWorkbookStore>();

        // Background cleanup service
        services.AddHostedService<PreviewCleanupService>();

        return services;
    }
}
