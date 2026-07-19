using ExcelAPI.Runtime;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Registers all Runtime-layer services.
/// These services are COM-free and build runtime forms from captured Excel data.
/// </summary>
public static class RuntimeServiceRegistration
{
    public static IServiceCollection AddRuntime(this IServiceCollection services)
    {
        // Field detection and type resolution
        services.AddSingleton<FieldTypeResolver>();
        services.AddSingleton<FieldDetector>();

        // Runtime form builder and serializer
        services.AddSingleton<FormRuntimeBuilder>();
        services.AddSingleton<RuntimeSerializer>();

        // COM Runtime Coordinate Generator (Phase 12)
        // Persists and loads Excel COM field rectangles as the single source of truth.
        services.AddSingleton<RuntimeCoordinateGenerator>();

        return services;
    }
}
