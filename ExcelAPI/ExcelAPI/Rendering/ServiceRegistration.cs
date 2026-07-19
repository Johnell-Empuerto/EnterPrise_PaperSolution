using ExcelAPI.Rendering;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Registers all Rendering pipeline services.
/// These services are COM-free and handle OpenXML parsing, geometry, and SkiaSharp rendering.
/// </summary>
public static class RenderingServiceRegistration
{
    public static IServiceCollection AddRendering(this IServiceCollection services)
    {
        // Core parsers and resolvers
        services.AddSingleton<OpenXmlParser>();
        services.AddSingleton<GeometryBuilder>();
        services.AddSingleton<PageGeometryResolver>();
        services.AddSingleton<MarginResolver>();
        services.AddSingleton<ScalingResolver>();
        services.AddSingleton<PrintLayoutEngine>();
        services.AddSingleton<CoordinateEngine>();
        services.AddSingleton<CellGeometryEngine>();

        // Fill and border engines
        services.AddSingleton<FillEngine>();
        services.AddSingleton<GridlineLayer>();
        services.AddSingleton<BorderEngine>();

        // Register IRenderLayer implementations (order = render order)
        services.AddSingleton<IRenderLayer, FillEngine>();
        services.AddSingleton<IRenderLayer, GridlineLayer>();
        services.AddSingleton<IRenderLayer, BorderEngine>();
        services.AddSingleton<IRenderLayer, TextEngine>();

        // Text Rendering Engine (Phase 11C / M4)
        services.AddSingleton<FontResolver>();
        services.AddSingleton<TextLayoutEngine>();
        services.AddSingleton<TextEngine>();

        // Style Resolution Engine (Phase 11E)
        services.AddSingleton<ThemeResolver>();
        services.AddSingleton<ColorResolver>();
        services.AddSingleton<StyleCache>();
        services.AddSingleton<StyleResolver>();

        // Image & Shape Engine (Phase 11F)
        services.AddSingleton<DrawingParser>();
        services.AddSingleton<ImageResolver>();
        services.AddSingleton<ImageEngine>();
        services.AddSingleton<ShapeResolver>();
        services.AddSingleton<ShapeEngine>();

        // Register IRenderLayer for images and shapes
        services.AddSingleton<IRenderLayer, ImageEngine>();
        services.AddSingleton<IRenderLayer, ShapeEngine>();

        // Production Export Engine (Phase 11G)
        services.AddSingleton<ExportOptions>();
        services.AddSingleton<PageRenderer>();
        services.AddSingleton<ExportCoordinator>();

        // RendererCoordinator depends on IEnumerable<IRenderLayer>
        services.AddSingleton<RendererCoordinator>();

        return services;
    }
}
