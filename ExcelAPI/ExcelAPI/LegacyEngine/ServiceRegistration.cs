using ExcelAPI.LegacyEngine.ClusterEngine;
using ExcelAPI.LegacyEngine.CoordinateEngine;
using ExcelAPI.LegacyEngine.Debug;
using ExcelAPI.LegacyEngine.ExcelEngine;
using ExcelAPI.LegacyEngine.LayoutEngine;
using ExcelAPI.LegacyEngine.Models;
using ExcelAPI.LegacyEngine.OpenXml;
using ExcelAPI.LegacyEngine.PublishEngine;
using ExcelAPI.LegacyEngine.VerificationEngine;

namespace ExcelAPI.LegacyEngine;

public static class ServiceRegistration
{
    public static IServiceCollection AddLegacyEngine(this IServiceCollection services, bool debugEnabled = false, bool useOpenXml = true)
    {
        // Excel Engine — OpenXML (no COM) or COM-based
        if (useOpenXml)
        {
            // OpenXML readers
            services.AddScoped<IStyleReader, StyleReader>();
            services.AddScoped<IPageSetupReader, PageSetupReader>();
            services.AddScoped<IPrintAreaReader, PrintAreaReader>();
            services.AddScoped<IColumnWidthReader, ColumnWidthReader>();
            services.AddScoped<IRowHeightReader, RowHeightReader>();
            services.AddScoped<IMergeCellReader, MergeCellReader>();
            services.AddScoped<ICommentReader, CommentReader>();
            services.AddScoped<IImageReader, ImageReader>();
            services.AddScoped<ICellDataReader, CellDataReader>();
            services.AddScoped<ICellStyleReader, CellStyleReader>();

            // OpenXML loader chain
            services.AddScoped<IOpenXmlWorksheetLoader, OpenXmlWorksheetLoader>();
            services.AddScoped<IOpenXmlWorkbookLoader, OpenXmlWorkbookLoader>();
            services.AddScoped<IWorkbookLoader, OpenXmlWorkbookLoaderAdapter>();

            // Dummy/stub for COM-based loaders (not used)
            services.AddScoped<IWorksheetLoader>(_ =>
                throw new InvalidOperationException("COM-based WorksheetLoader is not registered when useOpenXml=true"));
            services.AddScoped<IPrintAreaLoader>(_ =>
                throw new InvalidOperationException("COM-based PrintAreaLoader is not registered when useOpenXml=true"));
        }
        else
        {
            // COM-based loaders (legacy, requires Excel installation)
            services.AddScoped<IWorkbookLoader, WorkbookLoader>();
            services.AddScoped<IWorksheetLoader, WorksheetLoader>();
            services.AddScoped<IPrintAreaLoader, PrintAreaLoader>();
        }

        // Layout Engine
        services.AddScoped<IColumnEngine, ColumnEngine>();
        services.AddScoped<IRowEngine, RowEngine>();
        services.AddScoped<IOriginEngine, OriginEngine>();
        services.AddScoped<IPageEngine, PageEngine>();

        // Cluster Engine
        services.AddScoped<IClusterBuilder, ClusterBuilder>();
        services.AddScoped<IClusterDetector, ClusterDetector>();

        // Coordinate Engine
        services.AddScoped<ICoordinateCalculator, CoordinateCalculator>();
        services.AddScoped<ILegacyCoordinateTransform, LegacyCoordinateTransform>();

        // Publish Engine
        services.AddScoped<IXmlGenerator, XmlGenerator>();
        services.AddScoped<IBackgroundExporter, BackgroundExporter>();
        services.AddScoped<IPublishEngine, PublishEngine.PublishEngine>();

        // Verification Engine
        services.AddScoped<IVerificationDatabaseReader, VerificationDatabaseReader>();
        services.AddScoped<IXmlComparer, XmlComparer>();
        services.AddScoped<IImageComparer, ImageComparer>();
        services.AddScoped<IVerificationReportGenerator, VerificationReportGenerator>();
        services.AddScoped<IVerificationEngine, VerificationEngine.VerificationEngine>();

        // Debug Snapshotter (conditionally enabled)
        if (debugEnabled)
            services.AddScoped<IDebugSnapshotter>(_ => new DebugSnapshotter(true));
        else
            services.AddScoped<IDebugSnapshotter>(_ => new DebugSnapshotter(false));

        return services;
    }
}
