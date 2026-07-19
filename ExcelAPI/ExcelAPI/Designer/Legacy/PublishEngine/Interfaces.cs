using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.PublishEngine;

public interface IXmlGenerator
{
    string Generate(WorkbookModel workbook, List<ClusterModel> clusters,
        Dictionary<(int sheetIndex, int clusterId), CoordinateRect> finalCoords);
}

public interface IBackgroundExporter
{
    Task<string?> ExportAsync(string excelFilePath, int sheetIndex, string outputDir);
}

public interface IPublishEngine
{
    string TransformName { get; }
    Task<PublishResult> PublishAsync(string excelFilePath, string? outputDir = null);
}
