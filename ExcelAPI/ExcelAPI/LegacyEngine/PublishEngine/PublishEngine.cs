using ExcelAPI.LegacyEngine.ClusterEngine;
using ExcelAPI.LegacyEngine.CoordinateEngine;
using ExcelAPI.LegacyEngine.Debug;
using ExcelAPI.LegacyEngine.ExcelEngine;
using ExcelAPI.LegacyEngine.LayoutEngine;
using ExcelAPI.LegacyEngine.Models;
using ExcelAPI.LegacyEngine.VerificationEngine;

namespace ExcelAPI.LegacyEngine.PublishEngine;

public class PublishEngine : IPublishEngine
{
    private readonly IWorkbookLoader _workbookLoader;
    private readonly IClusterBuilder _clusterBuilder;
    private readonly ICoordinateCalculator _coordCalculator;
    private readonly ILegacyCoordinateTransform _coordinateTransform;
    private readonly IOriginEngine _originEngine;
    private readonly IXmlGenerator _xmlGenerator;
    private readonly IBackgroundExporter _backgroundExporter;
    private readonly IDebugSnapshotter? _debug;

    public string TransformName => _coordinateTransform.Name;

    public PublishEngine(
        IWorkbookLoader workbookLoader,
        IClusterBuilder clusterBuilder,
        ICoordinateCalculator coordCalculator,
        ILegacyCoordinateTransform coordinateTransform,
        IOriginEngine originEngine,
        IXmlGenerator xmlGenerator,
        IBackgroundExporter backgroundExporter,
        IDebugSnapshotter? debug = null)
    {
        _workbookLoader = workbookLoader;
        _clusterBuilder = clusterBuilder;
        _coordCalculator = coordCalculator;
        _coordinateTransform = coordinateTransform;
        _originEngine = originEngine;
        _xmlGenerator = xmlGenerator;
        _backgroundExporter = backgroundExporter;
        _debug = debug;
    }

    public async Task<PublishResult> PublishAsync(string excelFilePath, string? outputDir = null)
    {
        var result = new PublishResult();

        try
        {
            outputDir ??= Path.Combine(Path.GetTempPath(), "PaperLessPublish",
                Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(outputDir);

            // Step 1: Load workbook
            var workbook = _workbookLoader.Load(excelFilePath);
            _debug?.CaptureWorkbook(workbook);
            _debug?.CaptureSheets(workbook);

            // Step 2: Detect clusters and compute coordinates per sheet
            var allClusters = new List<ClusterModel>();
            var finalCoords = new Dictionary<(int sheetIndex, int clusterId), CoordinateRect>();

            foreach (var sheet in workbook.Sheets)
            {
                _debug?.CaptureColumns(sheet);
                _debug?.CaptureRows(sheet);

                var (originX, originY) = _originEngine.CalculateOrigin(sheet);

                var clusters = _clusterBuilder.BuildAll(sheet, sheet.SheetIndex);
                allClusters.AddRange(clusters);

                _debug?.CaptureClusters(clusters);

                foreach (var cluster in clusters)
                {
                    var raw = _coordCalculator.CalculateRaw(cluster, sheet);
                    cluster.CoordinatesBeforeTransform = _coordCalculator.CalculateNormalized(
                        cluster, sheet, originX, originY);

                    _debug?.CaptureCoordinatesBefore(cluster);

                    var final = _coordinateTransform.Transform(
                        cluster.CoordinatesBeforeTransform, cluster, sheet);
                    cluster.CoordinatesAfterTransform = final;

                    _debug?.CaptureCoordinatesAfter(cluster);

                    finalCoords[(sheet.SheetIndex, cluster.ClusterId)] = final;
                }
            }

            // Step 3: Generate XML
            string xml = _xmlGenerator.Generate(workbook, allClusters, finalCoords);
            result.XmlData = xml;
            _debug?.CaptureXml(xml);

            // Step 4: Export background images
            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                var bgPath = await _backgroundExporter.ExportAsync(
                    excelFilePath, i + 1, outputDir);
                result.Sheets.Add(new PublishSheetInfo
                {
                    SheetNo = i + 1,
                    SheetName = workbook.Sheets[i].Name,
                    WidthPoints = workbook.Sheets[i].PageSetup.PageWidthPoints,
                    HeightPoints = workbook.Sheets[i].PageSetup.PageHeightPoints,
                    BackgroundImagePath = bgPath
                });
                if (i == 0) result.BackgroundImagePath = bgPath;
                _debug?.CaptureBackground(bgPath);
            }

            // Step 5: Build cluster list for result
            foreach (var kvp in finalCoords)
            {
                var cluster = allClusters.FirstOrDefault(c =>
                    c.SheetIndex == kvp.Key.sheetIndex &&
                    c.ClusterId == kvp.Key.clusterId);
                if (cluster == null) continue;

                result.Clusters.Add(new PublishCluster
                {
                    ClusterId = cluster.ClusterId,
                    SheetNo = cluster.SheetIndex,
                    CellAddress = cluster.CellAddress,
                    Left = kvp.Value.Left,
                    Top = kvp.Value.Top,
                    Right = kvp.Value.Right,
                    Bottom = kvp.Value.Bottom
                });
            }

            _debug?.Flush(outputDir);
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }
}
