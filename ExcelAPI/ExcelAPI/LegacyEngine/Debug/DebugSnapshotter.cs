using System.Text.Json;
using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.Debug;

public class DebugSnapshotter : IDebugSnapshotter
{
    public bool Enabled { get; }

    private readonly DebugSnapshot _snapshot = new();
    private WorkbookModel? _workbook;
    private List<ClusterModel>? _clusters;
    private readonly Dictionary<int, CoordinateRect> _coordsBefore = new();
    private readonly Dictionary<int, CoordinateRect> _coordsAfter = new();
    private string? _xml;
    private string? _backgroundPath;
    private string? _verificationJson;
    private readonly List<object> _columnsData = new();
    private readonly List<object> _rowsData = new();

    public DebugSnapshotter(bool enabled)
    {
        Enabled = enabled;
    }

    public void CaptureWorkbook(WorkbookModel workbook)
    {
        if (!Enabled) return;
        _workbook = workbook;
    }

    public void CaptureSheets(WorkbookModel workbook)
    {
        if (!Enabled) return;
        _snapshot.SheetsPath = "sheets.json";
    }

    public void CaptureColumns(SheetModel sheet)
    {
        if (!Enabled) return;
        _columnsData.AddRange(sheet.Columns.Select(c => new
        {
            sheet.SheetIndex,
            sheet.Name,
            c.Index,
            c.WidthPoints
        }));
        _snapshot.ColumnsPath = "columns.json";
    }

    public void CaptureRows(SheetModel sheet)
    {
        if (!Enabled) return;
        _rowsData.AddRange(sheet.Rows.Select(r => new
        {
            sheet.SheetIndex,
            sheet.Name,
            r.Index,
            r.HeightPoints
        }));
        _snapshot.RowsPath = "rows.json";
    }

    public void CaptureClusters(List<ClusterModel> clusters)
    {
        if (!Enabled) return;
        _clusters = clusters;
        _snapshot.ClustersPath = "clusters.json";
    }

    public void CaptureCoordinatesBefore(ClusterModel cluster)
    {
        if (!Enabled || cluster.CoordinatesBeforeTransform == null) return;
        _coordsBefore[cluster.ClusterId] = cluster.CoordinatesBeforeTransform;
        _snapshot.CoordinatesBeforePath = "coordinates_before_transform.json";
    }

    public void CaptureCoordinatesAfter(ClusterModel cluster)
    {
        if (!Enabled || cluster.CoordinatesAfterTransform == null) return;
        _coordsAfter[cluster.ClusterId] = cluster.CoordinatesAfterTransform;
        _snapshot.CoordinatesAfterPath = "coordinates_after_transform.json";
    }

    public void CaptureXml(string xml)
    {
        if (!Enabled) return;
        _xml = xml;
        _snapshot.XmlPath = "xml_generated.xml";
    }

    public void CaptureBackground(string? backgroundPath)
    {
        if (!Enabled) return;
        _backgroundPath = backgroundPath;
        _snapshot.BackgroundPath = backgroundPath != null
            ? Path.GetFileName(backgroundPath) : null;
    }

    public void CaptureVerification(string verificationJson)
    {
        if (!Enabled) return;
        _verificationJson = verificationJson;
        _snapshot.VerificationPath = "verification.json";
    }

    public void Flush(string outputDir)
    {
        if (!Enabled) return;

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        string debugDir = Path.Combine(outputDir, "debug");
        Directory.CreateDirectory(debugDir);

        if (_snapshot.SheetsPath != null && _workbook != null)
        {
            var sheetsData = _workbook.Sheets.Select(s => new
            {
                s.Name,
                s.SheetIndex,
                PrintArea = s.PrintArea.Address,
                PageSetup = new
                {
                    s.PageSetup.PageWidthPoints,
                    s.PageSetup.PageHeightPoints,
                    s.PageSetup.MarginLeft,
                    s.PageSetup.MarginRight,
                    s.PageSetup.MarginTop,
                    s.PageSetup.MarginBottom,
                    s.PageSetup.CenterHorizontally,
                    s.PageSetup.CenterVertically,
                    s.PageSetup.Zoom,
                    s.PageSetup.Orientation
                },
                ColumnCount = s.Columns.Count,
                RowCount = s.Rows.Count,
                CellCount = s.Cells.Count,
                s.ContentWidthPoints,
                s.ContentHeightPoints
            }).ToList();
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.SheetsPath),
                JsonSerializer.Serialize(sheetsData, options));
        }

        if (_snapshot.ColumnsPath != null && _columnsData.Count > 0)
        {
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.ColumnsPath),
                JsonSerializer.Serialize(_columnsData, options));
        }

        if (_snapshot.RowsPath != null && _rowsData.Count > 0)
        {
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.RowsPath),
                JsonSerializer.Serialize(_rowsData, options));
        }

        if (_snapshot.ClustersPath != null && _clusters != null)
        {
            var clusterData = _clusters.Select(c => new
            {
                c.ClusterId,
                c.SheetIndex,
                c.StartRow,
                c.EndRow,
                c.StartColumn,
                c.EndColumn,
                c.CellAddress,
                c.WidthPoints,
                c.HeightPoints,
                c.LeftPoints,
                c.TopPoints,
                c.AreaPercent,
                c.AspectRatio
            }).ToList();
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.ClustersPath),
                JsonSerializer.Serialize(clusterData, options));
        }

        if (_snapshot.CoordinatesBeforePath != null)
        {
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.CoordinatesBeforePath),
                JsonSerializer.Serialize(_coordsBefore, options));
        }

        if (_snapshot.CoordinatesAfterPath != null)
        {
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.CoordinatesAfterPath),
                JsonSerializer.Serialize(_coordsAfter, options));
        }

        if (_snapshot.XmlPath != null && _xml != null)
        {
            File.WriteAllText(Path.Combine(debugDir, _snapshot.XmlPath), _xml);
        }

        if (_snapshot.VerificationPath != null && _verificationJson != null)
        {
            File.WriteAllText(
                Path.Combine(debugDir, _snapshot.VerificationPath), _verificationJson);
        }

        _snapshot.SaveTo(debugDir);
    }
}
