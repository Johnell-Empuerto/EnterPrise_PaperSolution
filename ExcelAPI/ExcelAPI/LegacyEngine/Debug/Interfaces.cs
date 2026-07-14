using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.Debug;

public interface IDebugSnapshotter
{
    bool Enabled { get; }
    void CaptureWorkbook(WorkbookModel workbook);
    void CaptureSheets(WorkbookModel workbook);
    void CaptureColumns(SheetModel sheet);
    void CaptureRows(SheetModel sheet);
    void CaptureClusters(List<ClusterModel> clusters);
    void CaptureCoordinatesBefore(ClusterModel cluster);
    void CaptureCoordinatesAfter(ClusterModel cluster);
    void CaptureXml(string xml);
    void CaptureBackground(string? backgroundPath);
    void CaptureVerification(string verificationJson);
    void Flush(string outputDir);
}
