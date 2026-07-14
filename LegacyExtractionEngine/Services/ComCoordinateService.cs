using System.Runtime.InteropServices;
using System.Text.Json;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Uses Excel COM Interop (late-binding) to read rendered cell positions
/// that exactly match what the legacy VB6 ConMas importer used.
/// 
/// The legacy system relied on Excel's own layout engine via Range.Left/Top/Width/Height,
/// NOT on OpenXML column-width estimates. This service replicates that behavior.
/// </summary>
public class ComCoordinateService : IDisposable
{
    private readonly IProgress<string> _progress;

    private dynamic? _excel;
    private dynamic? _workbook;
    private bool _disposed;

    public const double DefaultPageWidth = 612.0;
    public const double DefaultPageHeight = 792.0;

    public ComCoordinateService(IProgress<string> progress)
    {
        _progress = progress;
    }

    /// <summary>
    /// Open the workbook via Excel COM and extract all coordinate data needed
    /// to compute legacy-compatible cluster ratios.
    /// 
    /// The PrintedOrigin is derived from the FIRST DB cluster's data using:
    ///   PrintedOriginX = DB_Ratio * PageWidth - COM_Left
    ///   PrintedOriginY = DB_Ratio * PageHeight - COM_Top
    /// This is the proven formula from the coordinate investigation.
    /// </summary>
    /// <param name="workbookPath">Path to the .xlsx file to open via COM.</param>
    /// <param name="dbClusters">DB cluster definitions with cell addresses AND stored ratios.</param>
    /// <param name="openXmlPrintArea">
    /// Optional PrintArea from OpenXML defined names (more reliable than COM PageSetup
    /// which may return printer defaults instead of workbook settings).
    /// </param>
    public ComExtraction Extract(string workbookPath, List<DefCluster> dbClusters,
        string? openXmlPrintArea = null)
    {
        _progress.Report("  Opening Excel via COM...");
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new InvalidOperationException("Excel COM is not available on this system.");

        _excel = Activator.CreateInstance(excelType);
        _excel.Visible = false;
        _excel.DisplayAlerts = false;
        _excel.ScreenUpdating = false;

        _progress.Report($"  Opening workbook: {Path.GetFileName(workbookPath)}");
        _workbook = _excel.Workbooks.Open(workbookPath);

        var result = new ComExtraction();

        // First sheet (legacy only uses sheet 1)
        dynamic ws = _workbook.Sheets[1];
        dynamic pageSetup = ws.PageSetup;

        // === Page Setup ===
        result.PageWidth = SafeReadDouble(() => (double)pageSetup.PageWidth, DefaultPageWidth);
        result.PageHeight = SafeReadDouble(() => (double)pageSetup.PageHeight, DefaultPageHeight);
        result.LeftMargin = SafeReadDouble(() => (double)pageSetup.LeftMargin, 0);
        result.RightMargin = SafeReadDouble(() => (double)pageSetup.RightMargin, 0);
        result.TopMargin = SafeReadDouble(() => (double)pageSetup.TopMargin, 0);
        result.BottomMargin = SafeReadDouble(() => (double)pageSetup.BottomMargin, 0);
        result.CenterHorizontally = SafeReadBool(() => (bool)pageSetup.CenterHorizontally);
        result.CenterVertically = SafeReadBool(() => (bool)pageSetup.CenterVertically);
        result.PrintArea = SafeReadString(() => (string)pageSetup.PrintArea);
        result.Zoom = SafeReadInt(() => (int)pageSetup.Zoom);
        result.FitToPagesWide = SafeReadInt(() => (int)pageSetup.FitToPagesWide);
        result.FitToPagesTall = SafeReadInt(() => (int)pageSetup.FitToPagesTall);
        result.PaperSize = SafeReadInt(() => (int)pageSetup.PaperSize);
        result.Orientation = SafeReadInt(() => (int)pageSetup.Orientation);

        // Fallback: if COM returns 0 for page dimensions (no printer), use defaults
        if (result.PageWidth < 1) result.PageWidth = DefaultPageWidth;
        if (result.PageHeight < 1) result.PageHeight = DefaultPageHeight;

        _progress.Report($"  Page: {result.PageWidth:F1}x{result.PageHeight:F1}, " +
            $"Margins(L,R,T,B): {result.LeftMargin:F2}, {result.RightMargin:F2}, " +
            $"{result.TopMargin:F2}, {result.BottomMargin:F2}, " +
            $"Center(H,V): {result.CenterHorizontally},{result.CenterVertically}");

        // === Read Range positions for each DB cluster address ===
        _progress.Report("  Reading COM positions for each DB cluster...");
        double firstComLeft = 0, firstComTop = 0;
        string? firstAddr = null;
        double firstDbLeft = 0, firstDbTop = 0;

        foreach (var dbCluster in dbClusters)
        {
            var addr = dbCluster.CellAddr?.Trim();
            if (string.IsNullOrEmpty(addr)) continue;

            try
            {
                dynamic range = ws.Range[addr];
                var cp = new ComPosition
                {
                    Address = addr,
                    Left = (double)range.Left,
                    Top = (double)range.Top,
                    Width = (double)range.Width,
                    Height = (double)range.Height
                };

                // Also try MergeArea for merged cells
                try
                {
                    dynamic merge = range.MergeArea;
                    if (merge != null && !string.IsNullOrEmpty((string)merge.Address) &&
                        (string)merge.Address != addr)
                    {
                        cp.MergeAddress = (string)merge.Address;
                        cp.MergeLeft = (double)merge.Left;
                        cp.MergeTop = (double)merge.Top;
                        cp.MergeWidth = (double)merge.Width;
                        cp.MergeHeight = (double)merge.Height;
                    }
                }
                catch { }

                result.ComPositions[addr] = cp;
                _progress.Report($"    {addr}: Left={cp.Left:F2}, Top={cp.Top:F2}, " +
                    $"Width={cp.Width:F2}, Height={cp.Height:F2}" +
                    (cp.MergeAddress != null ? $" (merge: {cp.MergeAddress})" : ""));

                // Save first cluster's data for PrintedOrigin derivation
                if (firstAddr == null)
                {
                    firstAddr = addr;
                    firstComLeft = cp.Left;
                    firstComTop = cp.Top;
                    firstDbLeft = ParseDouble(dbCluster.LeftPosition);
                    firstDbTop = ParseDouble(dbCluster.TopPosition);
                }
            }
            catch (Exception ex)
            {
                _progress.Report($"    Warning: Could not read {addr}: {ex.Message}");
            }
        }

        _progress.Report($"  Extracted {result.ComPositions.Count} cluster positions");

        // === Compute PrintedOrigin from first DB cluster ===
        // PROVEN FORMULA: PrintedOriginX = DB_Left_Ratio * PageWidth - COM_Left
        // This is derived from: DB_Ratio = (COM_Left + PrintedOriginX) / PageWidth
        // => PrintedOriginX = DB_Ratio * PageWidth - COM_Left
        if (firstAddr != null)
        {
            result.PrintedOriginX = RoundEx(firstDbLeft * result.PageWidth - firstComLeft);
            result.PrintedOriginY = RoundEx(firstDbTop * result.PageHeight - firstComTop);

            // Also derive gap for Right/Bottom from the first cluster
            // Gap_x = DB_Right * PageWidth - (COM_Left + COM_Width + PrintedOriginX)
            // Gap_y = DB_Bottom * PageHeight - (COM_Top + COM_Height + PrintedOriginY)
            if (result.ComPositions.TryGetValue(firstAddr, out var firstPos))
            {
                double dbRight = ParseDouble(dbClusters.First(c => c.CellAddr?.Trim() == firstAddr)?.RightPosition);
                double dbBottom = ParseDouble(dbClusters.First(c => c.CellAddr?.Trim() == firstAddr)?.BottomPosition);
                result.ClusterGapX = RoundEx(dbRight * result.PageWidth - (firstPos.Left + firstPos.Width + result.PrintedOriginX));
                result.ClusterGapY = RoundEx(dbBottom * result.PageHeight - (firstPos.Top + firstPos.Height + result.PrintedOriginY));
            }

            _progress.Report($"  PrintedOrigin derived from first cluster {firstAddr}: " +
                $"DB_L/R={firstDbLeft:F7}/{firstDbTop:F7}, " +
                $"COM_L/T={firstComLeft:F2}/{firstComTop:F2}, " +
                $"Origin=({result.PrintedOriginX:F2},{result.PrintedOriginY:F2}), " +
                $"Gap=({result.ClusterGapX:F2},{result.ClusterGapY:F2})");
        }
        else
        {
            // Fallback: compute from page setup
            result.PrintedOriginX = result.CenterHorizontally
                ? (result.PageWidth - (result.PageWidth - result.LeftMargin - result.RightMargin)) / 2.0
                : result.LeftMargin;
            result.PrintedOriginY = result.CenterVertically
                ? (result.PageHeight - (result.PageHeight - result.TopMargin - result.BottomMargin)) / 2.0
                : result.TopMargin;
            _progress.Report($"  WARNING: Could not derive PrintedOrigin from DB data. " +
                $"Fallback from margins: ({result.PrintedOriginX:F2},{result.PrintedOriginY:F2})");
        }

        return result;
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_workbook != null)
            {
                _workbook.Close(false);
                Marshal.ReleaseComObject(_workbook);
            }
            if (_excel != null)
            {
                _excel.Quit();
                Marshal.ReleaseComObject(_excel);
            }
        }
        catch { /* cleanup best-effort */ }
    }

    private static double SafeReadDouble(Func<double> getter, double fallback = 0)
    {
        try { return getter(); }
        catch { return fallback; }
    }

    private static bool SafeReadBool(Func<bool> getter)
    {
        try { return getter(); }
        catch { return false; }
    }

    private static string SafeReadString(Func<string> getter)
    {
        try { return getter() ?? ""; }
        catch { return ""; }
    }

    private static int SafeReadInt(Func<int> getter)
    {
        try { return getter(); }
        catch { return 0; }
    }

    /// <summary>
    /// Clean up print area strings from OpenXML format to COM-valid format.
    /// OpenXML may return format like ${B}$1:${M}$44 (with curly braces),
    /// but COM expects $B$1:$M$44 (standard absolute reference).
    /// </summary>
    private static string? CleanPrintArea(string? printArea)
    {
        if (string.IsNullOrEmpty(printArea)) return null;
        // Remove curly braces that OpenXML may insert
        var cleaned = printArea.Replace("{", "").Replace("}", "");
        // Ensure at least one $ is present for absolute reference
        if (cleaned.Length > 0 && !cleaned.Contains("$"))
        {
            // Add $ format: A1:B2 -> $A$1:$B$2
            var parts = cleaned.Split(':');
            if (parts.Length == 2)
            {
                cleaned = $"${parts[0]}:${parts[1]}";
            }
        }
        return cleaned;
    }

    /// <summary>
    /// Round exactly as the legacy VB6 RoundEx function does:
    /// Math.Round((float)value, 7, MidpointRounding.AwayFromZero)
    /// </summary>
    public static double RoundEx(double value)
    {
        return Math.Round((float)value, 7, MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// Compute the legacy cluster ratio from a COM position.
    /// Formula: RoundEx((position + printedOrigin) / pageDimension, 7)
    /// </summary>
    public static double ComputeRatio(double comPos, double printedOrigin, double pageDim)
    {
        if (pageDim <= 0) return 0;
        return RoundEx((comPos + printedOrigin) / pageDim);
    }
}

/// <summary>
/// All coordinate data extracted from Excel COM for one workbook session.
/// </summary>
public class ComExtraction
{
    // Page setup
    public double PageWidth { get; set; }
    public double PageHeight { get; set; }
    public double LeftMargin { get; set; }
    public double RightMargin { get; set; }
    public double TopMargin { get; set; }
    public double BottomMargin { get; set; }
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public string? PrintArea { get; set; }
    public int? Zoom { get; set; }
    public int? FitToPagesWide { get; set; }
    public int? FitToPagesTall { get; set; }
    public int? PaperSize { get; set; }
    public int? Orientation { get; set; }

    // Print area dimensions
    public double PrintAreaWidth { get; set; }
    public double PrintAreaHeight { get; set; }

    // UsedRange dimensions (COM)
    public double UsedRangeLeft { get; set; }
    public double UsedRangeTop { get; set; }
    public double UsedRangeWidth { get; set; }
    public double UsedRangeHeight { get; set; }

    // Derived content dimensions
    public double ContentWidth { get; set; }
    public double ContentHeight { get; set; }

    // Printed origin offset (derived from content + page dimensions)
    public double PrintedOriginX { get; set; }
    public double PrintedOriginY { get; set; }

    // Cluster gap for Right/Bottom (derived from first cluster's DB data)
    public double ClusterGapX { get; set; }
    public double ClusterGapY { get; set; }

    // Per-cluster COM positions
    public Dictionary<string, ComPosition> ComPositions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>
/// COM position data for one cell range.
/// </summary>
public class ComPosition
{
    public string Address { get; set; } = "";
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }

    // Merge area (if cell is part of a merge)
    public string? MergeAddress { get; set; }
    public double MergeLeft { get; set; }
    public double MergeTop { get; set; }
    public double MergeWidth { get; set; }
    public double MergeHeight { get; set; }
}
