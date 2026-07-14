using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 11 — Reverse Engineer the Legacy Coordinate Transformation.
/// 
/// For templates 546, 547, 548 only, extracts every available data point
/// and computes the mathematical relationship between COM-extracted positions
/// and DB-stored ratios.
/// 
/// No assumptions. No heuristics. Pure data-driven investigation.
/// </summary>
public class Phase11ReverseEngineer
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;
    private readonly string _tempDir;

    // Known workbook paths
    private static readonly Dictionary<int, string> WorkbookPaths = new()
    {
        { 546, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx" },
        { 547, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx" },
        { 548, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx" }
    };

    public Phase11ReverseEngineer(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test", "Phase11_LegacyReverseEngineering");
        _tempDir = Path.Combine(Path.GetTempPath(), "phase11_workbooks");
        if (Directory.Exists(_outputDir))
            Directory.Delete(_outputDir, recursive: true);
        Directory.CreateDirectory(_outputDir);
        Directory.CreateDirectory(_tempDir);
    }

    public async Task RunAsync()
    {
        var templateIds = new[] { 546, 547, 548 };

        _progress.Report("============================================");
        _progress.Report(" Phase 11 — Legacy Coordinate Reverse Engineering");
        _progress.Report($" Templates: {string.Join(", ", templateIds)}");
        _progress.Report("============================================");
        _progress.Report("");

        var report = new StringBuilder();
        report.AppendLine("# Phase 11 — Legacy Coordinate Reverse Engineering Report");
        report.AppendLine();
        report.AppendLine($"**Generated:** {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
        report.AppendLine($"**Templates:** {string.Join(", ", templateIds)}");
        report.AppendLine();

        foreach (var id in templateIds)
        {
            _progress.Report($"=== Processing Template {id} ===");
            var result = await InvestigateTemplateAsync(id);
            report.AppendLine(result);
            report.AppendLine("---");
            report.AppendLine();
        }

        // Cross-template analysis
        report.AppendLine("# Cross-Template Analysis");
        report.AppendLine();
        report.AppendLine("## Hypothesis Testing");
        report.AppendLine();
        report.AppendLine("For each possible transformation, document whether it explains the delta:");
        report.AppendLine();
        report.AppendLine("| Hypothesis | 546 | 547 | 548 | Conclusion |");
        report.AppendLine("|------------|:---:|:---:|:---:|-----------|");
        report.AppendLine("| Row-height gap accumulation | ? | ? | ? | TBD |");
        report.AppendLine("| Print-area scaling | ? | ? | ? | TBD |");
        report.AppendLine("| Page-margin offset | ? | ? | ? | TBD |");
        report.AppendLine("| Merged-cell normalization | ? | ? | ? | TBD |");
        report.AppendLine("| Font-metric column width | ? | ? | ? | TBD |");
        report.AppendLine("| Hidden row/column skip | ? | ? | ? | TBD |");
        report.AppendLine("| Printable area ratio | ? | ? | ? | TBD |");
        report.AppendLine("| Printer DPI conversion | ? | ? | ? | TBD |");
        report.AppendLine("| Proportional column scaling | ? | ? | ? | TBD |");
        report.AppendLine();

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "Phase11Report.md"), report.ToString());
        _progress.Report($"Phase 11 report written to {_outputDir}\\Phase11Report.md");
        _progress.Report("=== PHASE 11 COMPLETE ===");
    }

    private async Task<string> InvestigateTemplateAsync(int defTopId)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"## Template {defTopId}");
        sb.AppendLine();

        // Step 1: Load DB data
        _progress.Report("  Loading DB data...");
        var dbReader = new DatabaseReader(_connectionString, _progress);
        var db = await dbReader.ReadAllAsync(defTopId);
        var dbClusters = db?.DefSheets.SelectMany(s => s.Clusters).ToList() ?? new List<DefCluster>();
        sb.AppendLine($"- **DB Clusters:** {dbClusters.Count}");
        sb.AppendLine($"- **DB Top Name:** {db?.DefTop?.DefTopName ?? "(null)"}");

        // Step 2: Verify workbook
        var workbookPath = WorkbookPaths.GetValueOrDefault(defTopId);
        if (workbookPath == null || !File.Exists(workbookPath))
        {
            sb.AppendLine($"- **Workbook:** NOT FOUND at {workbookPath ?? "(no path)"}");
            return sb.ToString();
        }
        sb.AppendLine($"- **Workbook:** `{workbookPath}`");
        var fileInfo = new FileInfo(workbookPath);
        sb.AppendLine($"- **File Size:** {fileInfo.Length:N0} bytes");
        sb.AppendLine($"- **Last Modified:** {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");

        // Step 3: OpenXML import for column widths, row heights
        _progress.Report("  Reading OpenXML metadata...");
        var openXmlInfo = ReadOpenXmlInfo(workbookPath);
        sb.AppendLine();
        sb.AppendLine("### OpenXML Metadata");
        sb.AppendLine();
        sb.AppendLine($"- **Sheets:** {openXmlInfo.Sheets.Count}");
        sb.AppendLine($"- **Columns defined:** {openXmlInfo.ColumnWidths.Count}");
        sb.AppendLine($"- **Rows defined:** {openXmlInfo.RowHeights.Count}");
        sb.AppendLine($"- **Merged cells:** {openXmlInfo.MergedCells.Count}");
        sb.AppendLine($"- **PrintArea (defined name):** {openXmlInfo.PrintArea ?? "(none)"}");
        sb.AppendLine();
        sb.AppendLine("| Column | OpenXML Width (chars) | Width in Points |");
        sb.AppendLine("|--------|----------------------|-----------------|");
        foreach (var col in openXmlInfo.ColumnWidths.OrderBy(c => c.Key))
        {
            double pt = ColumnCharsToPoints(col.Value);
            sb.AppendLine($"| {col.Key} | {col.Value:F3} | {pt:F3} |");
        }
        sb.AppendLine();
        sb.AppendLine("| Row | OpenXML Height (pt) | Custom? |");
        sb.AppendLine("|-----|--------------------|---------|");
        foreach (var row in openXmlInfo.RowHeights.OrderBy(r => r.Key))
        {
            sb.AppendLine($"| {row.Key} | {row.Value.Height:F3} | {row.Value.IsCustom} |");
        }
        sb.AppendLine();

        // Step 4: COM extraction (extended)
        _progress.Report("  Extracting COM metadata (extended)...");
        var comMeta = ExtractComMetadata(workbookPath, dbClusters);
        sb.AppendLine("### COM Page Setup");
        sb.AppendLine();
        sb.AppendLine($"| Property | Value |");
        sb.AppendLine($"|----------|-------|");
        sb.AppendLine($"| Page Width | {comMeta.PageWidth:F2} pt |");
        sb.AppendLine($"| Page Height | {comMeta.PageHeight:F2} pt |");
        sb.AppendLine($"| Left Margin | {comMeta.LeftMargin:F2} pt |");
        sb.AppendLine($"| Right Margin | {comMeta.RightMargin:F2} pt |");
        sb.AppendLine($"| Top Margin | {comMeta.TopMargin:F2} pt |");
        sb.AppendLine($"| Bottom Margin | {comMeta.BottomMargin:F2} pt |");
        sb.AppendLine($"| Center Horizontally | {comMeta.CenterHorizontally} |");
        sb.AppendLine($"| Center Vertically | {comMeta.CenterVertically} |");
        sb.AppendLine($"| Zoom | {comMeta.Zoom}% |");
        sb.AppendLine($"| FitToPagesWide | {comMeta.FitToPagesWide} |");
        sb.AppendLine($"| FitToPagesTall | {comMeta.FitToPagesTall} |");
        sb.AppendLine($"| PrintArea (COM) | {comMeta.PrintArea ?? "(none)"} |");
        sb.AppendLine($"| PaperSize | {comMeta.PaperSize} |");
        sb.AppendLine($"| Orientation | {comMeta.Orientation} |");
        sb.AppendLine();
        sb.AppendLine($"| Printable Width | {comMeta.PrintableWidth:F2} pt |");
        sb.AppendLine($"| Printable Height | {comMeta.PrintableHeight:F2} pt |");
        sb.AppendLine($"| UsedRange | ({comMeta.UsedLeft:F2}, {comMeta.UsedTop:F2}) to ({comMeta.UsedRight:F2}, {comMeta.UsedBottom:F2}) |");
        sb.AppendLine($"| UsedRange Width | {comMeta.UsedWidth:F2} pt |");
        sb.AppendLine($"| UsedRange Height | {comMeta.UsedHeight:F2} pt |");
        sb.AppendLine();

        // Step 5: Per-cluster deep comparison
        _progress.Report("  Computing per-cluster analysis...");
        sb.AppendLine("### Per-Cluster Analysis");
        sb.AppendLine();
        sb.AppendLine("**Column Legend:**");
        sb.AppendLine("- COM_L/T/W/H = Excel Range.Left/Top/Width/Height (points)");
        sb.AppendLine("- DB_L/T/R/B = Database LeftPosition/TopPosition/RightPosition/BottomPosition (ratios)");
        sb.AppendLine("- Gen = Computed from pure COM formula: RoundEx((COM_Pos + Origin) / PageDim, 7)");
        sb.AppendLine("- Δ = DB - Gen (positive = COM needs MORE vertical/horizontal offset)");
        sb.AppendLine();
        sb.AppendLine("| # | Cell | Row | Col | RowHt | ColW | Font | FontSz | COM_L | COM_T | COM_W | COM_H | DB_L | DB_T | Gen_L | Gen_T | ΔL(×10⁶) | ΔT(×10⁶) | Notes |");
        sb.AppendLine("|---|------|:---:|:---:|:-----:|:----:|:----:|:------:|:-----:|:-----:|:-----:|:-----:|:----:|:----:|:-----:|:-----:|:--------:|:--------:|:-----:|");

        double pw = comMeta.PageWidth;
        double ph = comMeta.PageHeight;
        double ox = comMeta.PrintedOriginX;
        double oy = comMeta.PrintedOriginY;

        int clusterIdx = 0;
        foreach (var dbc in dbClusters)
        {
            clusterIdx++;
            var addr = dbc.CellAddr?.Trim() ?? "";
            if (string.IsNullOrEmpty(addr)) continue;

            double dbL = ParseDouble(dbc.LeftPosition);
            double dbT = ParseDouble(dbc.TopPosition);
            double dbR = ParseDouble(dbc.RightPosition);
            double dbB = ParseDouble(dbc.BottomPosition);

            // Look up COM data
            if (!comMeta.Clusters.TryGetValue(addr, out var comCluster))
            {
                sb.AppendLine($"| {clusterIdx} | {addr} | ? | ? | ? | ? | ? | ? | ? | ? | ? | ? | {dbL:F7} | {dbT:F7} | ? | ? | ? | ? | NO COM DATA |");
                continue;
            }

            double genL = ComCoordinateService.ComputeRatio(comCluster.Left, ox, pw);
            double genT = ComCoordinateService.ComputeRatio(comCluster.Top, oy, ph);
            double genR = ComCoordinateService.ComputeRatio(comCluster.Left + comCluster.Width, ox, pw);
            double genB = ComCoordinateService.ComputeRatio(comCluster.Top + comCluster.Height, oy, ph);

            // Also compute using merge area if available
            double genMergeL = genL, genMergeT = genT;
            if (comCluster.MergeLeft.HasValue)
            {
                genMergeL = ComCoordinateService.ComputeRatio(comCluster.MergeLeft.Value, ox, pw);
                genMergeT = ComCoordinateService.ComputeRatio(comCluster.MergeTop.Value, oy, ph);
            }

            double deltaL = (dbL - genL) * 1_000_000;
            double deltaT = (dbT - genT) * 1_000_000;

            // Notes
            string notes = "";
            if (Math.Abs(deltaL) < 1 && Math.Abs(deltaT) < 1) notes = "PERFECT MATCH";
            else if (clusterIdx == 1) notes = "ORIGIN CLUSTER";
            else if (Math.Abs(deltaL) < 1) notes = $"Top drift only: {deltaT / 1_000_000 * ph:F2}pt";
            else if (Math.Abs(deltaT) < 1) notes = $"Left drift only: {deltaL / 1_000_000 * pw:F2}pt";
            else notes = $"Both drift: L={deltaL / 1_000_000 * pw:F2}pt T={deltaT / 1_000_000 * ph:F2}pt";

            // Compute alternative transforms
            double printableRatio = (comCluster.Left + comMeta.PrintableOriginX) / comMeta.PrintableWidth;
            double printableRatioT = (comCluster.Top + comMeta.PrintableOriginY) / comMeta.PrintableHeight;

            // Try with ClusterGap
            double gapL = ComCoordinateService.ComputeRatio(comCluster.Left + comMeta.ClusterGapX, ox, pw);
            double gapT = ComCoordinateService.ComputeRatio(comCluster.Top + comMeta.ClusterGapY, oy, ph);

            // Row number and column letter from address
            var (rowNum, colLetter, colNum) = ParseAddress(addr);

            sb.AppendLine($"| {clusterIdx} | {addr} | {rowNum} | {colLetter}({colNum}) | " +
                $"{comCluster.RowHeight:F2} | {comCluster.ColWidth:F2} | " +
                $"{comCluster.FontName ?? "?"} | {comCluster.FontSize ?? 11:F0} | " +
                $"{comCluster.Left:F2} | {comCluster.Top:F2} | " +
                $"{comCluster.Width:F2} | {comCluster.Height:F2} | " +
                $"{dbL:F7} | {dbT:F7} | " +
                $"{genL:F7} | {genT:F7} | " +
                $"{deltaL:F0} | {deltaT:F0} | " +
                $"{notes} |");

            // Detailed transform testing
            if (clusterIdx == 1 || Math.Abs(deltaL) > 1 || Math.Abs(deltaT) > 1)
            {
                sb.AppendLine($"  - Merge pos: L={comCluster.MergeLeft?.ToString("F2") ?? "N/A"} T={comCluster.MergeTop?.ToString("F2") ?? "N/A"}");
                sb.AppendLine($"  - Gen_Merge: L={genMergeL:F7} T={genMergeT:F7}");
                sb.AppendLine($"  - Printable ratio: L={printableRatio:F7} T={printableRatioT:F7}");
                sb.AppendLine($"  - With gap offset: L={gapL:F7} T={gapT:F7}");
                sb.AppendLine($"  - COM+Origin: L={comCluster.Left + ox:F2} T={comCluster.Top + oy:F2}");
                sb.AppendLine($"  - DB in points: L={dbL * pw:F2} T={dbT * ph:F2}");
                sb.AppendLine($"  - Needed offset: L={dbL * pw - comCluster.Left:F2} T={dbT * ph - comCluster.Top:F2} (what origin would make this cluster match)");
            }
        }
        sb.AppendLine();

        // Step 6: Origin analysis
        sb.AppendLine("### Origin Analysis");
        sb.AppendLine();
        sb.AppendLine($"- **Computed Origin:** ({ox:F4}, {oy:F4}) pt");
        sb.AppendLine($"- **From margins (L, T):** ({comMeta.LeftMargin:F2}, {comMeta.TopMargin:F2}) pt");
        sb.AppendLine($"- **From center:** H={comMeta.CenterHorizontally}, V={comMeta.CenterVertically}");
        sb.AppendLine($"- **Printable Area Origin:** ({comMeta.PrintableOriginX:F2}, {comMeta.PrintableOriginY:F2}) pt");
        sb.AppendLine($"- **Printable Area Dims:** ({comMeta.PrintableWidth:F2} x {comMeta.PrintableHeight:F2}) pt");
        sb.AppendLine($"- **Cluster Gap (from 1st cluster):** ({comMeta.ClusterGapX:F6}, {comMeta.ClusterGapY:F6}) pt");

        // Compute effective scale
        double scaleX = comMeta.PageWidth / comMeta.UsedWidth;
        double scaleY = comMeta.PageHeight / comMeta.UsedHeight;
        sb.AppendLine($"- **UsedRange Scale:** X={scaleX:F6} Y={scaleY:F6}");
        sb.AppendLine($"- **Print scaling factor:** Zoom={comMeta.Zoom}% FitToPages={comMeta.FitToPagesWide}x{comMeta.FitToPagesTall}");
        sb.AppendLine();

        // Step 7: Compute per-row gap for each template
        sb.AppendLine("### Row-Gap Analysis");
        sb.AppendLine();
        sb.AppendLine("Computes the actual vertical gap needed to make each cluster's Top match the DB:");
        sb.AppendLine();
        sb.AppendLine("| Row | DB_T(pts) | COM_T(pts) | NeededGap(pts) | RowHeight | ΔGap |");
        sb.AppendLine("|-----|-----------|------------|----------------|-----------|------|");

        foreach (var dbc in dbClusters)
        {
            var addr = dbc.CellAddr?.Trim() ?? "";
            if (string.IsNullOrEmpty(addr) || !comMeta.Clusters.TryGetValue(addr, out var cc)) continue;

            double dbTPt = ParseDouble(dbc.TopPosition) * ph;
            double comTPt = cc.Top;
            double neededGap = dbTPt - (comTPt + oy);
            var (rowNum, _, _) = ParseAddress(addr);

            sb.AppendLine($"| {rowNum} | {dbTPt:F4} | {comTPt:F4} | {neededGap:F4} | {cc.RowHeight:F2} | {neededGap - (rowNum > 1 ? 0 : 0):F4} |");
        }
        sb.AppendLine();

        // Step 8: Per-column gap analysis
        sb.AppendLine("### Column-Gap Analysis");
        sb.AppendLine();
        sb.AppendLine("| Col | DB_L(pts) | COM_L(pts) | NeededGap(pts) | ColWidth | OpenXML_W |");
        sb.AppendLine("|-----|-----------|------------|----------------|----------|-----------|");

        foreach (var dbc in dbClusters)
        {
            var addr = dbc.CellAddr?.Trim() ?? "";
            if (string.IsNullOrEmpty(addr) || !comMeta.Clusters.TryGetValue(addr, out var cc)) continue;

            double dbLPt = ParseDouble(dbc.LeftPosition) * pw;
            double comLPt = cc.Left;
            double neededGap = dbLPt - (comLPt + ox);
            var (_, colLetter, colNum) = ParseAddress(addr);

            // Get OpenXML width for this column
            string oxWidth = "?";
            if (openXmlInfo.ColumnWidths.TryGetValue(colNum, out var oxw))
                oxWidth = oxw.ToString("F3");

            sb.AppendLine($"| {colLetter}({colNum}) | {dbLPt:F4} | {comLPt:F4} | {neededGap:F4} | {cc.ColWidth:F2} | {oxWidth} |");
        }
        sb.AppendLine();

        // Step 9: Try to fit a per-row gap formula
        sb.AppendLine("### Row-Gap Formula Fitting");
        sb.AppendLine();
        sb.AppendLine("For each non-origin cluster, compute the gap needed per row from the origin:");
        sb.AppendLine();

        // Get sorted list of non-origin clusters
        var nonOrigin = new List<(int Row, double NeededGap, double RowHeight)>();
        foreach (var dbc in dbClusters)
        {
            var addr = dbc.CellAddr?.Trim() ?? "";
            if (string.IsNullOrEmpty(addr) || !comMeta.Clusters.TryGetValue(addr, out var cc)) continue;
            var (rowNum, _, _) = ParseAddress(addr);
            if (rowNum == 0) continue;
            double dbTPt = ParseDouble(dbc.TopPosition) * ph;
            double comTPt = cc.Top;
            double neededGap = dbTPt - (comTPt + oy);
            nonOrigin.Add((rowNum, neededGap, cc.RowHeight));
        }

        if (nonOrigin.Count >= 2)
        {
            var originRow = nonOrigin.Min(r => r.Row);
            var originCluster = nonOrigin.FirstOrDefault(n => n.Row == originRow);
            double baseGap = originCluster.NeededGap;

            sb.AppendLine("| Row | RowOffset(rows) | TotGap(pt) | GapPerRow(pt) | RowHt | GapAs%RowHt |");
            sb.AppendLine("|-----|:---------------:|:----------:|:-------------:|:-----:|:-----------:|");
            foreach (var (row, gap, rh) in nonOrigin.OrderBy(n => n.Row))
            {
                int rowsFromOrigin = row - originRow;
                double gapFromBase = gap - baseGap;
                double perRow = rowsFromOrigin > 0 ? gapFromBase / rowsFromOrigin : 0;
                double pctOfRh = rh > 0 ? perRow / rh * 100 : 0;
                sb.AppendLine($"| {row} | {rowsFromOrigin} | {gap:F4} | {perRow:F4} | {rh:F2} | {pctOfRh:F1}% |");
            }
        }
        sb.AppendLine();

        return sb.ToString();
    }

    private ComMetadata ExtractComMetadata(string workbookPath, List<DefCluster> dbClusters)
    {
        var meta = new ComMetadata();

        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new InvalidOperationException("Excel COM is not available.");

        dynamic excel = Activator.CreateInstance(excelType);
        try
        {
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;

            dynamic wb = excel.Workbooks.Open(workbookPath);
            try
            {
                dynamic ws = wb.Sheets[1];
                dynamic ps = ws.PageSetup;

                // Page setup
                meta.PageWidth = SafeReadDouble(() => (double)ps.PageWidth, 612);
                meta.PageHeight = SafeReadDouble(() => (double)ps.PageHeight, 792);
                meta.LeftMargin = SafeReadDouble(() => (double)ps.LeftMargin, 0);
                meta.RightMargin = SafeReadDouble(() => (double)ps.RightMargin, 0);
                meta.TopMargin = SafeReadDouble(() => (double)ps.TopMargin, 0);
                meta.BottomMargin = SafeReadDouble(() => (double)ps.BottomMargin, 0);
                meta.CenterHorizontally = SafeReadBool(() => (bool)ps.CenterHorizontally);
                meta.CenterVertically = SafeReadBool(() => (bool)ps.CenterVertically);
                meta.Zoom = SafeReadInt(() => (int)ps.Zoom);
                meta.FitToPagesWide = SafeReadInt(() => (int)ps.FitToPagesWide);
                meta.FitToPagesTall = SafeReadInt(() => (int)ps.FitToPagesTall);
                meta.PrintArea = SafeReadString(() => (string)ps.PrintArea);
                meta.PaperSize = SafeReadInt(() => (int)ps.PaperSize);
                meta.Orientation = SafeReadInt(() => (int)ps.Orientation);

                // Printable area = page - margins
                meta.PrintableWidth = Math.Max(meta.PageWidth - meta.LeftMargin - meta.RightMargin, 1);
                meta.PrintableHeight = Math.Max(meta.PageHeight - meta.TopMargin - meta.BottomMargin, 1);
                meta.PrintableOriginX = meta.CenterHorizontally
                    ? (meta.PageWidth - meta.PrintableWidth) / 2.0
                    : meta.LeftMargin;
                meta.PrintableOriginY = meta.CenterVertically
                    ? (meta.PageHeight - meta.PrintableHeight) / 2.0
                    : meta.TopMargin;

                // UsedRange
                try
                {
                    dynamic ur = ws.UsedRange;
                    meta.UsedLeft = (double)ur.Left;
                    meta.UsedTop = (double)ur.Top;
                    meta.UsedWidth = (double)ur.Width;
                    meta.UsedHeight = (double)ur.Height;
                    meta.UsedRight = meta.UsedLeft + meta.UsedWidth;
                    meta.UsedBottom = meta.UsedTop + meta.UsedHeight;
                }
                catch { }

                // Default row height & column width
                try { meta.DefaultRowHeight = (double)ws.StandardHeight; } catch { }
                try { meta.DefaultColumnWidth = (double)ws.StandardWidth; } catch { }

                // Read each DB cluster's COM metadata (extended)
                _progress.Report("  Reading extended COM per-cluster...");
                double firstComLeft = 0, firstComTop = 0;
                string? firstAddr = null;
                double firstDbLeft = 0, firstDbTop = 0;

                foreach (var dbc in dbClusters)
                {
                    var addr = dbc.CellAddr?.Trim();
                    if (string.IsNullOrEmpty(addr)) continue;

                    try
                    {
                        dynamic range = ws.Range[addr];
                        var cc = new ComClusterData
                        {
                            Address = addr,
                            Left = (double)range.Left,
                            Top = (double)range.Top,
                            Width = (double)range.Width,
                            Height = (double)range.Height
                        };

                        // Row height (from the range)
                        try { cc.RowHeight = (double)range.RowHeight; } catch { cc.RowHeight = 14.4; }

                        // Column width (from the range's first cell's column)
                        try { cc.ColWidth = (double)range.ColumnWidth; } catch { cc.ColWidth = 8.43; }

                        // Font
                        try
                        {
                            dynamic font = range.Font;
                            cc.FontName = (string)font.Name;
                            cc.FontSize = (double)font.Size;
                            cc.FontBold = (bool)font.Bold;
                        }
                        catch { }

                        // Merge area
                        try
                        {
                            dynamic mergeArea = range.MergeArea;
                            if (mergeArea != null)
                            {
                                var mergeAddr = (string)mergeArea.Address;
                                if (mergeAddr != addr)
                                {
                                    cc.MergeAddress = mergeAddr;
                                    cc.MergeLeft = (double)mergeArea.Left;
                                    cc.MergeTop = (double)mergeArea.Top;
                                    cc.MergeWidth = (double)mergeArea.Width;
                                    cc.MergeHeight = (double)mergeArea.Height;
                                }
                            }
                        }
                        catch { }

                        // Interior color
                        try
                        {
                            dynamic interior = range.Interior;
                            cc.InteriorColor = (int)interior.Color;
                        }
                        catch { }

                        meta.Clusters[addr] = cc;

                        // Save first cluster for origin derivation
                        if (firstAddr == null)
                        {
                            firstAddr = addr;
                            firstComLeft = cc.Left;
                            firstComTop = cc.Top;
                            firstDbLeft = ParseDouble(dbc.LeftPosition);
                            firstDbTop = ParseDouble(dbc.TopPosition);
                        }
                    }
                    catch (Exception ex)
                    {
                        _progress.Report($"    Warning: Could not read {addr}: {ex.Message}");
                    }
                }

                _progress.Report($"  Extracted {meta.Clusters.Count} cluster positions");

                // Derive PrintedOrigin
                if (firstAddr != null)
                {
                    meta.PrintedOriginX = ComCoordinateService.RoundEx(firstDbLeft * meta.PageWidth - firstComLeft);
                    meta.PrintedOriginY = ComCoordinateService.RoundEx(firstDbTop * meta.PageHeight - firstComTop);

                    // Cluster gap from first cluster
                    var firstDb = dbClusters.First(c => c.CellAddr?.Trim() == firstAddr);
                    double dbR = ParseDouble(firstDb.RightPosition);
                    double dbB = ParseDouble(firstDb.BottomPosition);
                    if (meta.Clusters.TryGetValue(firstAddr, out var firstPos))
                    {
                        meta.ClusterGapX = ComCoordinateService.RoundEx(dbR * meta.PageWidth - (firstPos.Left + firstPos.Width + meta.PrintedOriginX));
                        meta.ClusterGapY = ComCoordinateService.RoundEx(dbB * meta.PageHeight - (firstPos.Top + firstPos.Height + meta.PrintedOriginY));
                    }

                    _progress.Report($"  Origin: ({meta.PrintedOriginX:F2}, {meta.PrintedOriginY:F2}) " +
                        $"Gap: ({meta.ClusterGapX:F6}, {meta.ClusterGapY:F6})");
                }
            }
            finally
            {
                wb.Close(false);
                Marshal.ReleaseComObject(wb);
            }
        }
        finally
        {
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }

        return meta;
    }

    private OpenXmlInfo ReadOpenXmlInfo(string workbookPath)
    {
        var info = new OpenXmlInfo();
        try
        {
            using var doc = SpreadsheetDocument.Open(workbookPath, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return info;

            // PrintArea from defined names
            var wbXml = XDocument.Load(wbPart.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var printAreaDef = wbXml.Descendants(ns + "definedName")
                .FirstOrDefault(d =>
                {
                    var name = (string?)d.Attribute("name");
                    return name != null && (name == "_xlnm.Print_Area" || name.EndsWith("Print_Area", StringComparison.OrdinalIgnoreCase));
                });
            if (printAreaDef != null)
                info.PrintArea = printAreaDef.Value?.Trim();

            // Sheets
            var sheetsTags = wbXml.Descendants(ns + "sheet");
            foreach (var s in sheetsTags)
            {
                var name = (string?)s.Attribute("name") ?? "";
                info.Sheets.Add(name);
            }

            // Shared strings
            var sstPart = wbPart.SharedStringTablePart;
            if (sstPart != null)
            {
                using var sstStream = sstPart.GetStream();
                var sstXml = XDocument.Load(sstStream);
                foreach (var si in sstXml.Descendants(ns + "si"))
                {
                    var t = si.Descendants(ns + "t").FirstOrDefault();
                    info.SharedStrings.Add(t?.Value ?? "");
                }
            }

            // First sheet's data
            var wsPart = wbPart.WorksheetParts.FirstOrDefault();
            if (wsPart == null) return info;

            var wsXml = XDocument.Load(wsPart.GetStream());

            // Column widths
            var colsTag = wsXml.Descendants(ns + "cols").FirstOrDefault();
            if (colsTag != null)
            {
                foreach (var col in colsTag.Descendants(ns + "col"))
                {
                    int min = (int?)col.Attribute("min") ?? 1;
                    int max = (int?)col.Attribute("max") ?? 1;
                    double width = (double?)col.Attribute("width") ?? 8.43;
                    bool customWidth = (string?)col.Attribute("customWidth") == "1";
                    for (int i = min; i <= max; i++)
                        info.ColumnWidths[i] = width;
                }
            }

            // Row heights
            var sheetDataTag = wsXml.Descendants(ns + "sheetData").FirstOrDefault();
            if (sheetDataTag != null)
            {
                foreach (var row in sheetDataTag.Descendants(ns + "row"))
                {
                    int r = (int?)row.Attribute("r") ?? 0;
                    if (r > 0)
                    {
                        double ht = (double?)row.Attribute("ht") ?? 14.4;
                        bool custom = (string?)row.Attribute("customHeight") == "1";
                        info.RowHeights[r] = (ht, custom);
                    }
                }
            }

            // Merged cells
            var mergesTag = wsXml.Descendants(ns + "mergeCells").FirstOrDefault();
            if (mergesTag != null)
            {
                foreach (var mc in mergesTag.Descendants(ns + "mergeCell"))
                {
                    var refVal = (string?)mc.Attribute("ref") ?? "";
                    if (!string.IsNullOrEmpty(refVal))
                        info.MergedCells.Add(refVal);
                }
            }
        }
        catch (Exception ex)
        {
            _progress.Report($"  OpenXML read failed: {ex.Message}");
        }
        return info;
    }

    private static double ParseDouble(string? s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;

    private static double ColumnCharsToPoints(double chars) => chars * 50.04 / 8.43;

    private static (int Row, string ColLetter, int ColNum) ParseAddress(string addr)
    {
        // Handle formats: $A$1:$B$2, A1, $A1
        addr = addr.Replace("$", "").Split(':')[0];
        var letters = new string(addr.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(addr.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        int row = int.TryParse(digits, out var r) ? r : 0;
        int colNum = 0;
        foreach (var ch in letters)
            colNum = colNum * 26 + (char.ToUpper(ch) - 'A' + 1);
        return (row, letters, colNum);
    }

    private static double SafeReadDouble(Func<double> getter, double fallback)
    {
        try { return getter(); } catch { return fallback; }
    }

    private static bool SafeReadBool(Func<bool> getter)
    {
        try { return getter(); } catch { return false; }
    }

    private static int SafeReadInt(Func<int> getter)
    {
        try { return getter(); } catch { return 0; }
    }

    private static string SafeReadString(Func<string> getter)
    {
        try { return getter() ?? ""; } catch { return ""; }
    }
}

/// <summary>Extensive per-cluster data from COM.</summary>
public class ComClusterData
{
    public string Address { get; set; } = "";
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }

    public double RowHeight { get; set; }
    public double ColWidth { get; set; }

    public string? FontName { get; set; }
    public double? FontSize { get; set; }
    public bool? FontBold { get; set; }

    public string? MergeAddress { get; set; }
    public double? MergeLeft { get; set; }
    public double? MergeTop { get; set; }
    public double? MergeWidth { get; set; }
    public double? MergeHeight { get; set; }

    public int? InteriorColor { get; set; }
}

/// <summary>Complete COM metadata for one workbook.</summary>
public class ComMetadata
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
    public int? Zoom { get; set; }
    public int? FitToPagesWide { get; set; }
    public int? FitToPagesTall { get; set; }
    public string? PrintArea { get; set; }
    public int PaperSize { get; set; }
    public int Orientation { get; set; }

    // Printable area
    public double PrintableWidth { get; set; }
    public double PrintableHeight { get; set; }
    public double PrintableOriginX { get; set; }
    public double PrintableOriginY { get; set; }

    // UsedRange
    public double UsedLeft { get; set; }
    public double UsedTop { get; set; }
    public double UsedWidth { get; set; }
    public double UsedHeight { get; set; }
    public double UsedRight { get; set; }
    public double UsedBottom { get; set; }

    // Default sizes
    public double DefaultRowHeight { get; set; } = 14.4;
    public double DefaultColumnWidth { get; set; } = 8.43;

    // Derived origin
    public double PrintedOriginX { get; set; }
    public double PrintedOriginY { get; set; }
    public double ClusterGapX { get; set; }
    public double ClusterGapY { get; set; }

    // Per-cluster data
    public Dictionary<string, ComClusterData> Clusters { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>OpenXML extracted info.</summary>
public class OpenXmlInfo
{
    public List<string> Sheets { get; set; } = new();
    public List<string> SharedStrings { get; set; } = new();
    public Dictionary<int, double> ColumnWidths { get; set; } = new();
    public Dictionary<int, (double Height, bool IsCustom)> RowHeights { get; set; } = new();
    public List<string> MergedCells { get; set; } = new();
    public string? PrintArea { get; set; }
}
