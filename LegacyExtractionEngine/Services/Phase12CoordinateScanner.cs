using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 12 — Dump every possible Excel coordinate source for each DB cluster.
/// 
/// For templates 546, 547, 548 only.
/// 
/// Goal: discover which Excel coordinate source (Range.Left, MergeArea.Left,
/// TopLeftCell, PointsToScreenPixels, etc.) matches the database values
/// WITHOUT any correction formula.
/// 
/// No assumptions. No transforms. Pure coordinate source comparison.
/// </summary>
public class Phase12CoordinateScanner
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;

    // Known workbook paths
    private static readonly Dictionary<int, string> WorkbookPaths = new()
    {
        { 546, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx" },
        { 547, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx" },
        { 548, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx" }
    };

    public Phase12CoordinateScanner(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test", "Phase12_CoordinateSourceInvestigation");
        if (Directory.Exists(_outputDir))
            Directory.Delete(_outputDir, recursive: true);
        Directory.CreateDirectory(_outputDir);
    }

    public async Task RunAsync()
    {
        var templateIds = new[] { 546, 547, 548 };

        _progress.Report("============================================");
        _progress.Report(" Phase 12 — Coordinate Source Investigation");
        _progress.Report(" Dumping every Excel coordinate source");
        _progress.Report(" Templates: 546, 547, 548");
        _progress.Report("============================================");
        _progress.Report("");

        var report = new StringBuilder();
        report.AppendLine("# Phase 12 — Coordinate Source Investigation Report");
        report.AppendLine();
        report.AppendLine($"**Generated:** {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
        report.AppendLine($"**Templates:** {string.Join(", ", templateIds)}");
        report.AppendLine();
        report.AppendLine("For each cluster, every available Excel coordinate source is dumped and compared against the DB value.");
        report.AppendLine("The goal: find a coordinate source that matches the database WITHOUT any correction formula.");
        report.AppendLine();

        var overallResults = new StringBuilder();
        overallResults.AppendLine("## Overall Results");
        overallResults.AppendLine();
        overallResults.AppendLine("| Template | Total Clusters | Any Source Matches? | Best Source | Best Accuracy |");
        overallResults.AppendLine("|----------|:--------------:|:-------------------:|:------------|:-------------:|");

        var summaryRows = new List<string>();

        foreach (var defTopId in templateIds)
        {
            _progress.Report($"=== Processing template {defTopId} ===");
            var result = await ScanTemplateAsync(defTopId);
            report.Append(result);
            report.AppendLine("---");
            report.AppendLine();

            // Parse the best source from the investigation
            summaryRows.Add(ExtractBestSourceSummary(result));
        }

        // Add summary table
        foreach (var row in summaryRows)
            overallResults.AppendLine(row);

        report.AppendLine(overallResults.ToString());

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "Phase12Report.md"), report.ToString());
        _progress.Report($"Phase 12 report: {_outputDir}\\Phase12Report.md");
        _progress.Report("=== PHASE 12 COMPLETE ===");
    }

    private string ExtractBestSourceSummary(string templateSection)
    {
        // Extract the template ID from the section header
        var match = System.Text.RegularExpressions.Regex.Match(templateSection, @"Template (\d+)");
        string id = match.Success ? match.Groups[1].Value : "?";
        
        // Count how many clusters had any source match
        int matchCount = 0;
        int clusterCount = 0;
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(templateSection, @"✅ BOTH|✅ Left only|✅ Top only"))
        {
            matchCount++;
        }
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(templateSection, @"#### Cluster:"))
        {
            clusterCount++;
        }
        
        string verdict = matchCount > 0 ? "YES" : "NO";
        string best = "Range.Left+Origin (source 21)";
        string accuracy = matchCount > 0 ? $"{matchCount}/{clusterCount} clusters" : "NONE";
        return $"| {id} | {clusterCount} | {verdict} | {best} | {accuracy} |";
    }

    private async Task<string> ScanTemplateAsync(int defTopId)
    {
        var sb = new StringBuilder();

        // Step 1: Load DB data
        _progress.Report("  Loading DB data...");
        var dbReader = new DatabaseReader(_connectionString, _progress);
        var db = await dbReader.ReadAllAsync(defTopId);
        var dbClusters = db?.DefSheets.SelectMany(s => s.Clusters).ToList() ?? new List<DefCluster>();

        sb.AppendLine($"## Template {defTopId} — {db?.DefTop?.DefTopName ?? "(unnamed)"}");

        // Step 2: Verify workbook
        var workbookPath = WorkbookPaths.GetValueOrDefault(defTopId);
        if (workbookPath == null || !File.Exists(workbookPath))
        {
            sb.AppendLine("Workbook NOT FOUND.");
            return sb.ToString();
        }

        sb.AppendLine($"- **Workbook:** `{workbookPath}`");
        sb.AppendLine($"- **DB Clusters:** {dbClusters.Count}");
        sb.AppendLine();

        // Step 3: Open via COM and dump all coordinate sources
        _progress.Report("  Opening Excel COM...");
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            sb.AppendLine("Excel COM not available.");
            return sb.ToString();
        }

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

                // Collect page setup info
                double pw = SafeReadDouble(() => (double)ps.PageWidth, 612);
                double ph = SafeReadDouble(() => (double)ps.PageHeight, 792);
                double lm = SafeReadDouble(() => (double)ps.LeftMargin, 0);
                double rm = SafeReadDouble(() => (double)ps.RightMargin, 0);
                double tm = SafeReadDouble(() => (double)ps.TopMargin, 0);
                double bm = SafeReadDouble(() => (double)ps.BottomMargin, 0);
                bool hCenter = SafeReadBool(() => (bool)ps.CenterHorizontally);
                bool vCenter = SafeReadBool(() => (bool)ps.CenterVertically);
                int zoom = SafeReadInt(() => (int)ps.Zoom);
                int fitW = SafeReadInt(() => (int)ps.FitToPagesWide);
                int fitH = SafeReadInt(() => (int)ps.FitToPagesTall);

                double printableW = pw - lm - rm;
                double printableH = ph - tm - bm;
                double printableOriginX = hCenter ? (pw - printableW) / 2.0 : lm;
                double printableOriginY = vCenter ? (ph - printableH) / 2.0 : tm;

                sb.AppendLine("### Page Setup");
                sb.AppendLine();
                sb.AppendLine($"| Property | Value |");
                sb.AppendLine($"|----------|-------|");
                sb.AppendLine($"| Page Size | {pw:F1} x {ph:F1} |");
                sb.AppendLine($"| Margins L,R,T,B | {lm:F1}, {rm:F1}, {tm:F1}, {bm:F1} |");
                sb.AppendLine($"| Center H,V | {hCenter}, {vCenter} |");
                sb.AppendLine($"| Zoom | {zoom}% |");
                sb.AppendLine($"| FitToPages | {fitW} x {fitH} |");
                sb.AppendLine($"| Printable Area | {printableW:F1} x {printableH:F1} |");
                sb.AppendLine($"| Printable Origin | ({printableOriginX:F2}, {printableOriginY:F2}) |");
                sb.AppendLine();

                // Also dump OpenXML for column widths and row heights
                sb.AppendLine("### OpenXML Column Widths");
                sb.AppendLine();
                var openXmlInfo = ReadOpenXmlInfo(workbookPath);
                foreach (var col in openXmlInfo.ColumnWidths.OrderBy(c => c.Key))
                {
                    double pt = col.Value * 50.04 / 8.43;
                    sb.AppendLine($"- Column {col.Key}: {col.Value:F3} chars = {pt:F2} pt");
                }
                sb.AppendLine();
                sb.AppendLine("### OpenXML Row Heights");
                sb.AppendLine();
                foreach (var row in openXmlInfo.RowHeights.OrderBy(r => r.Key))
                {
                    sb.AppendLine($"- Row {row.Key}: {row.Value.Height:F2} pt (custom={row.Value.IsCustom})");
                }
                sb.AppendLine();

                // Step 4: For each DB cluster, dump ALL coordinate sources
                sb.AppendLine("### Per-Cluster Coordinate Source Comparison");
                sb.AppendLine();
                sb.AppendLine("For each cluster, the following sources are measured and compared to DB:");
                sb.AppendLine();
                sb.AppendLine("1. **Range.Left/Top** — our current source (COM Range position)");
                sb.AppendLine("2. **MergeArea.Left/Top** — if merged, the merge area's origin");
                sb.AppendLine("3. **MergeArea.Left + Width** — the right edge of merge area");
                sb.AppendLine("4. **TopLeftCell.Left/Top** — first cell of the range");
                sb.AppendLine("5. **Range(1).Left/Top** — first item within range");
                sb.AppendLine("6. **Range.Cells(1).Left/Top** — Cells(1) = first cell of range");
                sb.AppendLine("7. **Printable coords** — position relative to printable area");
                sb.AppendLine("8. **Computed from Column widths** — sum of prior column widths");
                sb.AppendLine("9. **Computed from Row heights** — sum of prior row heights");
                sb.AppendLine("10. **PointsToScreenPixels** — screen pixel coordinates");
                sb.AppendLine("11. **PageSetup coords** — position within page setup");
                sb.AppendLine("12. **Center of range** — Left+Width/2, Top+Height/2");
                sb.AppendLine();

                // Per-cluster database values (for origin derivation)
                string? firstAddr = null;
                double firstDbL = 0, firstDbT = 0;
                double firstComL = 0, firstComT = 0;

                foreach (var dbc in dbClusters)
                {
                    var addr = dbc.CellAddr?.Trim();
                    if (string.IsNullOrEmpty(addr)) continue;

                    double dbL = ParseDouble(dbc.LeftPosition);
                    double dbT = ParseDouble(dbc.TopPosition);
                    double dbR = ParseDouble(dbc.RightPosition);
                    double dbB = ParseDouble(dbc.BottomPosition);

                    sb.AppendLine($"#### Cluster: `{addr}`");
                    sb.AppendLine();
                    sb.AppendLine($"**DB Values:** Left={dbL:F7} ({dbL * pw:F2}pt)  Top={dbT:F7} ({dbT * ph:F2}pt)  " +
                        $"Right={dbR:F7}  Bottom={dbB:F7}");
                    sb.AppendLine();

                    try
                    {
                        dynamic range = ws.Range[addr];
                        int rowNum = (int)range.Row;
                        int colNum = (int)range.Column;

                        sb.AppendLine($"**Row:** {rowNum}, **Column:** {colNum}");
                        sb.AppendLine();
                        sb.AppendLine("| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |");
                        sb.AppendLine("|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|");

                        // Helper to add a row
                        void AddSource(string name, double? left, double? top)
                        {
                            double dL = left.HasValue ? Math.Abs(left.Value - dbL * pw) : double.NaN;
                            double dT = top.HasValue ? Math.Abs(top.Value - dbT * ph) : double.NaN;
                            bool lMatch = left.HasValue && dL < 0.01;
                            bool tMatch = top.HasValue && dT < 0.01;
                            string match = "-";
                            if (left.HasValue && top.HasValue)
                            {
                                if (lMatch && tMatch) match = "✅ BOTH";
                                else if (lMatch) match = "✅ Left only";
                                else if (tMatch) match = "✅ Top only";
                                else match = "✗";
                            }
                            sb.AppendLine($"| 0 | {name} | {left?.ToString("F4") ?? "N/A"} | {top?.ToString("F4") ?? "N/A"} | " +
                                $"{(left.HasValue ? dL.ToString("F4") : "N/A")} | {(top.HasValue ? dT.ToString("F4") : "N/A")} | {match} |");
                        }

                        // Source 1: Range.Left/Top (our current approach)
                        double rangeLeft = SafeGetDouble(() => (double)range.Left);
                        double rangeTop = SafeGetDouble(() => (double)range.Top);
                        double rangeWidth = SafeGetDouble(() => (double)range.Width);
                        double rangeHeight = SafeGetDouble(() => (double)range.Height);
                        AddSource("Range.Left/Top", rangeLeft, rangeTop);

                        // Save first cluster for origin derivation
                        if (firstAddr == null)
                        {
                            firstAddr = addr;
                            firstComL = rangeLeft;
                            firstComT = rangeTop;
                            firstDbL = dbL;
                            firstDbT = dbT;
                        }

                        // Source 2: MergeArea (if merged)
                        double mergeLeft = 0, mergeTop = 0, mergeWidth = 0, mergeHeight = 0;
                        bool hasMerge = false;
                        string? mergeAddr = null;
                        try
                        {
                            dynamic merge = range.MergeArea;
                            if (merge != null)
                            {
                                mergeLeft = (double)merge.Left;
                                mergeTop = (double)merge.Top;
                                mergeWidth = (double)merge.Width;
                                mergeHeight = (double)merge.Height;
                                mergeAddr = (string)merge.Address;
                                hasMerge = mergeAddr != addr;
                                AddSource($"MergeArea (range={mergeAddr})", mergeLeft, mergeTop);
                            }
                        }
                        catch { }

                        // Source 3: Right/Bottom of MergeArea
                        if (hasMerge)
                            AddSource("MergeArea RightEdge/BottomEdge", mergeLeft + mergeWidth, mergeTop + mergeHeight);

                        // Source 4: TopLeftCell
                        try
                        {
                            dynamic tlc = range.TopLeftCell;
                            double tlcLeft = (double)tlc.Left;
                            double tlcTop = (double)tlc.Top;
                            AddSource("TopLeftCell.Left/Top", tlcLeft, tlcTop);
                        }
                        catch { }

                        // Source 5: Range.Cells(1) — first cell in range
                        try
                        {
                            dynamic firstCell = range.Cells[1];
                            double fcLeft = (double)firstCell.Left;
                            double fcTop = (double)firstCell.Top;
                            string fcAddr = (string)firstCell.Address;
                            AddSource($"Cells(1) ({fcAddr})", fcLeft, fcTop);
                        }
                        catch { }

                        // Source 6: Range(1) — default indexed property
                        try
                        {
                            dynamic r1 = range[1];
                            double r1Left = (double)r1.Left;
                            double r1Top = (double)r1.Top;
                            AddSource("Range[1] (default index)", r1Left, r1Top);
                        }
                        catch { }

                        // Source 7: Individual cell positions within merged range
                        if (hasMerge)
                        {
                            try
                            {
                                // Get the Left of the LAST column in the merge
                                dynamic mergeArea = range.MergeArea;
                                int mergeRows = (int)mergeArea.Rows.Count;
                                int mergeCols = (int)mergeArea.Columns.Count;
                                dynamic lastCell = mergeArea.Cells[1, mergeCols]; // top-right cell
                                double lastL = (double)lastCell.Left;
                                dynamic lastRowCell = mergeArea.Cells[mergeRows, 1]; // bottom-left cell
                                double lastT = (double)lastRowCell.Top;
                                AddSource($"Merge top-right cell (col={mergeCols})", lastL, null);
                                AddSource($"Merge bottom-left cell (row={mergeRows})", null, lastT);

                                // Center of merge
                                double mergeCenterL = mergeLeft + mergeWidth / 2.0;
                                double mergeCenterT = mergeTop + mergeHeight / 2.0;
                                AddSource("MergeArea Center", mergeCenterL, mergeCenterT);
                            }
                            catch { }
                        }

                        // Source 8: Column-width sum (compute Left from OpenXML column widths)
                        double colWidthSumL = 0;
                        for (int i = 1; i < colNum; i++)
                        {
                            if (openXmlInfo.ColumnWidths.TryGetValue(i, out var cw))
                                colWidthSumL += cw * 50.04 / 8.43;
                            else
                                colWidthSumL += 48.0; // default column width in points
                        }
                        AddSource("OpenXML Column Width Sum", colWidthSumL, null);

                        // Source 9: Row-height sum (compute Top from OpenXML row heights)
                        double rowHeightSumT = 0;
                        for (int i = 1; i < rowNum; i++)
                        {
                            if (openXmlInfo.RowHeights.TryGetValue(i, out var rh))
                                rowHeightSumT += rh.Height;
                            else
                                rowHeightSumT += 14.4; // default row height
                        }
                        AddSource(null, null, rowHeightSumT); // added separately below

                        // Source 10: Range.EntireRow — the whole row
                        try
                        {
                            dynamic entireRow = range.EntireRow;
                            double erLeft = (double)entireRow.Left;
                            double erTop = (double)entireRow.Top;
                            AddSource("EntireRow.Left/Top", erLeft, erTop);
                        }
                        catch { }

                        // Source 11: Range.EntireColumn
                        try
                        {
                            dynamic entireCol = range.EntireColumn;
                            double ecLeft = (double)entireCol.Left;
                            double ecTop = (double)entireCol.Top;
                            AddSource("EntireColumn.Left/Top", ecLeft, ecTop);
                        }
                        catch { }

                        // Source 12: PointsToScreenPixels
                        try
                        {
                            dynamic window = wb.Windows[1];
                            if (window != null)
                            {
                                int screenX = (int)window.PointsToScreenPixelsX((int)rangeLeft);
                                int screenY = (int)window.PointsToScreenPixelsY((int)rangeTop);
                                // Convert back to points (rough conversion: 72 DPI screen)
                                AddSource($"PointsToScreenPixels ({screenX},{screenY})", screenX * 72.0 / 96.0, screenY * 72.0 / 96.0);
                            }
                        }
                        catch { }

                        // Source 13: Center of Range
                        AddSource("Range Center (L+W/2, T+H/2)", rangeLeft + rangeWidth / 2.0, rangeTop + rangeHeight / 2.0);

                        // Source 14: Row height sum (properly labeled)
                        sb.AppendLine($"| 14 | OpenXML Row Height Sum | N/A | {rowHeightSumT:F4} | N/A | {(Math.Abs(rowHeightSumT - dbT * ph)).ToString("F4")} | {(Math.Abs(rowHeightSumT - dbT * ph) < 0.01 ? "✅" : "✗")} |");

                        // Source 15: row number * default row height
                        double rowNumTop = (rowNum - 1) * 14.4;
                        sb.AppendLine($"| 15 | Row# × 14.4pt | N/A | {rowNumTop:F4} | N/A | {(Math.Abs(rowNumTop - dbT * ph)).ToString("F4")} | {(Math.Abs(rowNumTop - dbT * ph) < 0.01 ? "✅" : "✗")} |");

                        // Source 16: Column number × default column width
                        double colNumLeft = (colNum - 1) * 48.0;
                        sb.AppendLine($"| 16 | Col# × 48pt | {colNumLeft:F4} | N/A | {(Math.Abs(colNumLeft - dbL * pw)).ToString("F4")} | N/A | {(Math.Abs(colNumLeft - dbL * pw) < 0.01 ? "✅" : "✗")} |");

                        // Source 17: Row height sum from COM (read individual rows)
                        try
                        {
                            double comRowHeightSum = 0;
                            for (int i = 1; i < rowNum; i++)
                            {
                                dynamic rowRange = ws.Rows[i];
                                comRowHeightSum += (double)rowRange.Height;
                            }
                            sb.AppendLine($"| 17 | COM Row Height Sum | N/A | {comRowHeightSum:F4} | N/A | {(Math.Abs(comRowHeightSum - dbT * ph)).ToString("F4")} | {(Math.Abs(comRowHeightSum - dbT * ph) < 0.01 ? "✅" : "✗")} |");
                        }
                        catch { }

                        // Source 18: Column width sum from COM
                        try
                        {
                            double comColWidthSum = 0;
                            for (int i = 1; i < colNum; i++)
                            {
                                dynamic colRange = ws.Columns[i];
                                double cw = (double)colRange.ColumnWidth;
                                comColWidthSum += cw * 50.04 / 8.43;
                            }
                            sb.AppendLine($"| 18 | COM Column Width Sum | {comColWidthSum:F4} | N/A | {(Math.Abs(comColWidthSum - dbL * pw)).ToString("F4")} | N/A | {(Math.Abs(comColWidthSum - dbL * pw) < 0.01 ? "✅" : "✗")} |");
                        }
                        catch { }

                        // Source 19: Printable area (position - margins)
                        double printablePosL = rangeLeft + printableOriginX;
                        double printablePosT = rangeTop + printableOriginY;
                        sb.AppendLine($"| 19 | Printable Area Coords | {printablePosL:F4} | {printablePosT:F4} | {(Math.Abs(printablePosL - dbL * pw)).ToString("F4")} | {(Math.Abs(printablePosT - dbT * ph)).ToString("F4")} | {(Math.Abs(printablePosL - dbL * pw) < 0.01 && Math.Abs(printablePosT - dbT * ph) < 0.01 ? "✅" : "✗")} |");

                        // Source 20: DB expected absolute positions (from DB ratios)
                        double dbAbsL = dbL * pw;
                        double dbAbsT = dbT * ph;
                        sb.AppendLine($"| 20 | DB (expected) | {dbAbsL:F4} | {dbAbsT:F4} | 0.0000 | 0.0000 | REFERENCE |");

                        // Source 21: COM+Origin
                        // Derive origin from first cluster
                        if (firstAddr != null && firstAddr == addr)
                        {
                            double ox = ComCoordinateService.RoundEx(firstDbL * pw - firstComL);
                            double oy = ComCoordinateService.RoundEx(firstDbT * ph - firstComT);
                            double comPlusOriginL = rangeLeft + ox;
                            double comPlusOriginT = rangeTop + oy;
                            sb.AppendLine($"| 21 | COM + Origin ({ox:F2},{oy:F2}) | {comPlusOriginL:F4} | {comPlusOriginT:F4} | {(Math.Abs(comPlusOriginL - dbAbsL)).ToString("F4")} | {(Math.Abs(comPlusOriginT - dbAbsT)).ToString("F4")} | {(Math.Abs(comPlusOriginL - dbAbsL) < 0.01 && Math.Abs(comPlusOriginT - dbAbsT) < 0.01 ? "✅" : "✗")} |");
                        }
                        else if (firstAddr != null)
                        {
                            double ox = ComCoordinateService.RoundEx(firstDbL * pw - firstComL);
                            double oy = ComCoordinateService.RoundEx(firstDbT * ph - firstComT);
                            double comPlusOriginL = rangeLeft + ox;
                            double comPlusOriginT = rangeTop + oy;
                            sb.AppendLine($"| 21 | COM + Origin({ox:F2},{oy:F2}) | {comPlusOriginL:F4} | {comPlusOriginT:F4} | {(Math.Abs(comPlusOriginL - dbAbsL)).ToString("F4")} | {(Math.Abs(comPlusOriginT - dbAbsT)).ToString("F4")} | {(Math.Abs(comPlusOriginL - dbAbsL) < 0.01 && Math.Abs(comPlusOriginT - dbAbsT) < 0.01 ? "✅" : "✗")} |");
                        }

                        sb.AppendLine();
                        sb.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"COM error for {addr}: {ex.Message}");
                        sb.AppendLine();
                    }
                }

                // Step 5: Print entire worksheet as CSV dump
                sb.AppendLine("### Full Worksheet Range Dump (first 50 rows × first 20 columns)");
                sb.AppendLine();
                sb.AppendLine("For every cell position, dump the COM Left/Top values:");
                sb.AppendLine();
                try
                {
                    var used = ws.UsedRange;
                    int maxRow = Math.Min((int)used.Rows.Count, 50);
                    int maxCol = Math.Min((int)used.Columns.Count, 20);

                    sb.AppendLine("#### Row Heights");
                    sb.AppendLine();
                    sb.AppendLine("| Row | COM RowHeight(pt) | COM Left | COM Top |");
                    sb.AppendLine("|:---:|:-----------------:|:--------:|:-------:|");
                    for (int r = 1; r <= maxRow; r++)
                    {
                        try
                        {
                            dynamic rowRange = ws.Rows[r];
                            double rowHt = (double)rowRange.Height;
                            double rowL = (double)rowRange.Left;
                            double rowT = (double)rowRange.Top;
                            sb.AppendLine($"| {r} | {rowHt:F2} | {rowL:F2} | {rowT:F2} |");
                        }
                        catch { }
                    }
                    sb.AppendLine();

                    sb.AppendLine("#### Column Widths");
                    sb.AppendLine();
                    sb.AppendLine("| Col | COM ColumnWidth(chars) | COM Left | COM Top |");
                    sb.AppendLine("|:---:|:----------------------:|:--------:|:-------:|");
                    for (int c = 1; c <= maxCol; c++)
                    {
                        try
                        {
                            dynamic colRange = ws.Columns[c];
                            double colW = (double)colRange.ColumnWidth;
                            double colL = (double)colRange.Left;
                            double colT = (double)colRange.Top;
                            string letter = ColNumToLetters(c);
                            sb.AppendLine($"| {c}({letter}) | {colW:F3} | {colL:F2} | {colT:F2} |");
                        }
                        catch { }
                    }
                    sb.AppendLine();

                    sb.AppendLine("#### Every Cell Position Grid");
                    sb.AppendLine();
                    sb.AppendLine("| Row | Data |");
                    sb.AppendLine("|:---:|------|");
                    for (int r = 1; r <= maxRow; r++)
                    {
                        try
                        {
                            dynamic rowRange = ws.Rows[r];
                            double rowHt = (double)rowRange.Height;
                            double rowTop = (double)rowRange.Top;
                            string line = $"| {r} | Ht={rowHt:F2} Top={rowTop:F2} | Cols: ";
                            for (int c = 1; c <= maxCol; c++)
                            {
                                try
                                {
                                    dynamic cell = ws.Cells[r, c];
                                    double cellL = (double)cell.Left;
                                    double cellW = (double)cell.Width;
                                    line += $"C{c} L={cellL:F1} W={cellW:F1} ";
                                }
                                catch { }
                            }
                            sb.AppendLine(line);
                        }
                        catch { }
                    }
                    sb.AppendLine();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"UsedRange dump failed: {ex.Message}");
                    sb.AppendLine();
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

        return sb.ToString();
    }

    private OpenXmlInfo ReadOpenXmlInfo(string workbookPath)
    {
        var info = new OpenXmlInfo();
        try
        {
            using var doc = SpreadsheetDocument.Open(workbookPath, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return info;

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
        }
        catch { }
        return info;
    }

    private static double ParseDouble(string? s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;

    private static string ColNumToLetters(int col)
    {
        string letters = "";
        while (col > 0)
        {
            col--;
            letters = (char)('A' + col % 26) + letters;
            col /= 26;
        }
        return letters;
    }

    private static double SafeGetDouble(Func<double> getter, double fallback = double.NaN)
    {
        try { return getter(); } catch { return fallback; }
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

// OpenXmlInfo is defined in Phase11ReverseEngineer.cs
