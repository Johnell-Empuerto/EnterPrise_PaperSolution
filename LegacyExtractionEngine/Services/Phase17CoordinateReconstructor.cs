using System.Runtime.InteropServices;
using System.Text;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 17 — Reconstruct the Legacy Coordinate Algorithm.
/// 
/// Based on decompiled Cimtops.Excel.dll and Cimtops.R2Cluster.dll:
/// 
/// The legacy engine does NOT use Range.Left/Range.Top for coordinate computation.
/// Instead it:
///   1. Reads the PrintArea from PageSetup
///   2. Iterates every column in the PrintArea → stores width in points
///   3. Iterates every row in the PrintArea → stores height in points
///   4. For each cluster, computes position as:
///        left_pt = sum(Col[n].Width) for all n < cluster.startColumn
///        top_pt  = sum(Row[n].Height) for all n < cluster.startRow
///        width_pt  = sum(Col[n].Width) for cluster.startColumn..cluster.endColumn
///        height_pt = sum(Row[n].Height) for cluster.startRow..cluster.endRow
///   5. Normalizes by page dimensions (PageSetup.PageWidth/PageHeight)
///   6. Applies RoundEx (banker's rounding to 7 decimal places)
/// </summary>
public class Phase17CoordinateReconstructor : IDisposable
{
    private readonly IProgress<string> _progress;
    private dynamic? _excel;
    private dynamic? _workbook;
    private bool _disposed;

    public const double DefaultPageWidth = 612.0;
    public const double DefaultPageHeight = 792.0;

    public Phase17CoordinateReconstructor(IProgress<string> progress)
    {
        _progress = progress;
    }

    public async Task RunAsync()
    {
        _progress.Report("============================================");
        _progress.Report(" Phase 17 — Legacy Coordinate Reconstruction");
        _progress.Report(" Column-Width/Row-Height Summation Algorithm");
        _progress.Report("============================================");
        _progress.Report("");

        var templates = new (int Id, string Path, string Label)[]
        {
            (546, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx", "FormTest - Copy"),
            (547, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx", "[V3.1_Sample]アンケート用紙"),
            (548, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx", "Sample A"),
        };

        var outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test", "Phase17_Reconstruction");
        Directory.CreateDirectory(outputDir);

        var report = new StringBuilder();
        report.AppendLine("# Phase 17 — Legacy Coordinate Reconstruction Report");
        report.AppendLine();
        report.AppendLine("**Algorithm:** Column-Width/Row-Height Summation (from decompiled Cimtops.Excel.dll)");
        report.AppendLine("**Date:** " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        report.AppendLine();
        report.AppendLine("## Formula");
        report.AppendLine();
        report.AppendLine("```");
        report.AppendLine("left_pt   = Σ Col[n].Width  for n < cluster.StartColumn");
        report.AppendLine("top_pt    = Σ Row[n].Height for n < cluster.StartRow");
        report.AppendLine("width_pt  = Σ Col[n].Width  for n = StartColumn..EndColumn");
        report.AppendLine("height_pt = Σ Row[n].Height for n = StartRow..EndRow");
        report.AppendLine("ratio     = RoundEx((pos_pt + origin_pt) / page_dim_pt, 7)");
        report.AppendLine("```");
        report.AppendLine();
        report.AppendLine("**Key differences from Phase 10:**");
        report.AppendLine("- Position = sum of column widths / row heights (NOT Range.Left/Range.Top)");
        report.AppendLine("- Width/Height = sum of column widths / row heights for the spanned range");
        report.AppendLine("- Normalization divisor = PageSetup.PageWidth/PageHeight");
        report.AppendLine();
        report.AppendLine("---");
        report.AppendLine();

        var allResults = new List<(int Id, List<(string Cell, double DbL, double DbT, double DbR, double DbB,
            double GenL, double GenT, double GenR, double GenB, bool L, bool T, bool R, bool B)>)>();

        foreach (var (id, path, label) in templates)
        {
            if (!File.Exists(path))
            {
                _progress.Report($"[SKIP] Template {id}: File not found: {path}");
                report.AppendLine($"## Template {id} — {label}");
                report.AppendLine();
                report.AppendLine($"**File not found:** `{path}`");
                report.AppendLine();
                continue;
            }

            _progress.Report($"=== Template {id}: {label} ===");
            var result = await AnalyzeTemplate(id, path, outputDir);
            allResults.Add((id, result));
        }

        // Generate report
        report.AppendLine("## Per-Template Detail");
        report.AppendLine();

        foreach (var (id, clusters) in allResults)
        {
            report.AppendLine($"### Template {id}");
            report.AppendLine();
            report.AppendLine("| Cluster | DB L | DB T | DB R | DB B | Gen L | Gen T | Gen R | Gen B | L? | T? | R? | B? |");
            report.AppendLine("|---------|------|------|------|------|-------|-------|-------|-------|:--:|:--:|:--:|:--:|");

            int lPass = 0, tPass = 0, rPass = 0, bPass = 0;
            int total = 0;

            foreach (var c in clusters)
            {
                total++;
                if (c.L) lPass++;
                if (c.T) tPass++;
                if (c.R) rPass++;
                if (c.B) bPass++;

                report.AppendLine($"| {c.Cell} | {c.DbL:F7} | {c.DbT:F7} | {c.DbR:F7} | {c.DbB:F7} | " +
                    $"{c.GenL:F7} | {c.GenT:F7} | {c.GenR:F7} | {c.GenB:F7} | " +
                    $"{(c.L ? "✅" : "❌")} | {(c.T ? "✅" : "❌")} | {(c.R ? "✅" : "❌")} | {(c.B ? "✅" : "❌")} |");
            }

            report.AppendLine();
            report.AppendLine($"**Results:** Left={lPass}/{total} ({100.0 * lPass / total:F0}%), " +
                $"Top={tPass}/{total} ({100.0 * tPass / total:F0}%), " +
                $"Right={rPass}/{total} ({100.0 * rPass / total:F0}%), " +
                $"Bottom={bPass}/{total} ({100.0 * bPass / total:F0}%)");
            report.AppendLine();
            report.AppendLine("---");
            report.AppendLine();
        }

        // Cross-template summary
        report.AppendLine("## Cross-Template Summary");
        report.AppendLine();
        report.AppendLine("| Template | Left% | Top% | Right% | Bottom% | Overall |");
        report.AppendLine("|----------|:----:|:----:|:------:|:-------:|:-------:|");

        foreach (var (id, clusters) in allResults)
        {
            int total = clusters.Count;
            int l = clusters.Count(c => c.L);
            int t = clusters.Count(c => c.T);
            int r = clusters.Count(c => c.R);
            int b = clusters.Count(c => c.B);
            int all = clusters.Count(c => c.L && c.T && c.R && c.B);
            report.AppendLine($"| {id} | {100.0 * l / total:F1}% | {100.0 * t / total:F1}% | " +
                $"{100.0 * r / total:F1}% | {100.0 * b / total:F1}% | {100.0 * all / total:F1}% |");
        }

        report.AppendLine();
        report.AppendLine("## Conclusion");
        report.AppendLine();
        report.AppendLine("The column-width/row-height summation algorithm was implemented and tested against");
        report.AppendLine("all three reference templates. See per-template detail above for exact match results.");
        report.AppendLine();
        report.AppendLine("**If Left/Top/Right/Bottom do not all match at 100%:**");
        report.AppendLine("");
        report.AppendLine("1. Check if the PrintArea boundary was correctly identified");
        report.AppendLine("2. Check if column widths and row heights from COM match OpenXML values");
        report.AppendLine("3. Check if the printed origin offset (page margins + centering) is correct");
        report.AppendLine("4. The publishing code in ConMasClient.exe may apply additional transforms");

        var reportPath = Path.Combine(outputDir, "Phase17Report.md");
        await File.WriteAllTextAsync(reportPath, report.ToString());
        _progress.Report($"Report written: {reportPath}");
    }

    private async Task<List<(string Cell, double DbL, double DbT, double DbR, double DbB,
        double GenL, double GenT, double GenR, double GenB, bool L, bool T, bool R, bool B)>>
        AnalyzeTemplate(int defTopId, string workbookPath, string outputDir)
    {
        var results = new List<(string, double, double, double, double,
            double, double, double, double, bool, bool, bool, bool)>();

        try
        {
            // Load DB clusters
            _progress.Report("  Loading DB clusters...");
            string connStr = "Host=localhost;Port=5432;Database=irepodb;Username=postgres;Password=cimtops";
            var dbReader = new DatabaseReader(connStr, _progress);
            var db = await dbReader.ReadAllAsync(defTopId);
            if (db?.DefSheets == null || db.DefSheets.Count == 0)
            {
                _progress.Report($"  SKIP: No DB data for template {defTopId}");
                return results;
            }

            var dbClusters = db.DefSheets.SelectMany(s => s.Clusters).ToList();
            _progress.Report($"  Found {dbClusters.Count} DB clusters");

            // Open workbook via COM
            _progress.Report("  Opening Excel via COM...");
            OpenWorkbook(workbookPath);
            dynamic ws = _workbook.Sheets[1];
            dynamic pageSetup = ws.PageSetup;

            // Read page setup
            double pageWidth = SafeReadDouble(() => (double)pageSetup.PageWidth, DefaultPageWidth);
            double pageHeight = SafeReadDouble(() => (double)pageSetup.PageHeight, DefaultPageHeight);
            if (pageWidth < 1) pageWidth = DefaultPageWidth;
            if (pageHeight < 1) pageHeight = DefaultPageHeight;

            // Read PrintArea from COM
            string? comPrintArea = SafeReadString(() => (string)pageSetup.PrintArea);
            _progress.Report($"  PageSetup.PageWidth={pageWidth:F1}, PageHeight={pageHeight:F1}");
            _progress.Report($"  PrintArea (COM): {comPrintArea}");

            // Get the PrintArea range
            string? rangeAddress = comPrintArea;
            if (string.IsNullOrEmpty(rangeAddress))
            {
                // Fallback: use entire used range
                dynamic usedRange = ws.UsedRange;
                rangeAddress = (string)usedRange.Address;
                _progress.Report($"  No PrintArea from COM — falling back to UsedRange: {rangeAddress}");
            }

            // Parse print area to get start/end rows and columns
            // Strip sheet name prefix before !
            if (rangeAddress.Contains("!"))
                rangeAddress = rangeAddress.Substring(rangeAddress.IndexOf('!') + 1);

            dynamic paRange = ws.Range[rangeAddress];
            int startRow = paRange.Row;
            int endRow = paRange.Row + paRange.Rows.Count - 1;
            int startCol = paRange.Column;
            int endCol = paRange.Column + paRange.Columns.Count - 1;
            _progress.Report($"  PrintArea range: rows={startRow}-{endRow}, cols={startCol}-{endCol}");

            // Read column widths from COM using Range.Width (points), not ColumnWidth (char units)
            // The decompiled Sheet.cs uses: obj2.Width (points), NOT obj2.ColumnWidth (characters)
            var colWidths = new Dictionary<int, double>();
            _progress.Report("  Reading column widths (via Range.Width in points)...");
            for (int c = startCol; c <= endCol; c++)
            {
                try
                {
                    dynamic colRange = ws.Columns[c];
                    colWidths[c] = (double)colRange.Width;  // Range.Width returns points!
                }
                catch
                {
                    // Fallback: ~48pt for default Calibri 11pt width
                    colWidths[c] = 48.0;
                }
            }

            // Read row heights from COM
            var rowHeights = new Dictionary<int, double>();
            _progress.Report("  Reading row heights...");
            for (int r = startRow; r <= endRow; r++)
            {
                try
                {
                    dynamic rowRange = ws.Rows[r];
                    rowHeights[r] = (double)rowRange.RowHeight;
                }
                catch
                {
                    rowHeights[r] = 15.0; // default
                }
            }

            // Log a few sample values
            _progress.Report($"  Sample row heights: R{startRow}={rowHeights[startRow]:F2}, " +
                $"R{startRow + 1}={rowHeights.GetValueOrDefault(startRow + 1, 0):F2}, ...");
            _progress.Report($"  Sample col widths: C{startCol}={colWidths[startCol]:F2}, " +
                $"C{startCol + 1}={colWidths.GetValueOrDefault(startCol + 1, 0):F2}, ...");

            // Compute total print area dimensions (sum of all columns/rows in print area)
            double totalColWidthPt = 0;
            for (int c = startCol; c <= endCol; c++)
                totalColWidthPt += colWidths.GetValueOrDefault(c, 48.0);

            double totalRowHeightPt = 0;
            for (int r = startRow; r <= endRow; r++)
                totalRowHeightPt += rowHeights.GetValueOrDefault(r, 14.4);

            double sheetArea = totalColWidthPt * totalRowHeightPt;
            _progress.Report($"  PrintArea total col width: {totalColWidthPt:F2}pt, total row height: {totalRowHeightPt:F2}pt, " +
                $"Sheet area: {sheetArea:F2} pts²");

            // Get first DB cluster to compute printed origin
            var firstDb = dbClusters.FirstOrDefault();
            double firstComLeftPt = 0, firstComTopPt = 0;
            double firstDbLeft = 0, firstDbTop = 0;
            double originX = 0, originY = 0;

            if (firstDb != null)
            {
                var addr = firstDb.CellAddr?.Trim() ?? "";
                var (clStartCol, clStartRow, clEndCol, clEndRow) = ParseAddress(addr);

                // Compute position in points
                firstComLeftPt = SumColWidths(colWidths, startCol, clStartCol - 1);
                firstComTopPt = SumRowHeights(rowHeights, startRow, clStartRow - 1);

                firstDbLeft = ParseDouble(firstDb.LeftPosition);
                firstDbTop = ParseDouble(firstDb.TopPosition);

                // Derive origin: DB_ratio = (pos_pt + origin_pt) / page_dim
                // => origin_pt = DB_ratio * page_dim - pos_pt
                originX = firstDbLeft * pageWidth - firstComLeftPt;
                originY = firstDbTop * pageHeight - firstComTopPt;

                _progress.Report($"  First cluster: {addr}");
                _progress.Report($"    COM: col={clStartCol}, row={clStartRow}");
                _progress.Report($"    leftPt={firstComLeftPt:F2}, topPt={firstComTopPt:F2}");
                _progress.Report($"    DB: left={firstDbLeft:F7}, top={firstDbTop:F7}");
                _progress.Report($"    Derived origin: X={originX:F2}pt, Y={originY:F2}pt");
            }

            // Process each DB cluster
            foreach (var dbc in dbClusters)
            {
                var addr = dbc.CellAddr?.Trim() ?? "";
                if (string.IsNullOrEmpty(addr)) continue;

                var (clStartCol, clStartRow, clEndCol, clEndRow) = ParseAddress(addr);

                // Phase 1: Try MergeArea if available
                int actualStartCol = clStartCol, actualEndCol = clEndCol;
                int actualStartRow = clStartRow, actualEndRow = clEndRow;

                try
                {
                    dynamic range = ws.Range[addr];
                    try
                    {
                        dynamic merge = range.MergeArea;
                        if (merge != null)
                        {
                            actualStartCol = merge.Column;
                            actualStartRow = merge.Row;
                            actualEndCol = actualStartCol + merge.Columns.Count - 1;
                            actualEndRow = actualStartRow + merge.Rows.Count - 1;
                        }
                    }
                    catch { }
                }
                catch { }

                // Compute position as sum of column widths / row heights
                double leftPt = SumColWidths(colWidths, startCol, actualStartCol - 1);
                double topPt = SumRowHeights(rowHeights, startRow, actualStartRow - 1);
                double widthPt = SumColWidths(colWidths, actualStartCol, actualEndCol);
                double heightPt = SumRowHeights(rowHeights, actualStartRow, actualEndRow);

                // Compute ratios using page dimensions
                double genL = ComputeRatio(leftPt + originX, pageWidth);
                double genT = ComputeRatio(topPt + originY, pageHeight);
                double genR = ComputeRatio(leftPt + widthPt + originX, pageWidth);
                double genB = ComputeRatio(topPt + heightPt + originY, pageHeight);

                // DB values
                double dbL = ParseDouble(dbc.LeftPosition);
                double dbT = ParseDouble(dbc.TopPosition);
                double dbR = ParseDouble(dbc.RightPosition);
                double dbB = ParseDouble(dbc.BottomPosition);

                // Compare
                const double tolerance = 0.000001;
                bool lOk = Math.Abs(genL - dbL) < tolerance;
                bool tOk = Math.Abs(genT - dbT) < tolerance;
                bool rOk = Math.Abs(genR - dbR) < tolerance;
                bool bOk = Math.Abs(genB - dbB) < tolerance;

                _progress.Report($"    {addr}: L={genL:F7}(db={dbL:F7}) {(lOk ? "✅" : "❌")}, " +
                    $"T={genT:F7}(db={dbT:F7}) {(tOk ? "✅" : "❌")}, " +
                    $"R={genR:F7}(db={dbR:F7}) {(rOk ? "✅" : "❌")}, " +
                    $"B={genB:F7}(db={dbB:F7}) {(bOk ? "✅" : "❌")}");

                results.Add((addr, dbL, dbT, dbR, dbB, genL, genT, genR, genB, lOk, tOk, rOk, bOk));
            }
        }
        catch (Exception ex)
        {
            _progress.Report($"  ERROR: {ex.GetType().Name}: {ex.Message}");
            _progress.Report($"  Stack: {ex.StackTrace}");
        }

        return results;
    }

    private void OpenWorkbook(string path)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new InvalidOperationException("Excel COM not available");

        _excel = Activator.CreateInstance(excelType);
        _excel.Visible = false;
        _excel.DisplayAlerts = false;
        _excel.ScreenUpdating = false;
        _workbook = _excel.Workbooks.Open(path);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_workbook != null) { _workbook.Close(false); Marshal.ReleaseComObject(_workbook); }
            if (_excel != null) { _excel.Quit(); Marshal.ReleaseComObject(_excel); }
        }
        catch { }
    }

    // ========== Helpers ==========

    private static double SumColWidths(Dictionary<int, double> colWidths, int fromCol, int toCol)
    {
        double sum = 0;
        for (int c = fromCol; c <= toCol; c++)
            sum += colWidths.GetValueOrDefault(c, 48.0);
        return sum;
    }

    private static double SumRowHeights(Dictionary<int, double> rowHeights, int fromRow, int toRow)
    {
        double sum = 0;
        for (int r = fromRow; r <= toRow; r++)
            sum += rowHeights.GetValueOrDefault(r, 15.0);
        return sum;
    }

    /// <summary>
    /// Parse an Excel address like "$A$1:$B$2" or "$A$1" into column/row numbers.
    /// </summary>
    private static (int startCol, int startRow, int endCol, int endRow) ParseAddress(string addr)
    {
        try
        {
            addr = addr.Replace("$", ""); // Strip absolute ref markers
            var parts = addr.Split(':');
            var (startCol, startRow) = ParseCellRef(parts[0]);
            if (parts.Length >= 2)
            {
                var (endCol, endRow) = ParseCellRef(parts[1]);
                return (startCol, startRow, endCol, endRow);
            }
            return (startCol, startRow, startCol, startRow);
        }
        catch
        {
            return (1, 1, 1, 1);
        }
    }

    private static (int col, int row) ParseCellRef(string ref_)
    {
        ref_ = ref_.Trim();
        int i = 0;
        // Extract column letters
        while (i < ref_.Length && char.IsLetter(ref_[i])) i++;
        var colStr = ref_.Substring(0, i);
        var rowStr = ref_.Substring(i);

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        int col = 0;
        foreach (char c in colStr)
        {
            col = col * 26 + (char.ToUpper(c) - 'A' + 1);
        }

        int row = int.TryParse(rowStr, out var r) ? r : 1;
        return (col, row);
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static double ComputeRatio(double valuePt, double pageDim)
    {
        if (pageDim <= 0) return 0;
        return Math.Round((float)(valuePt / pageDim), 7, MidpointRounding.AwayFromZero);
    }

    private static double SafeReadDouble(Func<double> getter, double fallback = 0)
    {
        try { return getter(); }
        catch { return fallback; }
    }

    private static string SafeReadString(Func<string> getter)
    {
        try { return getter() ?? ""; }
        catch { return ""; }
    }
}
