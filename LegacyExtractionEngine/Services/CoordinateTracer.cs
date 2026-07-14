using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Uses Excel COM Interop to read actual rendered cell positions (Range.Left/Top/Width/Height)
/// and reverse-engineers the coordinate transformation that the legacy VB6 importer used.
/// 
/// The legacy VB6 code used Excel's COM API to get final rendered positions after
/// print scaling, zoom, margins, and centering were applied. This tracer extracts
/// those positions and compares them with both OpenXML-calculated and DB-stored values
/// to determine the exact transformation formula.
/// </summary>
public class CoordinateTracer
{
    private readonly IProgress<string> _progress;
    private readonly string _connectionString;
    private readonly string _outputDir;

    public CoordinateTracer(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test");
    }

    public async Task TraceAllAsync()
    {
        _progress.Report("=== Coordinate System Investigation ===");

        // Template 546 (simple, already 100%)
        await TraceTemplateAsync(546,
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx",
            "Template546");

        // Template 547 (complex, positions differ)
        var wb547 = @"C:\Users\MCF-JOHNELLEEMPUERTO\AppData\Local\Temp\opencode\547_workbook_copy.xlsx";
        if (!File.Exists(wb547))
        {
            var dir = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents";
            wb547 = Directory.GetFiles(dir, "*V3.1*", SearchOption.TopDirectoryOnly).FirstOrDefault() ?? "";
        }
        if (File.Exists(wb547))
            await TraceTemplateAsync(547, wb547, "Template547");

        _progress.Report("=== Investigation Complete ===");
    }

    private async Task TraceTemplateAsync(int defTopId, string workbookPath, string templateDir)
    {
        _progress.Report($"--- Template {defTopId} ({templateDir}) ---");

        // Load DB data
        var dbReader = new DatabaseReader(_connectionString, _progress);
        var db = await dbReader.ReadAllAsync(defTopId);
        if (db == null) { _progress.Report("  SKIP: No DB data"); return; }

        var clusters = db.DefSheets.SelectMany(s => s.Clusters).ToList();
        if (clusters.Count == 0) { _progress.Report("  SKIP: No DB clusters"); return; }

        var outputDir = Path.Combine(_outputDir, templateDir);
        Directory.CreateDirectory(outputDir);

        // Try to open workbook via Excel COM
        var traceResults = new List<CoordinateTrace>();
        dynamic? excel = null;
        dynamic? wb = null;

        try
        {
            _progress.Report("  Opening Excel via COM...");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                _progress.Report("  ERROR: Excel COM not available");
                return;
            }
            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;

            // Open workbook
            _progress.Report($"  Opening workbook: {Path.GetFileName(workbookPath)}");
            wb = excel.Workbooks.Open(workbookPath);

            // Get page setup info
            dynamic ws = wb.Sheets[1];
            dynamic pageSetup = ws.PageSetup;

            // Read page setup properties
            double pageWidth = 612;  // default letter
            double pageHeight = 792;
            try { pageWidth = (double)pageSetup.PageWidth; } catch { }
            try { pageHeight = (double)pageSetup.PageHeight; } catch { }

            var psInfo = new PageSetupInfo
            {
                PaperSize = (int?)pageSetup.PaperSize,
                Zoom = (int?)pageSetup.Zoom,
                FitToPagesWide = (int?)pageSetup.FitToPagesWide,
                FitToPagesTall = (int?)pageSetup.FitToPagesTall,
                LeftMargin = (double)pageSetup.LeftMargin,
                RightMargin = (double)pageSetup.RightMargin,
                TopMargin = (double)pageSetup.TopMargin,
                BottomMargin = (double)pageSetup.BottomMargin,
                HeaderMargin = (double)pageSetup.HeaderMargin,
                FooterMargin = (double)pageSetup.FooterMargin,
                CenterHorizontally = (bool)pageSetup.CenterHorizontally,
                CenterVertically = (bool)pageSetup.CenterVertically,
                Orientation = (int?)pageSetup.Orientation,
                PageWidth = pageWidth,
                PageHeight = pageHeight,
                PrintArea = (string?)pageSetup.PrintArea,
                Order = (int?)pageSetup.Order,
                FirstPageNumber = (int?)pageSetup.FirstPageNumber,
                // Additional scaling info
                PrintQuality = pageSetup.PrintQuality is int[] pq ? pq : null,
            };

            // Log page setup
            _progress.Report($"  Page: {pageWidth:F1}x{pageHeight:F1}, Zoom={psInfo.Zoom}, " +
                $"FitToPages={psInfo.FitToPagesWide}x{psInfo.FitToPagesTall}, " +
                $"Margins(L,R,T,B): {psInfo.LeftMargin:F3}, {psInfo.RightMargin:F3}, " +
                $"{psInfo.TopMargin:F3}, {psInfo.BottomMargin:F3}");

            // Get the document's UsedRange dimensions from the first sheet
            dynamic sheet = wb.Sheets[1];
            dynamic usedRange = sheet.UsedRange;
            psInfo.UsedRangeRows = usedRange.Rows.Count;
            psInfo.UsedRangeColumns = usedRange.Columns.Count;

            // Read page breaks
            try { psInfo.HPageBreaks = sheet.HPageBreaks.Count; } catch { }
            try { psInfo.VPageBreaks = sheet.VPageBreaks.Count; } catch { }

            // Process each DB cluster
            foreach (var dbCluster in clusters)
            {
                var addr = dbCluster.CellAddr?.Trim() ?? "";
                if (string.IsNullOrEmpty(addr)) continue;

                _progress.Report($"  Tracing {addr}...");

                // Read COM positions
                dynamic range = sheet.Range[addr];
                double comLeft = 0, comTop = 0, comWidth = 0, comHeight = 0;
                try { comLeft = (double)range.Left; } catch { }
                try { comTop = (double)range.Top; } catch { }
                try { comWidth = (double)range.Width; } catch { }
                try { comHeight = (double)range.Height; } catch { }

                // Parse DB position ratios
                double dbLeft = ParseDouble(dbCluster.LeftPosition);
                double dbTop = ParseDouble(dbCluster.TopPosition);
                double dbRight = ParseDouble(dbCluster.RightPosition);
                double dbBottom = ParseDouble(dbCluster.BottomPosition);

                // Try different coordinate transformations
                var trace = new CoordinateTrace
                {
                    Address = addr,
                    // COM workbook coordinates (points from top-left of worksheet)
                    ComLeft = comLeft,
                    ComTop = comTop,
                    ComWidth = comWidth,
                    ComHeight = comHeight,
                    // DB stored ratios
                    DbLeft = dbLeft,
                    DbTop = dbTop,
                    DbRight = dbRight,
                    DbBottom = dbBottom,
                    // Page setup info
                    PageWidth = pageWidth,
                    PageHeight = pageHeight,
                    LeftMargin = psInfo.LeftMargin,
                    RightMargin = psInfo.RightMargin,
                    TopMargin = psInfo.TopMargin,
                    BottomMargin = psInfo.BottomMargin,
                    Zoom = psInfo.Zoom,
                    FitToPagesWide = psInfo.FitToPagesWide,
                    FitToPagesTall = psInfo.FitToPagesTall,
                    CenterHorizontally = psInfo.CenterHorizontally,
                    CenterVertically = psInfo.CenterVertically,
                };

                // Try transformations
                trace = ComputeTransformations(trace, psInfo);
                traceResults.Add(trace);
            }

            // Close workbook
            wb.Close(false);
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(sheet);

            _progress.Report($"  Traced {traceResults.Count} clusters");
        }
        catch (Exception ex)
        {
            _progress.Report($"  COM Error: {ex.Message}");
        }
        finally
        {
            // Cleanup COM
            if (wb != null) try { wb.Close(false); } catch { }
            if (excel != null) try { excel.Quit(); } catch { }
        }

        // Generate analysis documents
        await GenerateClusterCoordinateTraceAsync(traceResults, clusters, outputDir);
        await GenerateCoordinateSystemComparisonAsync(traceResults, outputDir);
        await GenerateCoordinateMathAsync(traceResults, outputDir);
        await GenerateLegacyAlgorithmAsync(traceResults, outputDir);
    }

    private CoordinateTrace ComputeTransformations(CoordinateTrace t, PageSetupInfo ps)
    {
        // ===== Transformation 1: Raw workbook coordinates → page ratio =====
        // Simple: COM_Left / PageWidth
        t.RawLeftRatio = t.ComWidth > 0 ? SafeRatio(t.ComLeft, t.PageWidth) : 0;
        t.RawTopRatio = t.ComHeight > 0 ? SafeRatio(t.ComTop, t.PageHeight) : 0;
        t.RawWidthRatio = SafeRatio(t.ComWidth, t.PageWidth);
        t.RawHeightRatio = SafeRatio(t.ComHeight, t.PageHeight);
        t.RawRightRatio = t.RawLeftRatio + t.RawWidthRatio;
        t.RawBottomRatio = t.RawTopRatio + t.RawHeightRatio;

        // ===== Transformation 2: Printable area coordinates =====
        // Subtract margins from page, compute position relative to printable area
        double printableWidth = t.PageWidth - ps.LeftMargin - ps.RightMargin;
        double printableHeight = t.PageHeight - ps.TopMargin - ps.BottomMargin;
        t.PrintableLeftRatio = SafeRatio(t.ComLeft - ps.LeftMargin, printableWidth);
        t.PrintableTopRatio = SafeRatio(t.ComTop - ps.TopMargin, printableHeight);
        t.PrintableRightRatio = t.PrintableLeftRatio + SafeRatio(t.ComWidth, printableWidth);
        t.PrintableBottomRatio = t.PrintableTopRatio + SafeRatio(t.ComHeight, printableHeight);

        // ===== Transformation 3: Scaled printable area =====
        // If FitToPages is set, scale the printable coordinates
        double scaleX = 1.0;
        double scaleY = 1.0;

        if (t.FitToPagesWide.HasValue && t.FitToPagesWide > 0)
        {
            // The content is scaled to fit t.FitToPagesWide pages
            // One page = printableWidth. t.FitToPagesWide pages = t.FitToPagesWide * printableWidth
            // Scale = printableWidth / (t.FitToPagesWide * printableWidth) = 1/t.FitToPagesWide
            // But actually, Excel scales so that content fills N pages wide
            scaleX = 1.0 / t.FitToPagesWide.Value;
        }
        if (t.FitToPagesTall.HasValue && t.FitToPagesTall > 0)
        {
            scaleY = 1.0 / t.FitToPagesTall.Value;
        }

        t.ScaleX = scaleX;
        t.ScaleY = scaleY;

        var scaledPrintableW = printableWidth;
        var scaledPrintableH = printableHeight;

        // Scaled renders: COM position in scaled printable area
        t.ScaledLeftRatio = SafeRatio(t.ComLeft - ps.LeftMargin, scaledPrintableW);
        t.ScaledTopRatio = SafeRatio(t.ComTop - ps.TopMargin, scaledPrintableH);

        // ===== Transformation 4: Centered+Margins =====
        // The "PrintedOrigin" = where the content starts on the printed page
        // If centered: PrintedOrigin = (PageWidth - ContentWidth) / 2
        // If not centered: PrintedOrigin = LeftMargin
        double printedOriginX = 0;
        double printedOriginY = 0;

        if (t.CenterHorizontally)
        {
            // The content area is centered on the page
            // PrintedOrigin = (PageWidth - printableWidth * scale) / 2
            printedOriginX = (t.PageWidth - printableWidth) / 2.0;
        }
        else
        {
            printedOriginX = ps.LeftMargin;
        }

        if (t.CenterVertically)
        {
            printedOriginY = (t.PageHeight - printableHeight) / 2.0;
        }
        else
        {
            printedOriginY = ps.TopMargin;
        }

        t.PrintedOriginX = printedOriginX;
        t.PrintedOriginY = printedOriginY;

        // ===== Transformation 5: Position relative to printed origin =====
        // Takes the COM position, subtracts the printed origin, divides by printable area
        t.OriginLeftRatio = SafeRatio(t.ComLeft - printedOriginX, printableWidth);
        t.OriginTopRatio = SafeRatio(t.ComTop - printedOriginY, printableHeight);

        // ===== Transformation 6: Page-left-margin ratio =====
        // Position relative to left margin, expressed as ratio of printable area
        // This is: (COM_Left - LeftMargin) / printableWidth
        // = printableLeftRatio (already computed above)

        // ===== Transformation 7: Right-side centered =====
        // Alternative: maybe the legacy code used right margin for right side
        t.MarginOnlyLeftRatio = SafeRatio(t.ComLeft, t.PageWidth - ps.RightMargin);
        t.MarginOnlyTopRatio = SafeRatio(t.ComTop, t.PageHeight - ps.BottomMargin);

        // ===== Transformation 8: Content-scaled =====
        // If scale is applied: COM position / (page * scale)
        if (scaleX > 0 && scaleX != 1.0)
        {
            t.ScaledLeftRatio = SafeRatio(t.ComLeft, t.PageWidth * scaleX);
            t.ScaledTopRatio = SafeRatio(t.ComTop, t.PageHeight * scaleY);
        }

        // ===== Compute differences from DB values =====
        t.DiffRawLeft = Math.Abs(t.RawLeftRatio - t.DbLeft);
        t.DiffRawTop = Math.Abs(t.RawTopRatio - t.DbTop);
        t.DiffPrintableLeft = Math.Abs(t.PrintableLeftRatio - t.DbLeft);
        t.DiffPrintableTop = Math.Abs(t.PrintableTopRatio - t.DbTop);
        t.DiffOriginLeft = Math.Abs(t.OriginLeftRatio - t.DbLeft);
        t.DiffOriginTop = Math.Abs(t.OriginTopRatio - t.DbTop);
        t.DiffMarginLeft = Math.Abs(t.MarginOnlyLeftRatio - t.DbLeft);
        t.DiffMarginTop = Math.Abs(t.MarginOnlyTopRatio - t.DbTop);

        // Find best match
        var diffs = new Dictionary<string, double>
        {
            ["Raw"] = t.DiffRawLeft + t.DiffRawTop,
            ["Printable"] = t.DiffPrintableLeft + t.DiffPrintableTop,
            ["Origin"] = t.DiffOriginLeft + t.DiffOriginTop,
            ["MarginOnly"] = t.DiffMarginLeft + t.DiffMarginTop,
        };
        t.BestTransform = diffs.OrderBy(d => d.Value).First().Key;
        t.BestDiff = diffs.Min(d => d.Value);

        return t;
    }

    private async Task GenerateClusterCoordinateTraceAsync(List<CoordinateTrace> traces, List<DefCluster> dbClusters, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Cluster Coordinate Trace");
        sb.AppendLine();
        sb.AppendLine("For every DB cluster, this trace shows:");
        sb.AppendLine("- COM Range.Left/Top (Excel's rendered position)");
        sb.AppendLine("- DB stored ratios");
        sb.AppendLine("- Multiple transformation attempts");
        sb.AppendLine("- Best matching transformation");
        sb.AppendLine();

        foreach (var t in traces.OrderBy(x => x.Address))
        {
            sb.AppendLine($"## {t.Address}");
            sb.AppendLine();
            sb.AppendLine("### Workbook Positions (from Excel COM)");
            sb.AppendLine($"| Property | Points |");
            sb.AppendLine($"|----------|-------:|");
            sb.AppendLine($"| Range.Left | {t.ComLeft:F4} |");
            sb.AppendLine($"| Range.Top | {t.ComTop:F4} |");
            sb.AppendLine($"| Range.Width | {t.ComWidth:F4} |");
            sb.AppendLine($"| Range.Height | {t.ComHeight:F4} |");
            sb.AppendLine();
            sb.AppendLine("### Page Setup");
            sb.AppendLine($"| Property | Value |");
            sb.AppendLine($"|----------|------:|");
            sb.AppendLine($"| Page Width | {t.PageWidth:F1} pt |");
            sb.AppendLine($"| Page Height | {t.PageHeight:F1} pt |");
            sb.AppendLine($"| Left Margin | {t.LeftMargin:F4} pt |");
            sb.AppendLine($"| Right Margin | {t.RightMargin:F4} pt |");
            sb.AppendLine($"| Top Margin | {t.TopMargin:F4} pt |");
            sb.AppendLine($"| Bottom Margin | {t.BottomMargin:F4} pt |");
            sb.AppendLine($"| Center Horizontally | {t.CenterHorizontally} |");
            sb.AppendLine($"| Center Vertically | {t.CenterVertically} |");
            sb.AppendLine($"| Zoom | {t.Zoom} |");
            sb.AppendLine($"| FitToPagesWide | {t.FitToPagesWide} |");
            sb.AppendLine($"| FitToPagesTall | {t.FitToPagesTall} |");
            sb.AppendLine();
            sb.AppendLine("### DB Stored Ratios");
            sb.AppendLine($"| Ratio | Value | Points |");
            sb.AppendLine($"|-------|------:|-------:|");
            sb.AppendLine($"| Left | {t.DbLeft:F7} | {t.DbLeft * t.PageWidth:F2} |");
            sb.AppendLine($"| Top | {t.DbTop:F7} | {t.DbTop * t.PageHeight:F2} |");
            sb.AppendLine($"| Right | {t.DbRight:F7} | {t.DbRight * t.PageWidth:F2} |");
            sb.AppendLine($"| Bottom | {t.DbBottom:F7} | {t.DbBottom * t.PageHeight:F2} |");
            sb.AppendLine();
            sb.AppendLine("### Transformation Attempts");

            var transformations = new[]
            {
                ("Raw", t.RawLeftRatio, t.RawTopRatio, t.DiffRawLeft, t.DiffRawTop,
                    "COM_Left / PageWidth"),
                ("Printable", t.PrintableLeftRatio, t.PrintableTopRatio, t.DiffPrintableLeft, t.DiffPrintableTop,
                    "(COM_Left - LeftMargin) / (PageWidth - LeftMargin - RightMargin)"),
                ("Origin", t.OriginLeftRatio, t.OriginTopRatio, t.DiffOriginLeft, t.DiffOriginTop,
                    "(COM_Left - PrintedOrigin) / printableWidth"),
                ("MarginOnly", t.MarginOnlyLeftRatio, t.MarginOnlyTopRatio, t.DiffMarginLeft, t.DiffMarginTop,
                    "COM_Left / (PageWidth - RightMargin)"),
            };

            sb.AppendLine("| Transform | Formula | Left | DB Left | Diff Left | Top | DB Top | Diff Top |");
            sb.AppendLine("|-----------|---------|------|---------|----------|-----|--------|---------|");
            foreach (var (name, left, top, dLeft, dTop, formula) in transformations)
            {
                var marker = (dLeft + dTop) < 0.001 ? " ✓" : "";
                sb.AppendLine($"| {name}{marker} | {formula} | " +
                    $"{left:F7} | {t.DbLeft:F7} | {dLeft:F7} | " +
                    $"{top:F7} | {t.DbTop:F7} | {dTop:F7} |");
            }

            sb.AppendLine();
            sb.AppendLine($"**Best Transform:** {t.BestTransform} (total diff: {t.BestDiff:F7})");
            sb.AppendLine();
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "ClusterCoordinateTrace.md"), sb.ToString());
        _progress.Report($"  Wrote ClusterCoordinateTrace.md");
    }

    private async Task GenerateCoordinateSystemComparisonAsync(List<CoordinateTrace> traces, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Coordinate System Comparison");
        sb.AppendLine();
        sb.AppendLine("## Overview");
        sb.AppendLine();
        sb.AppendLine("This document compares three coordinate systems:");
        sb.AppendLine("1. **OpenXML Calculated** — Our current engine (estimates from column widths)");
        sb.AppendLine("2. **Excel COM** — Excel's own rendered positions (Range.Left/Top)");
        sb.AppendLine("3. **Database** — The legacy system's stored ratios");
        sb.AppendLine();
        sb.AppendLine("## Per-Cluster Comparison");
        sb.AppendLine();
        sb.AppendLine("| Address | DB Left | COM Left | DB Top | COM Top | DB→COM x | DB→COM y |");
        sb.AppendLine("|---------|---------|----------|--------|---------|----------|----------|");

        foreach (var t in traces.OrderBy(x => x.Address))
        {
            // DB points = ratio * page dimension
            double dbLeftPts = t.DbLeft * t.PageWidth;
            double dbTopPts = t.DbTop * t.PageHeight;

            sb.AppendLine($"| {t.Address} | " +
                $"{t.DbLeft:F7} ({dbLeftPts:F2}pt) | " +
                $"{t.ComLeft:F2}pt | " +
                $"{t.DbTop:F7} ({dbTopPts:F2}pt) | " +
                $"{t.ComTop:F2}pt | " +
                $"{dbLeftPts - t.ComLeft:+0.00;-0.00}pt | " +
                $"{dbTopPts - t.ComTop:+0.00;-0.00}pt |");
        }

        sb.AppendLine();
        sb.AppendLine("## Coordinate Math");

        // Analyze the relationship between COM and DB positions
        if (traces.Count > 1)
        {
            double avgRatioX = traces.Average(t =>
                t.ComLeft > 0 ? (t.DbLeft * t.PageWidth) / t.ComLeft : 0);
            double avgRatioY = traces.Average(t =>
                t.ComTop > 0 ? (t.DbTop * t.PageHeight) / t.ComTop : 0);

            double avgOffsetX = traces.Average(t =>
                (t.DbLeft * t.PageWidth) - t.ComLeft);
            double avgOffsetY = traces.Average(t =>
                (t.DbTop * t.PageHeight) - t.ComTop);

            sb.AppendLine();
            sb.AppendLine("## Summary Statistics");
            sb.AppendLine();
            sb.AppendLine($"- Average DB/COM ratio X: {avgRatioX:F6}");
            sb.AppendLine($"- Average DB/COM ratio Y: {avgRatioY:F6}");
            sb.AppendLine($"- Average offset X: {avgOffsetX:F2} pt");
            sb.AppendLine($"- Average offset Y: {avgOffsetY:F2} pt");
            sb.AppendLine();

            // Determine the relationship
            // If ratios cluster around 1.0: DB = COM / PageDimension (raw)
            // If ratios cluster around printableWidth/PageWidth: DB = (COM - margin) / printableWidth
            double printableW = traces[0].PageWidth - traces[0].LeftMargin - traces[0].RightMargin;
            double printableH = traces[0].PageHeight - traces[0].TopMargin - traces[0].BottomMargin;

            double printableRatioX = printableW / traces[0].PageWidth;
            double printableRatioY = printableH / traces[0].PageHeight;

            // Check if the printable area ratio matches avgRatioX
            if (Math.Abs(avgRatioX - printableRatioX) < 0.1)
            {
                sb.AppendLine($"- **Hypothesis: DB coordinates are relative to printable area (minus margins)**");
                sb.AppendLine($"  - Page width: {traces[0].PageWidth:F1} pt");
                sb.AppendLine($"  - Printable width: {printableW:F1} pt");
                sb.AppendLine($"  - Printable/Page ratio: {printableRatioX:F6}");
                sb.AppendLine($"  - DB Left = (COM_Left - LeftMargin) / PrintableWidth");
            }
            else
            {
                sb.AppendLine($"- Raw ratio: DB/COM = {avgRatioX:F6}");
                sb.AppendLine($"- Printable/Page ratio: {printableRatioX:F6}");
                sb.AppendLine($"- **No immediate match found** — investigating further...");
            }

            // Check for simple margin offset
            sb.AppendLine();
            sb.AppendLine("### Margin-Adjusted Hypothesis");
            foreach (var t in traces.Take(3))
            {
                double dbLeftFromMargin = (t.DbLeft * t.PageWidth) - t.LeftMargin;
                double dbTopFromMargin = (t.DbTop * t.PageHeight) - t.TopMargin;
                sb.AppendLine($"- {t.Address}: COM_Left={t.ComLeft:F2}, DB_left*Page-margin={dbLeftFromMargin:F2}, diff={t.ComLeft - dbLeftFromMargin:F2}");
            }
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "CoordinateSystemComparison.md"), sb.ToString());
        _progress.Report($"  Wrote CoordinateSystemComparison.md");
    }

    private async Task GenerateCoordinateMathAsync(List<CoordinateTrace> traces, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Coordinate Math");
        sb.AppendLine();
        sb.AppendLine("## Exact Transformation Formulas");
        sb.AppendLine();

        if (traces.Count == 0) { sb.AppendLine("No trace data available."); await File.WriteAllTextAsync(Path.Combine(dir, "CoordinateMath.md"), sb.ToString()); return; }

        // Find the best matching transformation across all clusters
        var t0 = traces.First();
        double printableW = t0.PageWidth - t0.LeftMargin - t0.RightMargin;
        double printableH = t0.PageHeight - t0.TopMargin - t0.BottomMargin;

        // Test the formula: DB = (COM - margin - alignmentOffset) / (page - margins)
        // For centered: alignmentOffset = (page - printable) / 2 = (L+R margins) / 2
        // For left-aligned: alignmentOffset = 0

        double centerOffsetX = (t0.LeftMargin + t0.RightMargin) / 2.0; // For horizontal centering
        double centerOffsetY = (t0.TopMargin + t0.BottomMargin) / 2.0; // For vertical centering

        sb.AppendLine("### Constants (from PageSetup)");
        sb.AppendLine();
        sb.AppendLine($"- PageWidth = {t0.PageWidth:F1} pt");
        sb.AppendLine($"- PageHeight = {t0.PageHeight:F2} pt");
        sb.AppendLine($"- LeftMargin = {t0.LeftMargin:F4} pt");
        sb.AppendLine($"- RightMargin = {t0.RightMargin:F4} pt");
        sb.AppendLine($"- TopMargin = {t0.TopMargin:F4} pt");
        sb.AppendLine($"- BottomMargin = {t0.BottomMargin:F4} pt");
        sb.AppendLine($"- PrintableWidth = PageWidth - LeftMargin - RightMargin = {printableW:F1} pt");
        sb.AppendLine($"- PrintableHeight = PageHeight - TopMargin - BottomMargin = {printableH:F1} pt");
        sb.AppendLine($"- CenterHorizontally = {t0.CenterHorizontally}");
        sb.AppendLine($"- CenterVertically = {t0.CenterVertically}");
        sb.AppendLine($"- FitToPagesWide = {t0.FitToPagesWide}");
        sb.AppendLine($"- FitToPagesTall = {t0.FitToPagesTall}");
        sb.AppendLine($"- Zoom = {t0.Zoom}");
        sb.AppendLine();

        // Test: printable area based formula
        sb.AppendLine("### Hypothesis 1: Printable Area Coordinates");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("DB_Left = (COM_Left - LeftMargin) / PrintableWidth");
        sb.AppendLine("DB_Top  = (COM_Top - TopMargin) / PrintableHeight");
        sb.AppendLine("```");
        sb.AppendLine();

        foreach (var t in traces)
        {
            double predLeft = (t.ComLeft - t.LeftMargin) / printableW;
            double predTop = (t.ComTop - t.TopMargin) / printableH;
            double predRight = predLeft + (t.ComWidth / printableW);
            double predBottom = predTop + (t.ComHeight / printableH);
            double diffL = Math.Abs(predLeft - t.DbLeft);
            double diffT = Math.Abs(predTop - t.DbTop);

            sb.AppendLine($"| {t.Address} | {predLeft:F7} | {t.DbLeft:F7} | {diffL:F7} | {predTop:F7} | {t.DbTop:F7} | {diffT:F7} |" +
                (diffL < 0.001 && diffT < 0.001 ? " ✓" : ""));
        }

        // Test: centered formula
        sb.AppendLine();
        sb.AppendLine("### Hypothesis 2: Centered Printable Coordinates");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("if CenterHorizontally:");
        sb.AppendLine("  PrintedOriginX = (PageWidth - PrintableWidth) / 2");
        sb.AppendLine("else:");
        sb.AppendLine("  PrintedOriginX = LeftMargin");
        sb.AppendLine("DB_Left = (COM_Left - PrintedOriginX) / PrintableWidth");
        sb.AppendLine("```");
        sb.AppendLine();

        foreach (var t in traces)
        {
            double originX = t.CenterHorizontally ? (t.PageWidth - printableW) / 2.0 : t.LeftMargin;
            double originY = t.CenterVertically ? (t.PageHeight - printableH) / 2.0 : t.TopMargin;
            double predLeft = (t.ComLeft - originX) / printableW;
            double predTop = (t.ComTop - originY) / printableH;
            double diffL = Math.Abs(predLeft - t.DbLeft);
            double diffT = Math.Abs(predTop - t.DbTop);

            sb.AppendLine($"| {t.Address} | {predLeft:F7} | {t.DbLeft:F7} | {diffL:F7} | {predTop:F7} | {t.DbTop:F7} | {diffT:F7} |" +
                (diffL < 0.001 && diffT < 0.001 ? " ✓" : ""));
        }

        // Test: Printed bounds - COM gives position from top-left of printable area
        // Range.Left includes margins! So:
        sb.AppendLine();
        sb.AppendLine("### Hypothesis 3: Page-Scaled Coordinates");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("DB_Left = COM_Left / PageWidth");
        sb.AppendLine("DB_Top  = COM_Top / PageHeight");
        sb.AppendLine("```");
        sb.AppendLine();

        foreach (var t in traces)
        {
            double diffL = Math.Abs(t.RawLeftRatio - t.DbLeft);
            double diffT = Math.Abs(t.RawTopRatio - t.DbTop);

            sb.AppendLine($"| {t.Address} | {t.RawLeftRatio:F7} | {t.DbLeft:F7} | {diffL:F7} | {t.RawTopRatio:F7} | {t.DbTop:F7} | {diffT:F7} |" +
                (diffL < 0.001 && diffT < 0.001 ? " ✓" : ""));
        }

        // Test: Zoom/Scale factored in
        sb.AppendLine();
        sb.AppendLine("### Hypothesis 4: Scaled Printable Coordinates");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("scale = Zoom / 100  (if Zoom is set)");
        sb.AppendLine("scale = 1 / FitToPagesWide  (if FitToPages is set)");
        sb.AppendLine("DB_Left = (COM_Left * scale - LeftMargin) / (PageWidth * scale - LeftMargin - RightMargin)");
        sb.AppendLine("```");
        sb.AppendLine();

        foreach (var t in traces)
        {
            double scale = t.Zoom.HasValue && t.Zoom > 0 ? t.Zoom.Value / 100.0 : 1.0;
            if (t.FitToPagesWide.HasValue && t.FitToPagesWide > 0)
                scale = 1.0 / t.FitToPagesWide.Value;

            double scaledPW = t.PageWidth * scale - t.LeftMargin - t.RightMargin;
            double scaledPH = t.PageHeight * scale - t.TopMargin - t.BottomMargin;

            if (scaledPW > 0 && scaledPH > 0)
            {
                double predLeft = (t.ComLeft * scale - t.LeftMargin) / scaledPW;
                double predTop = (t.ComTop * scale - t.TopMargin) / scaledPH;
                double diffL = Math.Abs(predLeft - t.DbLeft);
                double diffT = Math.Abs(predTop - t.DbTop);

                sb.AppendLine($"| {t.Address} | {predLeft:F7} | {t.DbLeft:F7} | {diffL:F7} | {predTop:F7} | {t.DbTop:F7} | {diffT:F7} |" +
                    (diffL < 0.001 && diffT < 0.001 ? " ✓" : ""));
            }
        }

        // Test: Pure COM position to page ratio (most likely for legacy system)
        sb.AppendLine();
        sb.AppendLine("### Hypothesis 5: Printed Page Coordinates (final rendered)");
        sb.AppendLine();
        sb.AppendLine("After all Excel layout is applied, Range.Left returns the position on the");
        sb.AppendLine("printed page (in points). The ratio is simply:");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("DB_Left = COM_Left / PageWidth");
        sb.AppendLine("DB_Top  = COM_Top / PageHeight");
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("This is the simplest hypothesis and matches Template 546.");
        sb.AppendLine();

        // Print summary
        sb.AppendLine();
        sb.AppendLine("## Best Hypothesis Determination");
        sb.AppendLine();

        // For each cluster, find the hypothesis with smallest total error
        var hypotheses = new[] { "Printable", "CenteredPrintable", "PageScaled", "ScaledPrintable" };
        foreach (var hyp in hypotheses)
        {
            double totalError = 0;
            foreach (var t in traces)
            {
                double error = hyp switch
                {
                    "Printable" => Math.Abs(t.PrintableLeftRatio - t.DbLeft) + Math.Abs(t.PrintableTopRatio - t.DbTop),
                    "CenteredPrintable" => Math.Abs(t.OriginLeftRatio - t.DbLeft) + Math.Abs(t.OriginTopRatio - t.DbTop),
                    "PageScaled" => Math.Abs(t.RawLeftRatio - t.DbLeft) + Math.Abs(t.RawTopRatio - t.DbTop),
                    _ => 999
                };
                totalError += error;
            }
            sb.AppendLine($"- {hyp}: total error = {totalError:F7}");
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "CoordinateMath.md"), sb.ToString());
        _progress.Report($"  Wrote CoordinateMath.md");
    }

    private async Task GenerateLegacyAlgorithmAsync(List<CoordinateTrace> traces, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Legacy Coordinate Algorithm");
        sb.AppendLine();
        sb.AppendLine("## Reverse-Engineered Formula");
        sb.AppendLine();

        if (traces.Count == 0) { sb.AppendLine("No trace data available."); return; }

        var t0 = traces.First();
        double printableW = t0.PageWidth - t0.LeftMargin - t0.RightMargin;
        double printableH = t0.PageHeight - t0.TopMargin - t0.BottomMargin;

        sb.AppendLine("### Step 1: Get workbook position from Excel COM");
        sb.AppendLine();
        sb.AppendLine("```vb");
        sb.AppendLine("Dim leftPts As Double");
        sb.AppendLine("Dim topPts As Double");
        sb.AppendLine("Dim widthPts As Double");
        sb.AppendLine("Dim heightPts As Double");
        sb.AppendLine("leftPts = Range(\"" + traces.First().Address + "\").Left");
        sb.AppendLine("topPts = Range(\"" + traces.First().Address + "\").Top");
        sb.AppendLine("widthPts = Range(\"" + traces.First().Address + "\").Width");
        sb.AppendLine("heightPts = Range(\"" + traces.First().Address + "\").Height");
        sb.AppendLine("```");
        sb.AppendLine();

        sb.AppendLine("### Step 2: Read PageSetup");
        sb.AppendLine();
        sb.AppendLine("```vb");
        sb.AppendLine($"Dim pageWidth As Double: pageWidth = {t0.PageWidth:F1}");
        sb.AppendLine($"Dim pageHeight As Double: pageHeight = {t0.PageHeight:F1}");
        sb.AppendLine($"Dim leftMargin As Double: leftMargin = {t0.LeftMargin:F4}");
        sb.AppendLine($"Dim rightMargin As Double: rightMargin = {t0.RightMargin:F4}");
        sb.AppendLine($"Dim topMargin As Double: topMargin = {t0.TopMargin:F4}");
        sb.AppendLine($"Dim bottomMargin As Double: bottomMargin = {t0.BottomMargin:F4}");
        sb.AppendLine($"Dim centerH As Boolean: centerH = {(t0.CenterHorizontally ? "True" : "False")}");
        sb.AppendLine($"Dim centerV As Boolean: centerV = {(t0.CenterVertically ? "True" : "False")}");
        sb.AppendLine($"Dim zoom As Integer: zoom = {t0.Zoom}");
        sb.AppendLine($"Dim fitW As Integer: fitW = {t0.FitToPagesWide}");
        sb.AppendLine($"Dim fitH As Integer: fitH = {t0.FitToPagesTall}");
        sb.AppendLine("```");
        sb.AppendLine();

        // Determine the confirmed formula based on best match
        double totalErrorPrintable = traces.Sum(t => Math.Abs(t.PrintableLeftRatio - t.DbLeft) + Math.Abs(t.PrintableTopRatio - t.DbTop));
        double totalErrorCentered = traces.Sum(t => Math.Abs(t.OriginLeftRatio - t.DbLeft) + Math.Abs(t.OriginTopRatio - t.DbTop));
        double totalErrorRaw = traces.Sum(t => Math.Abs(t.RawLeftRatio - t.DbLeft) + Math.Abs(t.RawTopRatio - t.DbTop));

        string bestFormula;
        if (totalErrorRaw < totalErrorPrintable && totalErrorRaw < totalErrorCentered)
            bestFormula = "Raw page ratio: DB = COM / PageDimension";
        else if (totalErrorPrintable < totalErrorCentered)
            bestFormula = "Printable area: DB = (COM - margin) / PrintableDimension";
        else
            bestFormula = "Centered printable: DB = (COM - printedOrigin) / PrintableDimension";

        sb.AppendLine("### Step 3: Compute printable area");
        sb.AppendLine();
        sb.AppendLine("```vb");
        sb.AppendLine("Dim printableWidth As Double");
        sb.AppendLine("printableWidth = pageWidth - leftMargin - rightMargin");
        sb.AppendLine($"' = {t0.PageWidth:F1} - {t0.LeftMargin:F4} - {t0.RightMargin:F4} = {printableW:F1}");
        sb.AppendLine("Dim printableHeight As Double");
        sb.AppendLine("printableHeight = pageHeight - topMargin - bottomMargin");
        sb.AppendLine($"' = {t0.PageHeight:F1} - {t0.TopMargin:F4} - {t0.BottomMargin:F4} = {printableH:F1}");
        sb.AppendLine("```");
        sb.AppendLine();

        sb.AppendLine($"### Best Fit: {bestFormula}");
        sb.AppendLine();

        // Show the formula with actual values
        foreach (var t in traces.Take(3))
        {
            double originX = t.CenterHorizontally ? (t.PageWidth - printableW) / 2.0 : t.LeftMargin;
            double originY = t.CenterVertically ? (t.PageHeight - printableH) / 2.0 : t.TopMargin;
            double predLeft = (t.ComLeft - originX) / printableW;
            double predTop = (t.ComTop - originY) / printableH;

            sb.AppendLine($"#### {t.Address}");
            sb.AppendLine();
            sb.AppendLine("```vb");
            sb.AppendLine($"leftPts = {t.ComLeft:F4}");
            sb.AppendLine($"topPts  = {t.ComTop:F4}");
            sb.AppendLine($"printedOriginX = {originX:F4}  ' {(t.CenterHorizontally ? "(pageWidth - printableWidth) / 2" : "leftMargin")}");
            sb.AppendLine($"printedOriginY = {originY:F4}  ' {(t.CenterVertically ? "(pageHeight - printableHeight) / 2" : "topMargin")}");
            sb.AppendLine($"leftRatio  = ({t.ComLeft:F4} - {originX:F4}) / {printableW:F1} = {predLeft:F7}  (DB: {t.DbLeft:F7}) {(Math.Abs(predLeft - t.DbLeft) < 0.001 ? "✓ MATCH" : "✗ MISMATCH")}");
            sb.AppendLine($"topRatio   = ({t.ComTop:F4} - {originY:F4}) / {printableH:F1} = {predTop:F7}  (DB: {t.DbTop:F7}) {(Math.Abs(predTop - t.DbTop) < 0.001 ? "✓ MATCH" : "✗ MISMATCH")}");
            sb.AppendLine("```");
            sb.AppendLine();
        }

        sb.AppendLine("## Final Algorithm");
        sb.AppendLine();
        sb.AppendLine("```vb");
        sb.AppendLine("' For each cluster cell address:");
        sb.AppendLine("Function ComputeClusterRatio(cellAddress As String) As (left, top, right, bottom)");
        sb.AppendLine("    Dim rng As Range");
        sb.AppendLine("    Set rng = ws.Range(cellAddress)");
        sb.AppendLine("    ");
        sb.AppendLine("    ' 1. Get COM positions");
        sb.AppendLine("    Dim leftPts As Double: leftPts = rng.Left");
        sb.AppendLine("    Dim topPts As Double: topPts = rng.Top");
        sb.AppendLine("    Dim widthPts As Double: widthPts = rng.Width");
        sb.AppendLine("    Dim heightPts As Double: heightPts = rng.Height");
        sb.AppendLine("    ");
        sb.AppendLine("    ' 2. Get page dimensions");
        sb.AppendLine("    Dim pw As Double: pw = ws.PageSetup.PageWidth");
        sb.AppendLine("    Dim ph As Double: ph = ws.PageSetup.PageHeight");
        sb.AppendLine("    Dim lm As Double: lm = ws.PageSetup.LeftMargin");
        sb.AppendLine("    Dim rm As Double: rm = ws.PageSetup.RightMargin");
        sb.AppendLine("    Dim tm As Double: tm = ws.PageSetup.TopMargin");
        sb.AppendLine("    Dim bm As Double: bm = ws.PageSetup.BottomMargin");
        sb.AppendLine("    ");
        sb.AppendLine("    ' 3. Compute printable area");
        sb.AppendLine("    Dim printW As Double: printW = pw - lm - rm");
        sb.AppendLine("    Dim printH As Double: printH = ph - tm - bm");
        sb.AppendLine("    ");
        sb.AppendLine("    ' 4. Compute printed origin (accounts for centering)");
        sb.AppendLine("    Dim originX As Double");
        sb.AppendLine("    If ws.PageSetup.CenterHorizontally Then");
        sb.AppendLine("        originX = (pw - printW) / 2");
        sb.AppendLine("    Else");
        sb.AppendLine("        originX = lm");
        sb.AppendLine("    End If");
        sb.AppendLine("    ");
        sb.AppendLine("    Dim originY As Double");
        sb.AppendLine("    If ws.PageSetup.CenterVertically Then");
        sb.AppendLine("        originY = (ph - printH) / 2");
        sb.AppendLine("    Else");
        sb.AppendLine("        originY = tm");
        sb.AppendLine("    End If");
        sb.AppendLine("    ");
        sb.AppendLine("    ' 5. Convert to ratios");
        sb.AppendLine("    left   = RoundEx((leftPts - originX) / printW, 7)");
        sb.AppendLine("    top    = RoundEx((topPts - originY) / printH, 7)");
        sb.AppendLine("    right  = RoundEx((leftPts + widthPts - originX) / printW, 7)");
        sb.AppendLine("    bottom = RoundEx((topPts + heightPts - originY) / printH, 7)");
        sb.AppendLine("    ");
        sb.AppendLine("    Return (left, top, right, bottom)");
        sb.AppendLine("End Function");
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("Note: `RoundEx` uses VB6's round-half-away-from-zero with float (x87 FPU 80-bit precision).");
        sb.AppendLine("In .NET, this is `Math.Round((float)value, 7, MidpointRounding.AwayFromZero)`.");

        await File.WriteAllTextAsync(Path.Combine(dir, "LegacyCoordinateAlgorithm.md"), sb.ToString());
        _progress.Report($"  Wrote LegacyCoordinateAlgorithm.md");
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static double SafeRatio(double numerator, double denominator)
    {
        if (Math.Abs(denominator) < 0.001) return 0;
        return Math.Round(numerator / denominator, 7);
    }

    private class PageSetupInfo
    {
        public int? PaperSize { get; set; }
        public int? Zoom { get; set; }
        public int? FitToPagesWide { get; set; }
        public int? FitToPagesTall { get; set; }
        public double LeftMargin { get; set; }
        public double RightMargin { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double HeaderMargin { get; set; }
        public double FooterMargin { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public int? Orientation { get; set; }
        public double PageWidth { get; set; }
        public double PageHeight { get; set; }
        public string? PrintArea { get; set; }
        public int? Order { get; set; }
        public int? FirstPageNumber { get; set; }
        public int[]? PrintQuality { get; set; }
        public int UsedRangeRows { get; set; }
        public int UsedRangeColumns { get; set; }
        public int HPageBreaks { get; set; }
        public int VPageBreaks { get; set; }
    }

    public class CoordinateTrace
    {
        public string Address { get; set; } = "";
        // COM positions (points)
        public double ComLeft { get; set; }
        public double ComTop { get; set; }
        public double ComWidth { get; set; }
        public double ComHeight { get; set; }
        // DB ratios
        public double DbLeft { get; set; }
        public double DbTop { get; set; }
        public double DbRight { get; set; }
        public double DbBottom { get; set; }
        // Page setup
        public double PageWidth { get; set; }
        public double PageHeight { get; set; }
        public double LeftMargin { get; set; }
        public double RightMargin { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public int? Zoom { get; set; }
        public int? FitToPagesWide { get; set; }
        public int? FitToPagesTall { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        // Computed transformations
        public double RawLeftRatio { get; set; }
        public double RawTopRatio { get; set; }
        public double RawRightRatio { get; set; }
        public double RawBottomRatio { get; set; }
        public double RawWidthRatio { get; set; }
        public double RawHeightRatio { get; set; }
        public double PrintableLeftRatio { get; set; }
        public double PrintableTopRatio { get; set; }
        public double PrintableRightRatio { get; set; }
        public double PrintableBottomRatio { get; set; }
        public double OriginLeftRatio { get; set; }
        public double OriginTopRatio { get; set; }
        public double MarginOnlyLeftRatio { get; set; }
        public double MarginOnlyTopRatio { get; set; }
        public double ScaledLeftRatio { get; set; }
        public double ScaledTopRatio { get; set; }
        public double DiffScaledLeft { get; set; }
        public double DiffScaledTop { get; set; }
        public double ScaleX { get; set; } = 1.0;
        public double ScaleY { get; set; } = 1.0;
        public double PrintedOriginX { get; set; }
        public double PrintedOriginY { get; set; }
        // Differences from DB
        public double DiffRawLeft { get; set; }
        public double DiffRawTop { get; set; }
        public double DiffPrintableLeft { get; set; }
        public double DiffPrintableTop { get; set; }
        public double DiffOriginLeft { get; set; }
        public double DiffOriginTop { get; set; }
        public double DiffMarginLeft { get; set; }
        public double DiffMarginTop { get; set; }
        public string BestTransform { get; set; } = "";
        public double BestDiff { get; set; }
    }
}
