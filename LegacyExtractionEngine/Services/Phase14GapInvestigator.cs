using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 14 — Identify the Root Cause of Template-Specific Gap Parameters.
/// 
/// Measures every workbook property and correlates against BaseGap, PerRowExtra, ColumnGap.
/// No curve fitting. No magic constants. Pure correlation analysis.
/// </summary>
public class Phase14GapInvestigator
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;

    private static readonly Dictionary<int, string> WorkbookPaths = new()
    {
        { 546, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx" },
        { 547, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx" },
        { 548, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx" }
    };

    public Phase14GapInvestigator(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test", "Phase14_GapOrigin");
        if (Directory.Exists(_outputDir)) Directory.Delete(_outputDir, true);
        Directory.CreateDirectory(_outputDir);
    }

    public async Task RunAsync()
    {
        _progress.Report("============================================");
        _progress.Report(" Phase 14 — Gap Origin Investigation");
        _progress.Report(" Templates: 546, 547, 548");
        _progress.Report("============================================");

        var report = new StringBuilder();
        report.AppendLine("# Phase 14 — Legacy Gap Origin Investigation Report");
        report.AppendLine();
        report.AppendLine($"**Generated:** {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
        report.AppendLine();
        report.AppendLine("## Objective");
        report.AppendLine();
        report.AppendLine("Identify the measurable Excel property that explains why the vertical gap");
        report.AppendLine("parameters differ between templates 546 (BaseGap=0.87pt), 547 (BaseGap=4.44pt),");
        report.AppendLine("and 548 (BaseGap=0pt). No curve fitting — every parameter must be explained");
        report.AppendLine("by a measurable workbook property.");
        report.AppendLine();

        var templateData = new List<TemplateMetrics>();

        foreach (var id in new[] { 546, 547, 548 })
        {
            _progress.Report($"=== Processing template {id} ===");
            var tm = await MeasureTemplateAsync(id);
            templateData.Add(tm);
            report.Append(tm.Section);
            report.AppendLine("---");
            report.AppendLine();
        }

        // Cross-template correlation
        report.AppendLine("# Cross-Template Correlation Analysis");
        report.AppendLine();
        report.AppendLine("| Property | 546 | 547 | 548 | Correlates with BaseGap? |");
        report.AppendLine("|----------|:---:|:---:|:---:|:------------------------:|");

        // Build property table
        var properties = new Dictionary<string, (double V546, double V547, double V548)>();

        foreach (var tm in templateData)
        {
            foreach (var kv in tm.Properties)
            {
                if (!properties.ContainsKey(kv.Key))
                    properties[kv.Key] = (0, 0, 0);
                var p = properties[kv.Key];
                if (tm.DefTopId == 546) p.V546 = kv.Value;
                else if (tm.DefTopId == 547) p.V547 = kv.Value;
                else p.V548 = kv.Value;
                properties[kv.Key] = p;
            }
        }

        double[] baseGaps = { templateData[0].BaseGap, templateData[1].BaseGap, templateData[2].BaseGap };
        double[] perRowExtras = { templateData[0].PerRowExtra, templateData[1].PerRowExtra, templateData[2].PerRowExtra };
        double[] colGaps = { templateData[0].ColGap, templateData[1].ColGap, templateData[2].ColGap };

        foreach (var kv in properties.OrderBy(p => p.Key))
        {
            double[] vals = { kv.Value.V546, kv.Value.V547, kv.Value.V548 };
            double rBg = PearsonCorrelation(vals, baseGaps);
            double rPr = PearsonCorrelation(vals, perRowExtras);
            double rCg = PearsonCorrelation(vals, colGaps);

            string correlates = "NONE";
            double maxR = Math.Max(Math.Abs(rBg), Math.Max(Math.Abs(rPr), Math.Abs(rCg)));
            if (maxR > 0.9) correlates = $"BaseGap(r={rBg:F2}) PerRowExtra(r={rPr:F2}) ColGap(r={rCg:F2})";
            else if (maxR > 0.7) correlates = $"Weak: BaseGap(r={rBg:F2}) PerRowExtra(r={rPr:F2}) ColGap(r={rCg:F2})";

            report.AppendLine($"| {kv.Key} | {kv.Value.V546:G4} | {kv.Value.V547:G4} | {kv.Value.V548:G4} | {correlates} |");
        }
        report.AppendLine();

        // Conclusion
        report.AppendLine("## Conclusions");
        report.AppendLine();

        // Find the best correlated property
        double bestCorrelation = 0;
        string bestProperty = "";
        foreach (var kv in properties.OrderBy(p => p.Key))
        {
            double[] vals = { kv.Value.V546, kv.Value.V547, kv.Value.V548 };
            double r = PearsonCorrelation(vals, baseGaps);
            if (Math.Abs(r) > Math.Abs(bestCorrelation))
            {
                bestCorrelation = r;
                bestProperty = kv.Key;
            }
        }

        report.AppendLine($"The property with the strongest correlation to BaseGap is **{bestProperty}** (r={bestCorrelation:F4}).");
        report.AppendLine();
        report.AppendLine("However, with only 3 data points, statistical significance is limited.");
        report.AppendLine("Any correlation above r=0.997 would be needed for p<0.05 with n=3.");
        report.AppendLine();

        if (Math.Abs(bestCorrelation) >= 0.99)
        {
            report.AppendLine($"**STRONG CANDIDATE:** {bestProperty} correlates with BaseGap at r={bestCorrelation:F4}.");
            report.AppendLine("This suggests a measurable explanation exists.");
        }
        else
        {
            report.AppendLine("**No single property strongly correlates with BaseGap across all 3 templates.**");
            report.AppendLine("The gaps likely originate from a multi-variable interaction or from processing");
            report.AppendLine("performed outside the workbook (legacy preprocessing, rendering pipeline, etc.).");
        }
        report.AppendLine();

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "Phase14GapOriginReport.md"), report.ToString());
        _progress.Report($"Report: {_outputDir}\\Phase14GapOriginReport.md");
    }

    private async Task<TemplateMetrics> MeasureTemplateAsync(int defTopId)
    {
        var tm = new TemplateMetrics { DefTopId = defTopId };

        // Load DB data for clusters and gap calculation
        var dbReader = new DatabaseReader(_connectionString, _progress);
        var db = await dbReader.ReadAllAsync(defTopId);
        var dbClusters = db?.DefSheets.SelectMany(s => s.Clusters).ToList() ?? new List<DefCluster>();

        // Read OpenXML first (for PrintArea, column widths, row heights)
        var workbookPath = WorkbookPaths.GetValueOrDefault(defTopId);
        var openXmlInfo = ReadOpenXmlInfo(workbookPath);

        // Parse PrintArea bounds
        int printAreaStartRow = 1, printAreaEndRow = 0;
        int printAreaStartCol = 1, printAreaEndCol = 0;
        if (!string.IsNullOrEmpty(openXmlInfo.PrintArea))
        {
            var pa = openXmlInfo.PrintArea.Replace("$", "").Split(':');
            if (pa.Length == 2)
            {
                var start = ParseCellRef(pa[0]);
                var end = ParseCellRef(pa[1]);
                printAreaStartRow = start.Row;
                printAreaEndRow = end.Row;
                printAreaStartCol = start.Col;
                printAreaEndCol = end.Col;
            }
        }

        tm.Properties["PrintAreaRows"] = printAreaEndRow - printAreaStartRow + 1;
        tm.Properties["PrintAreaCols"] = printAreaEndCol - printAreaStartCol + 1;

        // Open COM for detailed measurements
        var excelType = Type.GetTypeFromProgID("Excel.Application");
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

                // Print settings
                tm.Properties["Zoom"] = SafeGetInt(() => (int)ps.Zoom);
                tm.Properties["FitToPagesWide"] = SafeGetInt(() => (int)ps.FitToPagesWide);
                tm.Properties["FitToPagesTall"] = SafeGetInt(() => (int)ps.FitToPagesTall);
                tm.Properties["PaperSize"] = SafeGetInt(() => (int)ps.PaperSize);
                tm.Properties["Orientation"] = SafeGetInt(() => (int)ps.Orientation);

                // Page dimensions
                double pw = SafeGetDouble(() => (double)ps.PageWidth, 612);
                double ph = SafeGetDouble(() => (double)ps.PageHeight, 792);
                double lm = SafeGetDouble(() => (double)ps.LeftMargin, 0);
                double rm = SafeGetDouble(() => (double)ps.RightMargin, 0);
                double tm_m = SafeGetDouble(() => (double)ps.TopMargin, 0);
                double bm = SafeGetDouble(() => (double)ps.BottomMargin, 0);
                tm.Properties["PageWidth"] = pw;
                tm.Properties["PageHeight"] = ph;
                tm.Properties["LeftMargin"] = lm;
                tm.Properties["RightMargin"] = rm;
                tm.Properties["TopMargin"] = tm_m;
                tm.Properties["BottomMargin"] = bm;
                tm.Properties["PrintableWidth"] = pw - lm - rm;
                tm.Properties["PrintableHeight"] = ph - tm_m - bm;

                // Default row height and column width
                double stdHt = SafeGetDouble(() => (double)ws.StandardHeight, 14.4);
                double stdWd = SafeGetDouble(() => (double)ws.StandardWidth, 8.43);
                tm.Properties["StandardRowHeight"] = stdHt;
                tm.Properties["StandardColumnWidth"] = stdWd;

                // Read ALL row heights within print area
                double totalContentHeight = 0;
                double minRowHt = double.MaxValue;
                double maxRowHt = 0;
                double sumRowHt = 0;
                int rowCount = 0;

                for (int r = printAreaStartRow; r <= Math.Max(printAreaEndRow, 50); r++)
                {
                    try
                    {
                        dynamic rowRange = ws.Rows[r];
                        double ht = (double)rowRange.Height;
                        totalContentHeight += ht;
                        minRowHt = Math.Min(minRowHt, ht);
                        maxRowHt = Math.Max(maxRowHt, ht);
                        sumRowHt += ht;
                        rowCount++;
                    }
                    catch { break; }
                }

                tm.Properties["TotalContentHeight"] = totalContentHeight;
                tm.Properties["MinRowHt"] = minRowHt;
                tm.Properties["MaxRowHt"] = maxRowHt;
                tm.Properties["AvgRowHt"] = rowCount > 0 ? sumRowHt / rowCount : 0;

                // Read ALL column widths within print area
                double totalContentWidth = 0;
                double minColWd = double.MaxValue;
                double maxColWd = 0;
                double sumColWd = 0;
                int colCount = 0;

                for (int c = printAreaStartCol; c <= Math.Max(printAreaEndCol, 20); c++)
                {
                    try
                    {
                        dynamic colRange = ws.Columns[c];
                        double chars = (double)colRange.ColumnWidth;
                        double pts = chars * 50.04 / 8.43;
                        totalContentWidth += pts;
                        minColWd = Math.Min(minColWd, pts);
                        maxColWd = Math.Max(maxColWd, pts);
                        sumColWd += pts;
                        colCount++;
                    }
                    catch { break; }
                }

                tm.Properties["TotalContentWidth"] = totalContentWidth;
                tm.Properties["MinColWd"] = minColWd;
                tm.Properties["MaxColWd"] = maxColWd;
                tm.Properties["AvgColWd"] = colCount > 0 ? sumColWd / colCount : 0;

                // Content-to-page ratios
                tm.Properties["ContentHtRatio"] = ph > 0 ? totalContentHeight / ph : 0;
                tm.Properties["ContentWdRatio"] = pw > 0 ? totalContentWidth / pw : 0;

                // Compute PrintArea dimensions from OpenXML (more reliable than COM)
                double openXmlContentWidth = 0;
                for (int c = printAreaStartCol; c <= printAreaEndCol; c++)
                {
                    if (openXmlInfo.ColumnWidths.TryGetValue(c, out var cxw))
                        openXmlContentWidth += cxw * 50.04 / 8.43;
                    else
                        openXmlContentWidth += stdWd * 50.04 / 8.43;
                }
                double openXmlContentHeight = 0;
                for (int r = printAreaStartRow; r <= printAreaEndRow; r++)
                {
                    if (openXmlInfo.RowHeights.TryGetValue(r, out var rxh))
                        openXmlContentHeight += rxh.Height;
                    else
                        openXmlContentHeight += stdHt;
                }
                tm.Properties["OpenXmlContentWidth"] = openXmlContentWidth;
                tm.Properties["OpenXmlContentHeight"] = openXmlContentHeight;
                tm.Properties["OpenXmlWdRatio"] = pw > 0 ? openXmlContentWidth / pw : 0;
                tm.Properties["OpenXmlHtRatio"] = ph > 0 ? openXmlContentHeight / ph : 0;

                // Horizontal scale factor (width constraint)
                double hScale = openXmlContentWidth > pw ? pw / openXmlContentWidth : 1.0;
                double vScale = openXmlContentHeight > ph ? ph / openXmlContentHeight : 1.0;
                tm.Properties["HScale"] = hScale;
                tm.Properties["VScale"] = vScale;
                tm.Properties["UniformScale"] = Math.Min(hScale, vScale);

                // Number of merged cells
                tm.Properties["MergedCells"] = openXmlInfo.MergedCells.Count;

                // Compute BaseGap from clusters
                string? firstAddr = null;
                double firstComL = 0, firstComT = 0, firstDbL = 0, firstDbT = 0;
                int originRow = 0;

                foreach (var dbc in dbClusters)
                {
                    var addr = dbc.CellAddr?.Trim() ?? "";
                    if (string.IsNullOrEmpty(addr)) continue;
                    try
                    {
                        dynamic range = ws.Range[addr];
                        int rowNum = (int)range.Row;
                        double comL = (double)range.Left;
                        double comT = (double)range.Top;

                        if (firstAddr == null)
                        {
                            firstAddr = addr;
                            firstComL = comL;
                            firstComT = comT;
                            firstDbL = ParseDouble(dbc.LeftPosition);
                            firstDbT = ParseDouble(dbc.TopPosition);
                            originRow = rowNum;
                        }
                    }
                    catch { }
                }

                if (firstAddr != null)
                {
                    double ox = ComCoordinateService.RoundEx(firstDbL * pw - firstComL);
                    double oy = ComCoordinateService.RoundEx(firstDbT * ph - firstComT);

                    // Compute gaps for subsequent clusters
                    double firstGapT = 0;
                    int firstOffset = 0;
                    double secondGapT = 0;
                    int secondOffset = 0;
                    int gapCount = 0;

                    foreach (var dbc in dbClusters)
                    {
                        var addr = dbc.CellAddr?.Trim() ?? "";
                        if (string.IsNullOrEmpty(addr) || addr == firstAddr) continue;
                        try
                        {
                            dynamic range = ws.Range[addr];
                            int rowNum = (int)range.Row;
                            double comT = (double)range.Top;
                            double dbT = ParseDouble(dbc.TopPosition);
                            double neededGap = dbT * ph - (comT + oy);
                            int offset = rowNum - originRow;

                            if (offset > 0)
                            {
                                gapCount++;
                                if (gapCount == 1)
                                {
                                    firstGapT = neededGap;
                                    firstOffset = offset;
                                }
                                else if (gapCount == 2)
                                {
                                    secondGapT = neededGap;
                                    secondOffset = offset;
                                }
                            }
                        }
                        catch { }
                    }

                    double baseGap = firstOffset > 0 ? firstGapT / firstOffset : 0;
                    double perRowExtra = (firstOffset > 0 && secondOffset > 0)
                        ? (secondGapT / secondOffset - firstGapT / firstOffset) / (secondOffset - firstOffset)
                        : 0;

                    tm.BaseGap = baseGap;
                    tm.PerRowExtra = perRowExtra;

                    // Column gap
                    double firstGapL = 0;
                    int firstColOffset = 0;
                    gapCount = 0;
                    foreach (var dbc in dbClusters)
                    {
                        var addr = dbc.CellAddr?.Trim() ?? "";
                        if (string.IsNullOrEmpty(addr) || addr == firstAddr) continue;
                        try
                        {
                            dynamic range = ws.Range[addr];
                            int colNum = (int)range.Column;
                            double comL = (double)range.Left;
                            double dbL = ParseDouble(dbc.LeftPosition);
                            double neededGap = dbL * pw - (comL + ox);
                            int colOffset = colNum - 1;

                            if (Math.Abs(neededGap) > 0.01 && colOffset > 0)
                            {
                                gapCount++;
                                if (gapCount == 1)
                                {
                                    firstGapL = neededGap;
                                    firstColOffset = colOffset;
                                }
                            }
                        }
                        catch { }
                    }
                    tm.ColGap = firstColOffset > 0 ? firstGapL / firstColOffset : 0;
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

        // Build section
        var sb = new StringBuilder();
        sb.AppendLine($"## Template {defTopId}");
        sb.AppendLine();
        sb.AppendLine($"**PrintArea:** {openXmlInfo.PrintArea ?? "(none)"}");
        sb.AppendLine($"**Rows:** {printAreaStartRow}–{printAreaEndRow} ({printAreaEndRow - printAreaStartRow + 1} rows)");
        sb.AppendLine($"**Cols:** {printAreaStartCol}–{printAreaEndCol} ({printAreaEndCol - printAreaStartCol + 1} cols)");
        sb.AppendLine($"**BaseGap:** {tm.BaseGap:F4}pt");
        sb.AppendLine($"**PerRowExtra:** {tm.PerRowExtra:F4}pt");
        sb.AppendLine($"**ColGap:** {tm.ColGap:F4}pt");
        sb.AppendLine();

        sb.AppendLine("### Print Settings");
        sb.AppendLine();
        sb.AppendLine("| Property | Value |");
        sb.AppendLine("|----------|-------|");
        sb.AppendLine($"| Zoom | {tm.Properties.GetValueOrDefault("Zoom")}% |");
        sb.AppendLine($"| FitToPagesWide | {tm.Properties.GetValueOrDefault("FitToPagesWide")} |");
        sb.AppendLine($"| FitToPagesTall | {tm.Properties.GetValueOrDefault("FitToPagesTall")} |");
        sb.AppendLine($"| Orientation | {tm.Properties.GetValueOrDefault("Orientation")} |");
        sb.AppendLine($"| Page Size | {tm.Properties.GetValueOrDefault("PageWidth")}x{tm.Properties.GetValueOrDefault("PageHeight")} |");
        sb.AppendLine();

        sb.AppendLine("### Content Dimensions");
        sb.AppendLine();
        sb.AppendLine("| Dimension | Value |");
        sb.AppendLine("|-----------|-------|");
        sb.AppendLine($"| TotalContentHeight (COM) | {tm.Properties.GetValueOrDefault("TotalContentHeight"):F2}pt |");
        sb.AppendLine($"| TotalContentWidth (COM) | {tm.Properties.GetValueOrDefault("TotalContentWidth"):F2}pt |");
        sb.AppendLine($"| OpenXmlContentHeight | {tm.Properties.GetValueOrDefault("OpenXmlContentHeight"):F2}pt |");
        sb.AppendLine($"| OpenXmlContentWidth | {tm.Properties.GetValueOrDefault("OpenXmlContentWidth"):F2}pt |");
        sb.AppendLine($"| ContentHtRatio (COM/Page) | {tm.Properties.GetValueOrDefault("ContentHtRatio"):F4} |");
        sb.AppendLine($"| ContentWdRatio (COM/Page) | {tm.Properties.GetValueOrDefault("ContentWdRatio"):F4} |");
        sb.AppendLine($"| OpenXmlHtRatio | {tm.Properties.GetValueOrDefault("OpenXmlHtRatio"):F4} |");
        sb.AppendLine($"| OpenXmlWdRatio | {tm.Properties.GetValueOrDefault("OpenXmlWdRatio"):F4} |");
        sb.AppendLine($"| HScale (width constraint) | {tm.Properties.GetValueOrDefault("HScale"):F4} |");
        sb.AppendLine($"| VScale (height constraint) | {tm.Properties.GetValueOrDefault("VScale"):F4} |");
        sb.AppendLine($"| UniformScale | {tm.Properties.GetValueOrDefault("UniformScale"):F4} |");
        sb.AppendLine($"| MergedCells | {tm.Properties.GetValueOrDefault("MergedCells")} |");
        sb.AppendLine($"| StandardRowHeight | {tm.Properties.GetValueOrDefault("StandardRowHeight"):F2}pt |");
        sb.AppendLine($"| MinRowHt | {tm.Properties.GetValueOrDefault("MinRowHt"):F2}pt |");
        sb.AppendLine($"| MaxRowHt | {tm.Properties.GetValueOrDefault("MaxRowHt"):F2}pt |");
        sb.AppendLine($"| AvgRowHt | {tm.Properties.GetValueOrDefault("AvgRowHt"):F2}pt |");
        sb.AppendLine($"| AvgColWd | {tm.Properties.GetValueOrDefault("AvgColWd"):F2}pt |");
        sb.AppendLine();

        tm.Section = sb.ToString();
        return tm;
    }

    private OpenXmlPrintArea ReadOpenXmlInfo(string workbookPath)
    {
        var info = new OpenXmlPrintArea();
        if (!File.Exists(workbookPath)) return info;

        try
        {
            using var doc = SpreadsheetDocument.Open(workbookPath, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return info;

            var wbXml = XDocument.Load(wbPart.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            // PrintArea from defined names
            var printAreaDef = wbXml.Descendants(ns + "definedName")
                .FirstOrDefault(d =>
                {
                    var name = (string?)d.Attribute("name");
                    return name != null && (name == "_xlnm.Print_Area" || name.EndsWith("Print_Area", StringComparison.OrdinalIgnoreCase));
                });
            if (printAreaDef != null)
            {
                var val = printAreaDef.Value?.Trim();
                if (!string.IsNullOrEmpty(val))
                {
                    var bang = val.IndexOf('!');
                    if (bang >= 0) val = val.Substring(bang + 1);
                    info.PrintArea = val.Replace("'", "");
                }
            }

            var wsPart = wbPart.WorksheetParts.FirstOrDefault();
            if (wsPart == null) return info;

            var wsXml = XDocument.Load(wsPart.GetStream());

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
        catch { }
        return info;
    }

    private static (int Row, int Col) ParseCellRef(string refStr)
    {
        var letters = new string(refStr.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(refStr.SkipWhile(char.IsLetter).ToArray());
        int row = int.TryParse(digits, out var r) ? r : 0;
        int col = 0;
        foreach (var ch in letters)
            col = col * 26 + (char.ToUpper(ch) - 'A' + 1);
        return (row, col);
    }

    private static double PearsonCorrelation(double[] x, double[] y)
    {
        if (x.Length != y.Length || x.Length < 2) return 0;
        double mx = x.Average(), my = y.Average();
        double num = 0, denX = 0, denY = 0;
        for (int i = 0; i < x.Length; i++)
        {
            double dx = x[i] - mx, dy = y[i] - my;
            num += dx * dy;
            denX += dx * dx;
            denY += dy * dy;
        }
        if (denX == 0 || denY == 0) return 0;
        return num / Math.Sqrt(denX * denY);
    }

    private static double ParseDouble(string? s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;

    private static double SafeGetDouble(Func<double> getter, double fallback = 0)
    {
        try { return getter(); } catch { return fallback; }
    }

    private static int SafeGetInt(Func<int> getter, int fallback = 0)
    {
        try { return getter(); } catch { return fallback; }
    }
}

public class TemplateMetrics
{
    public int DefTopId { get; set; }
    public double BaseGap { get; set; }
    public double PerRowExtra { get; set; }
    public double ColGap { get; set; }
    public Dictionary<string, double> Properties { get; set; } = new();
    public string Section { get; set; } = "";
}

public class OpenXmlPrintArea
{
    public string? PrintArea { get; set; }
    public Dictionary<int, double> ColumnWidths { get; set; } = new();
    public Dictionary<int, (double Height, bool IsCustom)> RowHeights { get; set; } = new();
    public List<string> MergedCells { get; set; } = new();
}
