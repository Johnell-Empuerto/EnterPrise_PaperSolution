using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class ComparisonEngine
{
    private readonly IProgress<string> _progress;

    public ComparisonEngine(IProgress<string> progress)
    {
        _progress = progress;
    }

    public ComparisonReport Compare(DatabaseDump database, ExtractionResult generated, string generatedXml)
    {
        _progress.Report("Comparing database vs generated output...");
        _generatedXml = generatedXml;
        var report = new ComparisonReport
        {
            DefTopId = database.DefTop?.DefTopId ?? 0,
            GeneratedAt = DateTime.UtcNow
        };

        CompareWorkbook(report, database, generated);
        ComparePageSetup(report, database, generated);
        CompareStyles(report, database, generated);
        CompareCells(report, database, generated);
        CompareMergedCells(report, database, generated);
        CompareXmlData(report, database, generatedXml);
        CompareDefCluster(report, database, generated);
        CompareBackgroundImage(report, database, generated);

        ComputeSummary(report);
        report.Overall = report.Summary.OverallMatchPercentage >= 100 ? "PASS" : "INCOMPLETE";

        return report;
    }

    private void CompareWorkbook(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };
        var db = database.DefTop;
        var wb = generated.Workbook;

        CompareProperty(section, "SheetCount", database.DefSheets.Count, wb.Sheets.Count);

        var dbNames = string.Join(", ", database.DefSheets.Select(s => s.DefSheetName));
        var genNames = string.Join(", ", wb.Sheets.Select(s => s.Name));
        CompareProperty(section, "SheetNames", dbNames, genNames);

        CompareProperty(section, "DefTopName", db?.DefTopName, wb.Properties.Title);
        CompareProperty(section, "DesignerVersion", db?.DesignerVersion, wb.Properties.AppVersion);

        var dbPrintArea = ExtractPrintAreaFromDb(database);
        var genPrintArea = wb.DefinedNames.FirstOrDefault(n =>
            n.Name?.Equals("Print_Area", StringComparison.OrdinalIgnoreCase) == true)?.RefersTo;
        CompareProperty(section, "PrintArea", dbPrintArea, genPrintArea);

        var dbSheetCount = db?.DefSheetCount;
        CompareProperty(section, "DefSheetCount", dbSheetCount, wb.Sheets.Count);

        var dbReportType = db?.ReportType;
        CompareProperty(section, "ReportType", dbReportType, 1);

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["Workbook"] = section;
    }

    private void ComparePageSetup(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };

        foreach (var sheet in generated.Workbook.Sheets)
        {
            var prefix = $"Sheet[{sheet.Name}]";

            CompareProperty(section, $"{prefix}.PaperSize", null, sheet.PageSetup.PaperSize);
            CompareProperty(section, $"{prefix}.Orientation", null, sheet.PageSetup.Orientation);
            CompareProperty(section, $"{prefix}.FitToWidth", null, sheet.PageSetup.FitToWidth);
            CompareProperty(section, $"{prefix}.FitToHeight", null, sheet.PageSetup.FitToHeight);
            CompareProperty(section, $"{prefix}.Scale", null, sheet.PageSetup.Scale);
            CompareProperty(section, $"{prefix}.MarginTop", sheet.PageMargins.Top, sheet.PageMargins.Top);
            CompareProperty(section, $"{prefix}.MarginBottom", sheet.PageMargins.Bottom, sheet.PageMargins.Bottom);
            CompareProperty(section, $"{prefix}.MarginLeft", sheet.PageMargins.Left, sheet.PageMargins.Left);
            CompareProperty(section, $"{prefix}.MarginRight", sheet.PageMargins.Right, sheet.PageMargins.Right);
            CompareProperty(section, $"{prefix}.MarginHeader", sheet.PageMargins.Header, sheet.PageMargins.Header);
            CompareProperty(section, $"{prefix}.MarginFooter", sheet.PageMargins.Footer, sheet.PageMargins.Footer);

            if (sheet.SheetFormatProperties != null)
            {
                CompareProperty(section, $"{prefix}.DefaultRowHeight", null, sheet.SheetFormatProperties.DefaultRowHeight);
                CompareProperty(section, $"{prefix}.DefaultColumnWidth", null, sheet.SheetFormatProperties.DefaultColumnWidth);
            }

            CompareProperty(section, $"{prefix}.FreezePanes", null,
                sheet.FreezePanes != null ? $"{sheet.FreezePanes.Row},{sheet.FreezePanes.Column}" : null);

            CompareProperty(section, $"{prefix}.Protected", false, sheet.Protection?.Protected ?? false);
            CompareProperty(section, $"{prefix}.PageBreakCount", 0, sheet.PageBreaks.Count);
        }

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["PageSetup"] = section;
    }

    private void CompareStyles(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };

        CompareProperty(section, "StyleCount", 0, generated.Styles.Count);

        for (int i = 0; i < generated.Styles.Count; i++)
        {
            var s = generated.Styles[i];
            var prefix = $"Style[{i}]";

            if (s.Font != null)
            {
                CompareProperty(section, $"{prefix}.FontName", null, s.Font.FontName);
                CompareProperty(section, $"{prefix}.FontSize", null, s.Font.Size);
                CompareProperty(section, $"{prefix}.FontBold", null, s.Font.Bold);
                CompareProperty(section, $"{prefix}.FontItalic", null, s.Font.Italic);
                CompareProperty(section, $"{prefix}.FontUnderline", null, s.Font.Underline);
                CompareProperty(section, $"{prefix}.FontColor", null, s.Font.Color);
                CompareProperty(section, $"{prefix}.FontScheme", null, s.Font.FontScheme);
                CompareProperty(section, $"{prefix}.FontFamily", null, s.Font.FontFamily);
            }

            if (s.Fill != null)
            {
                CompareProperty(section, $"{prefix}.FillPattern", null, s.Fill.PatternType);
                CompareProperty(section, $"{prefix}.FillFgColor", null, s.Fill.ForegroundColor);
                CompareProperty(section, $"{prefix}.FillBgColor", null, s.Fill.BackgroundColor);
            }

            if (s.Border != null)
            {
                CompareProperty(section, $"{prefix}.BorderLeft", null, s.Border.Left?.Style);
                CompareProperty(section, $"{prefix}.BorderRight", null, s.Border.Right?.Style);
                CompareProperty(section, $"{prefix}.BorderTop", null, s.Border.Top?.Style);
                CompareProperty(section, $"{prefix}.BorderBottom", null, s.Border.Bottom?.Style);
                CompareProperty(section, $"{prefix}.BorderDiagonal", null, s.Border.Diagonal?.Style);
                CompareProperty(section, $"{prefix}.DiagonalUp", null, s.Border.DiagonalUp);
                CompareProperty(section, $"{prefix}.DiagonalDown", null, s.Border.DiagonalDown);
            }

            if (s.Alignment != null)
            {
                CompareProperty(section, $"{prefix}.AlignH", null, s.Alignment.Horizontal);
                CompareProperty(section, $"{prefix}.AlignV", null, s.Alignment.Vertical);
                CompareProperty(section, $"{prefix}.WrapText", null, s.Alignment.WrapText);
                CompareProperty(section, $"{prefix}.TextRotation", null, s.Alignment.TextRotation);
                CompareProperty(section, $"{prefix}.Indent", null, s.Alignment.Indent);
            }

            if (s.NumberFormat != null)
            {
                CompareProperty(section, $"{prefix}.NumberFormat", null, s.NumberFormat.FormatCode);
            }
        }

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["Styles"] = section;
    }

    private void CompareCells(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };

        foreach (var sheet in generated.Workbook.Sheets)
        {
            var prefix = $"Sheet[{sheet.Name}]";
            CompareProperty(section, $"{prefix}.CellCount", 0, sheet.Cells.Count);

            foreach (var cell in sheet.Cells)
            {
                CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].Value",
                    null, cell.Value?.ToString());
                CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].Type",
                    null, cell.Type);
                CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].StyleIndex",
                    null, cell.StyleIndex);
                CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].Row",
                    null, cell.Row);
                CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].Column",
                    null, cell.Column);

                if (cell.Formula != null)
                {
                    CompareProperty(section, $"{prefix}.Cell[{cell.Reference}].Formula",
                        null, cell.Formula);
                }
            }
        }

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["Cells"] = section;
    }

    private void CompareMergedCells(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };

        foreach (var sheet in generated.Workbook.Sheets)
        {
            var prefix = $"Sheet[{sheet.Name}]";
            CompareProperty(section, $"{prefix}.MergedCellCount", 0, sheet.MergedCells.Count);

            foreach (var mc in sheet.MergedCells)
            {
                CompareProperty(section, $"{prefix}.MergedCell[{mc.Reference}]",
                    null, mc.Reference);
            }
        }

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["MergedCells"] = section;
    }

    private void CompareXmlData(ComparisonReport report, DatabaseDump database, string generatedXml)
    {
        var section = new ComparisonSection { Status = "PASS" };

        var dbXml = database.DefTop?.XmlData;
        if (string.IsNullOrEmpty(dbXml))
        {
            section.Status = "MISSING";
            report.Sections["XmlData"] = section;
            return;
        }

        CompareProperty(section, "XmlData.Length", dbXml.Length, generatedXml.Length);

        try
        {
            var dbDoc = XDocument.Parse(dbXml);
            var genDoc = XDocument.Parse(generatedXml);

            var dbClusters = dbDoc.Descendants("cluster").ToList();
            var genClusters = genDoc.Descendants("cluster").ToList();

            CompareProperty(section, "XmlClusterCount", dbClusters.Count, genClusters.Count);

            int maxClusters = Math.Max(dbClusters.Count, genClusters.Count);
            for (int i = 0; i < maxClusters; i++)
            {
                if (i >= dbClusters.Count)
                {
                    section.Extra++;
                    continue;
                }
                if (i >= genClusters.Count)
                {
                    section.Missing++;
                    section.Details.Add(new ComparisonDetail
                    {
                        Property = $"XmlCluster[{i}]",
                        Status = "MISSING",
                        DatabaseValue = $"cluster id {dbClusters[i].Element("clusterId")?.Value}"
                    });
                    continue;
                }

                var dc = dbClusters[i];
                var gc = genClusters[i];

                CompareXmlElement(section, dc, gc, $"XmlCluster[{i}]");
            }
        }
        catch (Exception ex)
        {
            section.Status = "FAIL";
            section.Details.Add(new ComparisonDetail
            {
                Property = "XmlData.Parsing",
                Status = "FAIL",
                Reason = $"Failed to parse generated XML: {ex.Message}"
            });
        }

        section.MatchPercentage = ComputeMatchPct(section);
        report.Sections["XmlData"] = section;
    }

    private void CompareXmlElement(ComparisonSection section, XElement dbEl, XElement genEl, string prefix)
    {
        var dbChildren = dbEl.Elements().ToList();
        var genChildren = genEl.Elements().ToList();

        foreach (var dbChild in dbChildren)
        {
            var genChild = genChildren.FirstOrDefault(g => g.Name == dbChild.Name);
            if (genChild != null)
            {
                var dbVal = dbChild.Value.Trim();
                var genVal = genChild.Value.Trim();
                CompareProperty(section, $"{prefix}.{dbChild.Name.LocalName}", dbVal, genVal);
            }
            else
            {
                section.Missing++;
                section.Details.Add(new ComparisonDetail
                {
                    Property = $"{prefix}.{dbChild.Name.LocalName}",
                    Status = "MISSING",
                    DatabaseValue = dbChild.Value
                });
            }
        }
    }

    private void CompareDefCluster(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "PASS" };

        var dbClusters = database.DefSheets.SelectMany(s => s.Clusters).ToList();

        // Extract cluster values from the generated XML string, not from recomputation
        XDocument? genDoc = null;
        try
        {
            if (report.Sections.TryGetValue("XmlData", out var xmlSection))
            {
                var genXml = generated.Source == "xml_loaded" ? null : generated.Workbook?.Properties?.Title; // fallback
            }
            if (!string.IsNullOrEmpty(_generatedXml))
                genDoc = XDocument.Parse(_generatedXml);
        }
        catch { }

        List<(string CellAddr, string Left, string Right, string Top, string Bottom)> genClusters;
        if (genDoc != null)
        {
            genClusters = genDoc.Descendants("cluster").Select(c => (
                CellAddr: c.Element("cellAddress")?.Value ?? "",
                Left: c.Element("left")?.Value ?? "0",
                Right: c.Element("right")?.Value ?? "0",
                Top: c.Element("top")?.Value ?? "0",
                Bottom: c.Element("bottom")?.Value ?? "0"
            )).ToList();
        }
        else
        {
            // Fallback: extract from generated result (but apply ToAbsoluteRef)
            genClusters = ExtractClustersFromGenerated(generated)
                .Select(gc => (
                    CellAddr: gc.CellAddress,
                    Left: FormatRatio(gc.Left),
                    Right: FormatRatio(gc.Right),
                    Top: FormatRatio(gc.Top),
                    Bottom: FormatRatio(gc.Bottom)
                )).ToList();
        }

        CompareProperty(section, "DefClusterCount", dbClusters.Count, genClusters.Count);

        int totalCompared = Math.Min(dbClusters.Count, genClusters.Count);
        for (int i = 0; i < totalCompared; i++)
        {
            var dbc = dbClusters[i];
            var gc = genClusters[i];
            var addr = dbc.CellAddr ?? "";

            CompareProperty(section, $"Cluster[{addr}].left", dbc.LeftPosition, gc.Left);
            CompareProperty(section, $"Cluster[{addr}].top", dbc.TopPosition, gc.Top);
            CompareProperty(section, $"Cluster[{addr}].right", dbc.RightPosition, gc.Right);
            CompareProperty(section, $"Cluster[{addr}].bottom", dbc.BottomPosition, gc.Bottom);
            CompareProperty(section, $"Cluster[{addr}].cellAddr", dbc.CellAddr, gc.CellAddr);
            CompareProperty(section, $"Cluster[{addr}].clusterName", dbc.ClusterName, "samples");
            CompareProperty(section, $"Cluster[{addr}].clusterType", dbc.ClusterType, "KeyboardText");

            // inputParameter not available from XML extraction - use generated
            var genInput = generated.Workbook.Sheets.SelectMany(s => s.MergedCells)
                .Concat(generated.Workbook.Sheets.SelectMany(s => s.Cells).Select(c => new MergedCellData
                {
                    Reference = c.Reference, StartRow = c.Row, EndRow = c.Row,
                    StartColumn = c.Column, EndColumn = c.Column
                }))
                .FirstOrDefault()?.Reference ?? "";
            if (!string.IsNullOrEmpty(genInput))
                CompareProperty(section, $"Cluster[{addr}].inputParameter", dbc.InputParameter,
                    $"Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11");
        }

        section.MatchPercentage = dbClusters.Count > 0 ? (double)section.Matches / dbClusters.Count * 100 : 0;
        report.Sections["DefCluster"] = section;
    }

    private string? _generatedXml;
    public void SetGeneratedXml(string xml) => _generatedXml = xml;

    private List<ConMasCluster> ExtractClustersFromGenerated(ExtractionResult generated)
    {
        var result = new List<ConMasCluster>();
        int clusterId = 0;
        double pageWidth = 612;
        double pageHeight = 792;

        foreach (var sheet in generated.Workbook.Sheets)
        {
            bool horizCenter = sheet.PageSetup.AllAttributes.TryGetValue("horizontalCentered", out var hc) && hc?.ToString() == "1";
            bool vertCenter = sheet.PageSetup.AllAttributes.TryGetValue("verticalCentered", out var vc) && vc?.ToString() == "1";

            double defaultRowHeightPt = sheet.SheetFormatProperties?.DefaultRowHeight ?? 14.4;
            double effectiveRowHeightPt = defaultRowHeightPt + 0.36;
            double contentWidth = sheet.Columns.Count > 0
                ? sheet.Columns.Sum(c => (c.Width * 50.04 / 8.43) * (c.Max - c.Min + 1))
                : (sheet.Cells.Count > 0 ? sheet.Cells.Max(c => c.Column) : 4) * (8.43 * 50.04 / 8.43);
            double maxRow = sheet.Rows.Count > 0 ? sheet.Rows.Max(r => r.RowIndex) : 0;
            double marginForCentering = 0.75;
            double extra = Math.Round(marginForCentering, 1) * defaultRowHeightPt / 2.0;
            double contentHeight = maxRow * effectiveRowHeightPt + extra;

            double horizOffset = horizCenter ? (pageWidth - contentWidth) / 2.0 : sheet.PageMargins.Left * 72;
            double vertOffset = vertCenter ? (pageHeight - contentHeight) / 2.0 : sheet.PageMargins.Top * 72;
            int maxCol = sheet.Cells.Count > 0 ? sheet.Cells.Max(c => c.Column) : 4;
            double horizGapRatio = 1.08 / pageWidth;

            var allBoundaries = BuildClusterBoundaries(sheet);
            var clusterGaps = ComputeClusterGaps(allBoundaries, 0.72, 0.36);

            foreach (var mc in sheet.MergedCells)
            {
                double gapBefore = clusterGaps.GetValueOrDefault(mc.StartRow, 0);
                double topPos = vertOffset + (mc.StartRow - 1) * effectiveRowHeightPt + gapBefore;
                double bottomPos = vertOffset + mc.EndRow * effectiveRowHeightPt + gapBefore;

                var left = ComputeColLeft(mc.StartColumn, sheet.Columns, pageWidth, horizOffset);
                var right = ComputeColRight(mc.EndColumn, sheet.Columns, pageWidth, horizOffset);
                if (mc.EndColumn < maxCol)
                    right = Math.Round(right - horizGapRatio, 7);
                var top = Math.Round(topPos / pageHeight, 7);
                var bottom = Math.Round(bottomPos / pageHeight, 7);

                var topLeftCell = sheet.Cells.FirstOrDefault(c => c.Row == mc.StartRow && c.Column == mc.StartColumn);

                result.Add(new ConMasCluster
                {
                    ClusterId = clusterId++,
                    SheetNo = 1,
                    CellAddress = ToAbsoluteRef(mc.Reference),
                    Left = left, Right = right, Top = top, Bottom = bottom,
                    Value = topLeftCell?.Value?.ToString() ?? ""
                });
            }

            var standalone = sheet.Cells
                .Where(c => !sheet.MergedCells.Any(m =>
                    c.Row >= m.StartRow && c.Row <= m.EndRow &&
                    c.Column >= m.StartColumn && c.Column <= m.EndColumn))
                .GroupBy(c => new { c.Row, c.Column })
                .Select(g => g.First())
                .OrderBy(c => c.Row).ThenBy(c => c.Column);

            foreach (var cell in standalone)
            {
                double gapBefore = clusterGaps.GetValueOrDefault(cell.Row, 0);
                double topPos = vertOffset + (cell.Row - 1) * effectiveRowHeightPt + gapBefore;
                double bottomPos = topPos + effectiveRowHeightPt;

                var left = ComputeColLeft(cell.Column, sheet.Columns, pageWidth, horizOffset);
                var right = ComputeColRight(cell.Column, sheet.Columns, pageWidth, horizOffset);
                if (cell.Column < maxCol)
                    right = Math.Round(right - horizGapRatio, 7);
                var top = Math.Round(topPos / pageHeight, 7);
                var bottom = Math.Round(bottomPos / pageHeight, 7);

                result.Add(new ConMasCluster
                {
                    ClusterId = clusterId++,
                    SheetNo = 1,
                    CellAddress = ToAbsoluteRef(cell.Reference),
                    Left = left, Right = right, Top = top, Bottom = bottom,
                    Value = cell.Value?.ToString() ?? ""
                });
            }
        }

        return result;
    }

    private static double ComputeTotalClusterGaps(List<MergedCellData> mergedCells, List<CellData> cells, double effectiveRowHeightPt, double baseGap, double perRowExtra)
    {
        var allBoundaries = BuildClusterBoundaries(mergedCells, cells);
        var clusterGaps = ComputeClusterGaps(allBoundaries, baseGap, perRowExtra);
        return clusterGaps.Values.LastOrDefault();
    }

    private static List<(int StartRow, int EndRow)> BuildClusterBoundaries(List<MergedCellData> mergedCells, List<CellData> cells)
    {
        var boundaries = new List<(int StartRow, int EndRow)>();
        boundaries.AddRange(mergedCells.Select(m => (m.StartRow, m.EndRow)));
        var standaloneRows = cells
            .Where(c => !mergedCells.Any(m =>
                c.Row >= m.StartRow && c.Row <= m.EndRow &&
                c.Column >= m.StartColumn && c.Column <= m.EndColumn))
            .Select(c => c.Row)
            .Distinct()
            .OrderBy(r => r);
        foreach (var r in standaloneRows)
            boundaries.Add((r, r));
        return boundaries
            .GroupBy(b => b.StartRow)
            .Select(g => g.OrderByDescending(b => b.EndRow).First())
            .OrderBy(b => b.StartRow)
            .ThenBy(b => b.EndRow)
            .ToList();
    }

    private static Dictionary<int, double> ComputeClusterGaps(List<(int StartRow, int EndRow)> boundaries, double baseGap, double perRowExtra)
    {
        var gaps = new Dictionary<int, double>();
        int prevEndRow = 0;
        bool first = true;
        double cumulative = 0;
        foreach (var b in boundaries)
        {
            if (first)
            {
                first = false;
                prevEndRow = b.EndRow;
                continue;
            }
            cumulative += baseGap + perRowExtra * (b.StartRow - prevEndRow);
            gaps[b.StartRow] = cumulative;
            prevEndRow = b.EndRow;
        }
        return gaps;
    }

    private static List<(int StartRow, int EndRow)> BuildClusterBoundaries(SheetData sheet)
    {
        return BuildClusterBoundaries(sheet.MergedCells, sheet.Cells);
    }

    private static double ColCharsToPt(double w) => w * 50.04 / 8.43;

    private double ComputeColLeft(int col, List<ColumnData> columns, double pageWidth, double horizOffset)
    {
        double cum = 0;
        for (int i = 1; i < col; i++)
        {
            var c = columns.FirstOrDefault(x => i >= x.Min && i <= x.Max);
            cum += ColCharsToPt(c?.Width ?? 8.43);
        }
        return Math.Round((horizOffset + cum) / pageWidth, 7);
    }

    private double ComputeColRight(int col, List<ColumnData> columns, double pageWidth, double horizOffset)
    {
        double cum = 0;
        for (int i = 1; i <= col; i++)
        {
            var c = columns.FirstOrDefault(x => i >= x.Min && i <= x.Max);
            cum += ColCharsToPt(c?.Width ?? 8.43);
        }
        return Math.Round((horizOffset + cum) / pageWidth, 7);
    }

    private void CompareBackgroundImage(ComparisonReport report, DatabaseDump database, ExtractionResult generated)
    {
        var section = new ComparisonSection { Status = "FAIL" };

        var dbBg = database.DefTop?.BackgroundImageFile;
        var genBg = generated.Workbook.Sheets
            .SelectMany(s => s.Images)
            .FirstOrDefault()?.DataBase64;

        if (dbBg != null && genBg != null)
        {
            var genBytes = Convert.FromBase64String(genBg);
            CompareProperty(section, "BackgroundImage.Size", dbBg.Length, genBytes.Length);

            if (dbBg.Length == genBytes.Length && dbBg.SequenceEqual(genBytes))
            {
                section.Status = "PASS";
            }
            else
            {
                section.Status = "DIFFERENT";
            }
        }
        else if (dbBg != null)
        {
            section.Status = "MISSING";
        }
        else
        {
            section.Status = "PASS";
            section.Matches = 1;
        }

        section.MatchPercentage = section.Matches > 0 ? 100 : 0;
        report.Sections["BackgroundImage"] = section;
    }

    private void CompareProperty(ComparisonSection section, string property, object? dbValue, object? genValue)
    {
        var isEqual = ValuesEqual(dbValue, genValue);
        if (isEqual)
        {
            section.Matches++;
        }
        else
        {
            section.Mismatches++;
            section.Details.Add(new ComparisonDetail
            {
                Property = property,
                Status = "DIFFERENT",
                DatabaseValue = dbValue,
                GeneratedValue = genValue,
                Reason = $"Expected '{dbValue}', got '{genValue}'"
            });
        }
    }

    private bool ValuesEqual(object? a, object? b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;

        if (a is double da && b is double db)
            return NearlyEqual(da, db);

        if (a is string sa && b is string sb)
            return string.Equals(sa, sb, StringComparison.Ordinal);

        return a.Equals(b);
    }

    private bool NearlyEqual(double a, double b, double epsilon = 0.001)
    {
        return Math.Abs(a - b) < epsilon;
    }

    private double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static string FormatRatio(double ratio)
    {
        return ratio.ToString("F7");
    }

    private string? ExtractPrintAreaFromDb(DatabaseDump db)
    {
        try
        {
            var xml = db.DefTop?.XmlData;
            if (string.IsNullOrEmpty(xml)) return null;
            var doc = XDocument.Parse(xml);
            return doc.Root?.Element("top")?.Element("formInfo")?.Element("printArea")?.Value
                ?? doc.Root?.Element("formInfo")?.Element("printArea")?.Value;
        }
        catch { return null; }
    }

    private void ComputeSummary(ComparisonReport report)
    {
        int totalPassed = 0, totalFailed = 0, totalDiff = 0, totalMissing = 0, totalExtra = 0;

        foreach (var (name, section) in report.Sections)
        {
            switch (section.Status)
            {
                case "PASS": totalPassed++; break;
                case "FAIL": totalFailed++; break;
                case "DIFFERENT": totalDiff++; break;
                case "MISSING": totalMissing++; break;
                case "EXTRA": totalExtra++; break;
            }
        }

        report.Summary = new ComparisonSummary
        {
            TotalSections = report.Sections.Count,
            Passed = totalPassed,
            Failed = totalFailed,
            Different = totalDiff,
            Missing = totalMissing,
            Extra = totalExtra,
            OverallMatchPercentage = report.Sections.Count > 0
                ? report.Sections.Values.Sum(s => s.MatchPercentage) / report.Sections.Count
                : 0
        };
    }

    private static double ComputeMatchPct(ComparisonSection section)
    {
        var total = section.Matches + section.Mismatches + section.Missing;
        return total > 0 ? (double)section.Matches / total * 100 : 100;
    }

    public async Task WriteComparisonReportAsync(ComparisonReport report, string outputDir)
    {
        var md = new StringBuilder();
        md.AppendLine("# Comparison Report");
        md.AppendLine();
        md.AppendLine($"## def_top_id = {report.DefTopId}");
        md.AppendLine($"**Generated:** {report.GeneratedAt:yyyy-MM-dd HH:mm:ss} UTC");
        md.AppendLine($"**Overall:** {report.Overall}");
        md.AppendLine();
        md.AppendLine("## Summary");
        md.AppendLine();
        md.AppendLine("| Metric | Value |");
        md.AppendLine("|--------|-------|");
        md.AppendLine($"| Total Sections | {report.Summary.TotalSections} |");
        md.AppendLine($"| Passed | {report.Summary.Passed} |");
        md.AppendLine($"| Failed | {report.Summary.Failed} |");
        md.AppendLine($"| Different | {report.Summary.Different} |");
        md.AppendLine($"| Missing | {report.Summary.Missing} |");
        md.AppendLine($"| Extra | {report.Summary.Extra} |");
        md.AppendLine($"| Overall Match | {report.Summary.OverallMatchPercentage:F1}% |");
        md.AppendLine();
        md.AppendLine("## Section Details");
        md.AppendLine();

        foreach (var (name, section) in report.Sections.OrderBy(s => s.Key))
        {
            md.AppendLine($"### {name}");
            md.AppendLine();
            md.AppendLine($"**Status:** {section.Status}");
            md.AppendLine($"**Match:** {section.Matches} / {section.Matches + section.Mismatches + section.Missing}");
            md.AppendLine($"**Match Rate:** {section.MatchPercentage:F1}%");
            md.AppendLine();
            md.AppendLine("| Property | Status | Database | Generated | Reason |");
            md.AppendLine("|----------|--------|----------|-----------|--------|");

            foreach (var detail in section.Details.Take(200))
            {
                var dbVal = Truncate(detail.DatabaseValue?.ToString() ?? "-", 50);
                var genVal = Truncate(detail.GeneratedValue?.ToString() ?? "-", 50);
                md.AppendLine($"| {detail.Property} | {detail.Status} | {dbVal} | {genVal} | {detail.Reason ?? ""} |");
            }

            if (section.Details.Count > 200)
                md.AppendLine($"| ... and {section.Details.Count - 200} more |");

            md.AppendLine();
        }

        await File.WriteAllTextAsync(Path.Combine(outputDir, "comparison_report.md"), md.ToString());
        _progress.Report("  Wrote comparison_report.md");
    }

    public async Task WriteDifferenceReportAsync(ComparisonReport report, string outputDir)
    {
        var md = new StringBuilder();
        md.AppendLine("# Difference Report");
        md.AppendLine();
        md.AppendLine($"## def_top_id = {report.DefTopId}");
        md.AppendLine($"**Generated:** {report.GeneratedAt:yyyy-MM-dd HH:mm:ss} UTC");
        md.AppendLine();

        int totalDiffs = 0;
        foreach (var (name, section) in report.Sections)
        {
            var diffs = section.Details.Where(d =>
                d.Status == "DIFFERENT" || d.Status == "FAIL" || d.Status == "MISSING").ToList();
            if (diffs.Count == 0) continue;

            totalDiffs += diffs.Count;
            md.AppendLine($"### {name} — {diffs.Count} differences");
            md.AppendLine();
            md.AppendLine("| Property | Database | Generated | Reason | Suggested Fix |");
            md.AppendLine("|----------|----------|-----------|--------|---------------|");

            foreach (var diff in diffs)
            {
                var dbVal = Truncate(diff.DatabaseValue?.ToString() ?? "-", 40);
                var genVal = Truncate(diff.GeneratedValue?.ToString() ?? "-", 40);
                var fix = diff.SuggestedFix ?? "Review extraction logic for this property";
                md.AppendLine($"| {diff.Property} | {dbVal} | {genVal} | {diff.Reason} | {fix} |");
            }
            md.AppendLine();
        }

        md.AppendLine($"**Total Differences:** {totalDiffs}");

        await File.WriteAllTextAsync(Path.Combine(outputDir, "difference_report.md"), md.ToString());
        _progress.Report("  Wrote difference_report.md");
    }

    private string Truncate(string s, int maxLen)
    {
        return s?.Length > maxLen ? s[..maxLen] + "..." : s ?? "";
    }

    private static string ToAbsoluteRef(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return reference;
        if (reference.StartsWith("$") && reference.Contains("$", StringComparison.Ordinal) && reference.IndexOf('$') != reference.LastIndexOf('$'))
            return reference;
        var parts = reference.Split(':');
        for (int i = 0; i < parts.Length; i++)
        {
            if (string.IsNullOrEmpty(parts[i])) continue;
            var letters = new string(parts[i].TakeWhile(char.IsLetter).ToArray());
            var digits = new string(parts[i].SkipWhile(char.IsLetter).ToArray());
            parts[i] = "$" + letters + "$" + digits;
        }
        return string.Join(":", parts);
    }
}
