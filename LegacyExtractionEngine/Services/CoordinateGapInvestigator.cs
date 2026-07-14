using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;
using LegacyExtractionEngine.Services.Importer;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 5 Investigation: Reverse-engineer the final coordinate transformation.
/// 
/// Current findings:
///   - Template 546, column C:  4.08 pt offset between COM Left and OpenXML sum
///   - Template 547, rows 7-10: ~4.76 pt/row drift in DB Top vs COM Top
/// 
/// This tool dumps every measurable Excel property to find the EXACT source
/// of these values — gridlines, cell padding, borders, printer scaling, etc.
/// </summary>
public class CoordinateGapInvestigator
{
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;

    public CoordinateGapInvestigator(IProgress<string> progress)
    {
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test");
    }

    public async Task InvestigateAllAsync()
    {
        _progress.Report("=== Phase 5: Coordinate Gap Investigation ===");
        _progress.Report("Objective: Find exact source of 4.08pt column gap and 4.76pt/row drift");
        _progress.Report("");

        await Investigate546ColumnsAsync();
        await Investigate547RowsAsync();
        await GenerateGapReportAsync();

        _progress.Report("=== Investigation Complete ===");
    }

    private async Task Investigate546ColumnsAsync()
    {
        _progress.Report("--- Template 546: Column Investigation ---");
        var wbPath = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx";

        // Get OpenXML column data
        var openXmlCols = new List<OpenXmlColInfo>();
        try
        {
            using var doc = SpreadsheetDocument.Open(wbPath, false);
            var wbPart = doc.WorkbookPart;
            var wsPart = wbPart?.WorksheetParts.FirstOrDefault();
            if (wsPart != null)
            {
                var wsXml = XDocument.Load(wsPart.GetStream());
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var colsEl = wsXml.Descendants(ns + "cols").FirstOrDefault();
                if (colsEl != null)
                {
                    foreach (var col in colsEl.Elements(ns + "col"))
                    {
                        openXmlCols.Add(new OpenXmlColInfo
                        {
                            Min = (int)col.Attribute("min"),
                            Max = (int)col.Attribute("max"),
                            Width = (double)col.Attribute("width"),
                            CustomWidth = (bool?)col.Attribute("customWidth")
                        });
                    }
                }
                else
                {
                    // No cols element — all default width
                    _progress.Report("  No <cols> element found — all columns at default width");
                }

                // Get default column width from sheet format properties
                var sheetFormat = wsXml.Descendants(ns + "sheetFormatPr").FirstOrDefault();
                if (sheetFormat != null)
                {
                    var defColWidth = (double?)sheetFormat.Attribute("defaultColWidth");
                    _progress.Report($"  SheetFormatPr defaultColWidth: {defColWidth}");
                }

                // Get the workbook's default font width
                var stylesPart = wbPart?.WorkbookStylesPart;
                if (stylesPart != null)
                {
                    var sx = XDocument.Load(stylesPart.GetStream());
                    var fonts = sx.Descendants(ns + "fonts").FirstOrDefault();
                    var fontCount = (int?)fonts?.Attribute("count");
                    _progress.Report($"  Font count: {fontCount}");

                    var firstFont = fonts?.Elements(ns + "font").FirstOrDefault();
                    if (firstFont != null)
                    {
                        var sz = firstFont.Element(ns + "sz");
                        var name = firstFont.Element(ns + "name");
                        if (sz != null) _progress.Report($"  Default font size: {sz.Attribute("val")?.Value}");
                        if (name != null) _progress.Report($"  Default font name: {name.Attribute("val")?.Value}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _progress.Report($"  OpenXML read error: {ex.Message}");
        }

        // Get COM column data
        var comCols = new List<ComColInfo>();
        dynamic? excel = null;
        dynamic? wb = null;
        try
        {
            _progress.Report("  Opening Excel via COM...");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) { _progress.Report("  ERROR: Excel COM not available"); return; }
            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            wb = excel.Workbooks.Open(wbPath);
            dynamic ws = wb.Sheets[1];

            _progress.Report("  Reading column properties via COM...");
            for (int colIdx = 1; colIdx <= 20; colIdx++)
            {
                try
                {
                    dynamic range = ws.Cells[1, colIdx];
                    dynamic entireCol = range.EntireColumn;
                    var ci = new ComColInfo
                    {
                        ColIndex = colIdx,
                        ColLetter = GetColumnLetter(colIdx),
                        RangeLeft = (double)range.Left,
                        RangeWidth = (double)range.Width,
                        ColumnWidth = (double)entireCol.ColumnWidth,
                    };

                    // Try more properties
                    try { ci.RowHeight = (double)range.RowHeight; } catch { }
                    try { ci.EntireRowHeight = (double)range.EntireRow.Height; } catch { }

                    // Get the column's left/width from EntireColumn
                    try { ci.EntireColLeft = (double)entireCol.Left; } catch { }
                    try { ci.EntireColWidth = (double)entireCol.Width; } catch { }

                    comCols.Add(ci);
                }
                catch { break; }
            }

            // Get page setup info
            dynamic pageSetup = ws.PageSetup;
            try
            {
                _progress.Report($"  PageSetup: LeftMargin={pageSetup.LeftMargin}, RightMargin={pageSetup.RightMargin}");
                _progress.Report($"  CenterHorizontally={pageSetup.CenterHorizontally}, CenterVertically={pageSetup.CenterVertically}");
                _progress.Report($"  Zoom={pageSetup.Zoom}, FitToPagesWide={pageSetup.FitToPagesWide}");
                _progress.Report($"  PaperSize={pageSetup.PaperSize}, Orientation={pageSetup.Orientation}");
            }
            catch { }

            // Check gridline/print settings
            try { _progress.Report($"  PrintGridlines={ws.PageSetup.PrintGridlines}"); } catch { }
            try { _progress.Report($"  DisplayGridlines={ws.DisplayGridlines}"); } catch { }

            // PointsToScreenPixels test
            try
            {
                var x = excel.ActiveWindow.PointsToScreenPixelsX(96);
                var y = excel.ActiveWindow.PointsToScreenPixelsY(0);
                _progress.Report($"  PointsToScreenPixelsX(96pt)={x}, PointsToScreenPixelsY(0pt)={y}");
            }
            catch (Exception ex) { _progress.Report($"  PointsToScreenPixels failed: {ex.Message}"); }
        }
        catch (Exception ex) { _progress.Report($"  COM Error: {ex.Message}"); }
        finally
        {
            if (wb != null) try { wb.Close(false); } catch { }
            if (excel != null) try { excel.Quit(); } catch { }
        }

        // Generate column report
        var sb = new StringBuilder();
        sb.AppendLine("# Template 546 — Column Coordinate Investigation");
        sb.AppendLine();
        sb.AppendLine("## Column Dump");
        sb.AppendLine();
        sb.AppendLine("| Col | Letter | OpenXML Width(chars) | COM ColumnWidth(chars) | COM Range.Left(pt) | COM Range.Width(pt) | COM EntireCol.Left(pt) | COM EntireCol.Width(pt) |");
        sb.AppendLine("|-----|--------|---------------------|----------------------|-------------------|--------------------|----------------------|-----------------------|");

        for (int i = 0; i < Math.Min(comCols.Count, 10); i++)
        {
            var cc = comCols[i];
            var oxCol = openXmlCols.FirstOrDefault(c => c.Min <= cc.ColIndex && c.Max >= cc.ColIndex);
            var oxW = oxCol?.Width.ToString("F6") ?? "(default 8.43)";
            sb.AppendLine($"| {cc.ColIndex} | {cc.ColLetter} | {oxW} | {cc.ColumnWidth:F6} | {cc.RangeLeft:F2} | {cc.RangeWidth:F2} | {cc.EntireColLeft:F2} | {cc.EntireColWidth:F2} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Column Width Conversions");
        sb.AppendLine();
        sb.AppendLine("Excel's internal column width conversion formula:");
        sb.AppendLine("- Character width → points: `points = (chars * 7 + 5) / 7 * font_width / 7 * 4`");
        sb.AppendLine("- Or more simply: points = chars * (max_digit_width * 256 + 12) / 256");
        sb.AppendLine();

        // Compute character-width conversion for each column
        sb.AppendLine("| Col | OpenXML chars | COM chars | COM Left | COM Width | OpenXML pts (50.04/8.43) | Derived factor |");
        sb.AppendLine("|-----|--------------|-----------|----------|-----------|------------------------|----------------|");
        for (int i = 0; i < Math.Min(comCols.Count, 10); i++)
        {
            var cc = comCols[i];
            var oxCol = openXmlCols.FirstOrDefault(c => c.Min <= cc.ColIndex && c.Max >= cc.ColIndex);
            double oxChars = oxCol?.Width ?? 8.43;
            double oxPts = oxChars * 50.04 / 8.43;
            double derivedFactor = cc.ColumnWidth > 0 ? cc.RangeWidth / cc.ColumnWidth : 0;
            sb.AppendLine($"| {cc.ColLetter} | {oxChars:F6} | {cc.ColumnWidth:F6} | {cc.RangeLeft:F2} | {cc.RangeWidth:F2} | {oxPts:F4} | {derivedFactor:F6} |");
        }

        // Compare cumulative left positions
        sb.AppendLine();
        sb.AppendLine("## Cumulative Left Position: COM vs OpenXML");
        sb.AppendLine();
        sb.AppendLine("| Col | COM Left | OpenXML Cumulative | Difference |");
        sb.AppendLine("|-----|----------|-------------------|------------|");

        double cumulativeOpenXml = 0;
        for (int i = 0; i < Math.Min(comCols.Count, 10); i++)
        {
            var cc = comCols[i];
            var oxCol = openXmlCols.FirstOrDefault(c => c.Min <= cc.ColIndex && c.Max >= cc.ColIndex);
            double oxChars = oxCol?.Width ?? 8.43;
            double oxPts = oxChars * 50.04 / 8.43;
            if (i > 0) cumulativeOpenXml += oxPts;
            double diff = cc.RangeLeft - cumulativeOpenXml;
            sb.AppendLine($"| {cc.ColLetter} | {cc.RangeLeft:F4} | {cumulativeOpenXml:F4} | {diff:+0.0000;-0.0000} |");
        }

        // Compute the EXACT difference at column C (index 3)
        sb.AppendLine();
        sb.AppendLine("## Column C (Index 3) — Root Cause Analysis");
        sb.AppendLine();
        double comLeftC = comCols.Count >= 3 ? comCols[2].RangeLeft : 0;
        double openXmlCumC = 0;
        for (int i = 0; i < 2 && i < comCols.Count; i++)
        {
            var cc = comCols[i];
            var oxCol = openXmlCols.FirstOrDefault(c => c.Min <= cc.ColIndex && c.Max >= cc.ColIndex);
            openXmlCumC += (oxCol?.Width ?? 8.43) * 50.04 / 8.43;
        }
        double gapC = comLeftC - openXmlCumC;
        sb.AppendLine($"COM Range.Left for Column C: {comLeftC:F4} pt");
        sb.AppendLine($"OpenXML cumulative (A+B): {openXmlCumC:F4} pt");
        sb.AppendLine($"Difference: {gapC:+0.0000;-0.0000} pt");
        sb.AppendLine();
        sb.AppendLine("### Possible Sources");
        sb.AppendLine($"1. Gridline width: ~{(gapC / 2):F4} pt per gridline between A-B and B-C");
        sb.AppendLine($"2. Cell padding (left + right): {gapC:F4} pt");
        sb.AppendLine($"3. Border thickness: check if columns have borders");
        sb.AppendLine($"4. Font metric conversion: COM vs OpenXML character width formula difference");

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "Template546", "column_gap_investigation.md"), sb.ToString());
        _progress.Report("  Wrote column_gap_investigation.md");
    }

    private async Task Investigate547RowsAsync()
    {
        _progress.Report("--- Template 547: Row Investigation ---");
        var wbPath = @"C:\Users\MCF-JOHNELLEEMPUERTO\AppData\Local\Temp\opencode\547_workbook_copy.xlsx";
        if (!File.Exists(wbPath))
        {
            var dir = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents";
            wbPath = Directory.GetFiles(dir, "*V3.1*", SearchOption.TopDirectoryOnly).FirstOrDefault() ?? "";
        }
        if (!File.Exists(wbPath)) { _progress.Report("  SKIP: workbook not found"); return; }

        // OpenXML row data
        var openXmlRows = new List<OpenXmlRowInfo>();
        try
        {
            using var doc = SpreadsheetDocument.Open(wbPath, false);
            var wbPart = doc.WorkbookPart;
            var wsPart = wbPart?.WorksheetParts.FirstOrDefault();
            if (wsPart != null)
            {
                var wsXml = XDocument.Load(wsPart.GetStream());
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                var rowsEl = wsXml.Descendants(ns + "sheetData").Elements(ns + "row");
                foreach (var row in rowsEl)
                {
                    openXmlRows.Add(new OpenXmlRowInfo
                    {
                        RowIndex = (int)row.Attribute("r"),
                        Height = (double?)row.Attribute("ht"),
                        CustomHeight = (bool?)row.Attribute("customHeight")
                    });
                }

                var sheetFormat = wsXml.Descendants(ns + "sheetFormatPr").FirstOrDefault();
                if (sheetFormat != null)
                {
                    _progress.Report($"  SheetFormatPr defaultRowHeight: {sheetFormat.Attribute("defaultRowHeight")?.Value}");
                }
            }
        }
        catch (Exception ex) { _progress.Report($"  OpenXML row read error: {ex.Message}"); }

        // COM row data
        var comRows = new List<ComRowInfo>();
        dynamic? excel = null;
        dynamic? wb = null;
        try
        {
            _progress.Report("  Opening Excel via COM for row investigation...");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) { _progress.Report("  ERROR: Excel COM not available"); return; }
            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            wb = excel.Workbooks.Open(wbPath);
            dynamic ws = wb.Sheets[1];

            _progress.Report("  Reading row properties via COM...");
            for (int rowIdx = 1; rowIdx <= 15; rowIdx++)
            {
                try
                {
                    dynamic range = ws.Cells[rowIdx, 1];  // column A for reference
                    dynamic entireRow = range.EntireRow;
                    var ri = new ComRowInfo
                    {
                        RowIndex = rowIdx,
                        RangeTop = (double)range.Top,
                        RangeHeight = (double)range.Height,
                        RowHeight = (double)entireRow.RowHeight,
                        EntireRowHeight = (double)entireRow.Height,
                    };

                    // Also read from the cluster column (I = index 9)
                    try
                    {
                        dynamic rangeI = ws.Cells[rowIdx, 9];
                        ri.RangeI_Top = (double)rangeI.Top;
                        ri.RangeI_Height = (double)rangeI.Height;
                    }
                    catch { }

                    comRows.Add(ri);
                }
                catch { break; }
            }

            // Print/display settings
            dynamic pageSetup = ws.PageSetup;
            try
            {
                _progress.Report($"  PageSetup: TopMargin={pageSetup.TopMargin}, BottomMargin={pageSetup.BottomMargin}");
                _progress.Report($"  Zoom={pageSetup.Zoom}, FitToPagesTall={pageSetup.FitToPagesTall}");
            }
            catch { }
            try { _progress.Report($"  PrintGridlines={ws.PageSetup.PrintGridlines}"); } catch { }
        }
        catch (Exception ex) { _progress.Report($"  COM Error: {ex.Message}"); }
        finally
        {
            if (wb != null) try { wb.Close(false); } catch { }
            if (excel != null) try { excel.Quit(); } catch { }
        }

        // Generate row report
        var sb = new StringBuilder();
        sb.AppendLine("# Template 547 — Row Coordinate Investigation");
        sb.AppendLine();
        sb.AppendLine("## Row Dump");
        sb.AppendLine();
        sb.AppendLine("| Row | COM Top(A) | COM Ht(A) | COM Top(I) | COM Ht(I) | RowHeight | OpenXML Ht | CustomHt |");
        sb.AppendLine("|-----|-----------|----------|-----------|----------|-----------|------------|----------|");

        foreach (var r in comRows.OrderBy(r => r.RowIndex))
        {
            var oxRow = openXmlRows.FirstOrDefault(o => o.RowIndex == r.RowIndex);
            var oxHt = oxRow?.Height?.ToString("F2") ?? "default";
            var cust = oxRow?.CustomHeight == true ? "yes" : "no";
            sb.AppendLine($"| {r.RowIndex} | {r.RangeTop:F2} | {r.RangeHeight:F2} | {r.RangeI_Top:F2} | {r.RangeI_Height:F2} | {r.RowHeight:F2} | {oxHt} | {cust} |");
        }

        // Compute per-row differences
        sb.AppendLine();
        sb.AppendLine("## Per-Row Difference Analysis");
        sb.AppendLine();
        sb.AppendLine("Comparing DB cluster positions vs COM Top:");
        sb.AppendLine();

        // We know from the DB: $I$6:$M$6 through $I$10:$M$10
        var dbClusters = new (int row, double dbTop, double dbBottom)[]
        {
            (6, 0.1654546, 0.1868182),
            (7, 0.1877273, 0.2095454),
            (8, 0.2104545, 0.2327273),
            (9, 0.2336364, 0.2559091),
            (10, 0.2568182, 0.2786364),
        };

        double printedOriginY = 65.04; // from first cluster derivation
        double pageHeight = 792;

        sb.AppendLine("| Row | COM Top | COM Top+Origin | DB Top(pts) | DB Top(pts) | Gap(pt) | Gap inc |");
        sb.AppendLine("|-----|---------|----------------|-------------|-------------|---------|---------|");

        double prevGap = 0;
        foreach (var (row, dbTop, _) in dbClusters)
        {
            var comRow = comRows.FirstOrDefault(r => r.RowIndex == row);
            if (comRow == null) continue;
            double comTop = comRow.RangeI_Top;
            double comTopPlusOrigin = comTop + printedOriginY;
            double dbTopPts = dbTop * pageHeight;
            double gap = dbTopPts - comTopPlusOrigin;
            double gapInc = gap - prevGap;
            sb.AppendLine($"| {row} | {comTop:F2} | {comTopPlusOrigin:F2} | {dbTopPts:F2} | {dbTop:F7} | {gap:+0.0000;-0.0000} | {gapInc:+0.0000;-0.0000} |");
            prevGap = gap;
        }

        // Check if gap = f(rowHeight, defaults, etc.)
        sb.AppendLine();
        sb.AppendLine("## Row Height vs Gap Analysis");
        sb.AppendLine();
        sb.AppendLine("Let's check if the gap increment equals something measurable:");
        sb.AppendLine();

        double gapSum = 0;
        int gapCount = 0;
        foreach (var (row, dbTop, _) in dbClusters.Skip(1)) // skip first which defines origin
        {
            var comRow = comRows.FirstOrDefault(r => r.RowIndex == row);
            if (comRow == null) continue;
            double comTop = comRow.RangeI_Top;
            double comTopPlusOrigin = comTop + printedOriginY;
            double dbTopPts = dbTop * pageHeight;
            double gap = dbTopPts - comTopPlusOrigin;
            gapSum += gap;
            gapCount++;

            sb.AppendLine($"Row {row}: Gap={gap:F4}pt, RowHeight={comRow.RowHeight:F4}pt, Gap/RowHeight={gap/comRow.RowHeight:F6}");
        }

        double avgGap = gapCount > 0 ? gapSum / gapCount : 0;
        sb.AppendLine();
        sb.AppendLine($"Average gap (rows 7-10): {avgGap:F4}pt");

        // Compare to known constants
        sb.AppendLine();
        sb.AppendLine("## Comparison with Known Constants");
        sb.AppendLine();
        double perRowExtraPt = 0.36;
        double clusterBaseGapPt = 0.72;
        sb.AppendLine($"PerRowExtraPt (from old engine): {perRowExtraPt} pt");
        sb.AppendLine($"ClusterBaseGapPt (from old engine): {clusterBaseGapPt} pt");
        sb.AppendLine($"Expected compounded gap: base + rowIdx * extra = {clusterBaseGapPt:F2} + N * {perRowExtraPt:F2}");

        // Check if the gap relates to default row height
        var row6 = comRows.FirstOrDefault(r => r.RowIndex == 6);
        if (row6 != null)
        {
            sb.AppendLine();
            sb.AppendLine($"Row 6 height: {row6.RowHeight:F4} pt (COM)");
            sb.AppendLine($"Row 6 height: {row6.EntireRowHeight:F4} pt (EntireRow)");
            double defaultHt = row6.RowHeight;
            sb.AppendLine($"If gap increment = defaultRowHeight * factor: {defaultHt * 0.36:F4} pt");

            // Check PerRowExtraPt as fraction of row height
            sb.AppendLine($"PerRowExtraPt / RowHeight = {perRowExtraPt / defaultHt:F6}");
        }

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "Template547", "row_gap_investigation.md"), sb.ToString());
        _progress.Report("  Wrote row_gap_investigation.md");
    }

    private async Task GenerateGapReportAsync()
    {
        _progress.Report("--- Generating Combined Gap Report ---");

        var sb = new StringBuilder();
        sb.AppendLine("# Phase 5: Coordinate Gap Analysis — Combined Report");
        sb.AppendLine();
        sb.AppendLine("## Current Status");
        sb.AppendLine();
        sb.AppendLine("| Gap | Value | Location | Status |");
        sb.AppendLine("|-----|-------|----------|--------|");
        sb.AppendLine("| Column offset (C) | 4.08 pt | Template 546, column C | Under investigation |");
        sb.AppendLine("| Row drift (7-10) | ~4.76 pt/row | Template 547, rows 7+ | Under investigation |");
        sb.AppendLine();
        sb.AppendLine("## Investigation Files");
        sb.AppendLine();
        sb.AppendLine("- `Template546/column_gap_investigation.md` — Column-level COM vs OpenXML dump");
        sb.AppendLine("- `Template547/row_gap_investigation.md` — Row-level COM vs DB comparison");
        sb.AppendLine();
        sb.AppendLine("## Next Steps");
        sb.AppendLine();
        sb.AppendLine("1. Read the investigation files above");
        sb.AppendLine("2. Determine if gaps come from: gridlines, cell padding, borders, font metrics, printer scaling");
        sb.AppendLine("3. Derive exact formula");
        sb.AppendLine("4. Update coordinate engine");
        sb.AppendLine("5. Validate both templates at 100%");

        await File.WriteAllTextAsync(Path.Combine(_outputDir, "coordinate_gap_report.md"), sb.ToString());
        _progress.Report("  Wrote coordinate_gap_report.md");
    }

    private static string GetColumnLetter(int colIndex)
    {
        if (colIndex < 1) return "";
        return ((char)('A' + colIndex - 1)).ToString();
    }

    private class OpenXmlColInfo
    {
        public int Min { get; set; }
        public int Max { get; set; }
        public double Width { get; set; }
        public bool? CustomWidth { get; set; }
    }

    private class ComColInfo
    {
        public int ColIndex { get; set; }
        public string ColLetter { get; set; } = "";
        public double RangeLeft { get; set; }
        public double RangeWidth { get; set; }
        public double ColumnWidth { get; set; }
        public double RowHeight { get; set; }
        public double EntireRowHeight { get; set; }
        public double EntireColLeft { get; set; }
        public double EntireColWidth { get; set; }
    }

    private class OpenXmlRowInfo
    {
        public int RowIndex { get; set; }
        public double? Height { get; set; }
        public bool? CustomHeight { get; set; }
    }

    private class ComRowInfo
    {
        public int RowIndex { get; set; }
        public double RangeTop { get; set; }
        public double RangeHeight { get; set; }
        public double RowHeight { get; set; }
        public double EntireRowHeight { get; set; }
        public double RangeI_Top { get; set; }
        public double RangeI_Height { get; set; }
    }
}
