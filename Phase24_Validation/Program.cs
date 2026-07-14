using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DDS = DocumentFormat.OpenXml.Drawing.Spreadsheet;

// Discover template paths
var templateDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\.."));
var formsDir = Path.Combine(templateDir, @"ExcelAPI\ExcelAPI\Forms");

var templates = new List<(int Id, string Path, string ComDumpPath)>();

// Template 546
var tpl546 = Path.Combine(templateDir, @"Investigation_546\original.xlsx");
var com546 = Path.Combine(templateDir, @"Test Folder Final Test\Template546\excel_com_dump.json");
if (File.Exists(tpl546)) templates.Add((546, tpl546, com546));

// Template 547 — search for V3.1 file in Documents
var docsDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
foreach (var f in Directory.GetFiles(docsDir, "*V3.1*.xlsx"))
{
    var com547 = Path.Combine(templateDir, @"Test Folder Final Test\Template547\excel_com_dump.json");
    templates.Add((547, f, com547));
    break;
}

// Template 548 — search Forms for "548" in filename
foreach (var f in Directory.GetFiles(formsDir, "*.xlsx"))
{
    var name = Path.GetFileNameWithoutExtension(f);
    if (name.Contains("548"))
    {
        var com548 = Path.Combine(templateDir, @"Test Folder Final Test\Template548\excel_com_dump.json");
        templates.Add((548, f, com548));
        break;
    }
}

if (templates.Count == 0)
{
    Console.WriteLine("No templates found. Exiting.");
    return;
}

int totalPass = 0, totalFail = 0;

foreach (var (id, path, comDumpPath) in templates)
{
    Console.WriteLine($"\n{'=',60}");
    Console.WriteLine($"  TEMPLATE {id}");
    Console.WriteLine($"  XLSX: {path}");
    Console.WriteLine($"  COM baseline: {(File.Exists(comDumpPath) ? comDumpPath : "(not found)")}");
    Console.WriteLine($"{'=',60}");

    if (!File.Exists(path))
    {
        Console.WriteLine($"  [SKIP] File not found: {path}");
        continue;
    }

    var ox = new OpenXmlExtractor(path);
    ox.ExtractAll();
    ox.PrintAll();

    if (File.Exists(comDumpPath))
    {
        var (p, f) = Compare(id, ox, File.ReadAllText(comDumpPath));
        totalPass += p;
        totalFail += f;
    }
    else
    {
        Console.WriteLine($"\n  [WARN] No COM dump — OpenXML values shown for manual inspection");
    }
}

Console.WriteLine($"\n{'=',60}");
Console.WriteLine($"  TOTAL: {totalPass} passed, {totalFail} failed");
Console.WriteLine($"{'=',60}");

// ── Comparison function ──
static (int pass, int fail) Compare(int templateId, OpenXmlExtractor ox, string comJson)
{
    using var comDoc = JsonDocument.Parse(comJson);
    var root = comDoc.RootElement;

    int pass = 0, fail = 0;
    void Check(string label, bool ok, string? detail = null)
    {
        if (ok) { Console.WriteLine($"  [PASS] {label}"); pass++; }
        else { Console.WriteLine($"  [FAIL] {label}" + (detail is not null ? $" — {detail}" : "")); fail++; }
    }

    // ── PageSetup ──
    var comPs = root.GetProperty("PageSetup");

    string? comPrintArea = comPs.TryGetProperty("PrintArea", out var pa) && pa.ValueKind == JsonValueKind.String
        ? pa.GetString() : null;
    Check("PrintArea", ox.PrintArea == comPrintArea || (ox.PrintArea is null && comPrintArea is null),
        $"OpenXML='{ox.PrintArea}' COM='{comPrintArea}'");

    if (ox.PageSetup is not null)
    {
        var ps = ox.PageSetup.Value;

        double comL = comPs.GetProperty("LeftMargin").GetDouble();
        double comR = comPs.GetProperty("RightMargin").GetDouble();
        double comT = comPs.GetProperty("TopMargin").GetDouble();
        double comB = comPs.GetProperty("BottomMargin").GetDouble();
            double comH = comPs.GetProperty("HeaderMargin").GetDouble();
        double comF = comPs.GetProperty("FooterMargin").GetDouble();

        Check("LeftMargin", Math.Abs(ps.LeftMargin * 72.0 - comL) < 0.5, $"OpenXML={ps.LeftMargin * 72:F2}pt COM={comL:F2}pt");
        Check("RightMargin", Math.Abs(ps.RightMargin * 72 - comR) < 0.5, $"OpenXML={ps.RightMargin * 72:F2}pt COM={comR:F2}pt");
        Check("TopMargin", Math.Abs(ps.TopMargin * 72 - comT) < 0.5, $"OpenXML={ps.TopMargin * 72:F2}pt COM={comT:F2}pt");
        Check("BottomMargin", Math.Abs(ps.BottomMargin * 72 - comB) < 0.5, $"OpenXML={ps.BottomMargin * 72:F2}pt COM={comB:F2}pt");
        Check("HeaderMargin", Math.Abs(ps.HeaderMargin * 72 - comH) < 0.5, $"OpenXML={ps.HeaderMargin * 72:F2}pt COM={comH:F2}pt");
        Check("FooterMargin", Math.Abs(ps.FooterMargin * 72 - comF) < 0.5, $"OpenXML={ps.FooterMargin * 72:F2}pt COM={comF:F2}pt");

        Check("CenterHorizontally", ps.CenterHorizontally == comPs.GetProperty("CenterHorizontally").GetBoolean());
        Check("CenterVertically", ps.CenterVertically == comPs.GetProperty("CenterVertically").GetBoolean());

        int comOrient = comPs.GetProperty("Orientation").GetInt32();
        string comOrientStr = comOrient == 2 ? "Landscape" : "Portrait";
        Check("Orientation", ps.Orientation == comOrientStr, $"OpenXML={ps.Orientation} COM={comOrientStr}");

        Check("PageWidth", Math.Abs(ps.PageWidth - comPs.GetProperty("PageWidth").GetDouble()) < 1,
            $"OpenXML={ps.PageWidth:F0} COM={comPs.GetProperty("PageWidth").GetDouble():F0}");
        Check("PageHeight", Math.Abs(ps.PageHeight - comPs.GetProperty("PageHeight").GetDouble()) < 1,
            $"OpenXML={ps.PageHeight:F0} COM={comPs.GetProperty("PageHeight").GetDouble():F0}");

        Check("PaperSize", ps.PaperSize == comPs.GetProperty("PaperSize").GetInt32());
        Check("FitToPagesWide", (int)ps.FitToPagesWide == comPs.GetProperty("FitToPagesWide").GetInt32());
        Check("FitToPagesTall", (int)ps.FitToPagesTall == comPs.GetProperty("FitToPagesTall").GetInt32());
    }
    else
    {
        fail++;
        Console.WriteLine($"  [FAIL] PageSetup — OpenXML returned null");
    }

    // Cluster comparison (informational)
    if (root.TryGetProperty("Clusters", out var comClusters))
    {
        Console.WriteLine($"  [INFO] COM has {comClusters.GetArrayLength()} clusters");

        foreach (var comC in comClusters.EnumerateArray())
        {
            string addr = comC.GetProperty("Address").GetString() ?? "";
            bool found = ox.MergeCells.Any(m =>
            {
                var mUpper = m.ToUpperInvariant();
                var aUpper = addr.ToUpperInvariant();
                return mUpper == aUpper || mUpper.Split(':')[0] == aUpper;
            });
            found |= ox.Comments.Any(c => c.Address.ToUpperInvariant() == addr.ToUpperInvariant());

            Console.WriteLine($"    {addr}: {(found ? "found in OpenXML" : "pixel-only, not in OpenXML")}");
        }
    }

    // Coordinate calculation
    if (ox.PageSetup is not null && ox.PrintAreaParsed is not null)
    {
        var ps = ox.PageSetup.Value;
        var printArea = ox.PrintAreaParsed.Value;

        double contentWidth = 0;
        for (int c = printArea.StartCol - 1; c < printArea.EndCol && c < ox.ColumnWidths.Count; c++)
            contentWidth += ox.ColumnWidths[c];

        double contentHeight = 0;
        foreach (var rh in ox.RowHeights)
            if (rh.Row >= printArea.StartRow && rh.Row <= printArea.EndRow)
                contentHeight += rh.Height;

        double originX = ps.CenterHorizontally
            ? (ps.PageWidth - contentWidth) / 2.0
            : ps.LeftMargin * 72;

        double originY = ps.CenterVertically
            ? (ps.PageHeight - contentHeight) / 2.0
            : ps.TopMargin * 72;

        Console.WriteLine($"  Coordinate calculation:");
        Console.WriteLine($"    ContentWidth={contentWidth:F2}pt ContentHeight={contentHeight:F2}pt");
        Console.WriteLine($"    OriginX={originX:F2}pt OriginY={originY:F2}pt");
    }

    Console.WriteLine($"\n  {'-',50}");
    Console.WriteLine($"  RESULTS: {pass} passed, {fail} failed");
    Console.WriteLine($"  {'-',50}");

    return (pass, fail);
}

// ── Data records ──
record struct PageSetupData(
    double LeftMargin, double RightMargin, double TopMargin, double BottomMargin,
    double HeaderMargin, double FooterMargin,
    bool CenterHorizontally, bool CenterVertically,
    int PaperSize, string Orientation,
    double PageWidth, double PageHeight,
    double FitToPagesWide, double FitToPagesTall
);

// ── OpenXML Extractor ──
class OpenXmlExtractor
{
    readonly string _path;
    public string FilePath => _path;

    public string? PrintArea { get; private set; }
    public (int StartCol, int StartRow, int EndCol, int EndRow)? PrintAreaParsed { get; private set; }
    public PageSetupData? PageSetup { get; private set; }
    public List<double> ColumnWidths { get; } = new();
    public List<(int Row, double Height)> RowHeights { get; } = new();
    public List<string> MergeCells { get; } = new();
    public List<(string Address, string Text)> Comments { get; } = new();
    public List<string> ImageAnchors { get; } = new();
    public List<string> SheetNames { get; } = new();

    public OpenXmlExtractor(string path) => _path = path;

    public void ExtractAll()
    {
        using var doc = SpreadsheetDocument.Open(_path, false);
        var wbPart = doc.WorkbookPart;
        if (wbPart is null) return;

        var allSheets = wbPart.Workbook.Descendants<Sheet>().ToList();
        foreach (var s in allSheets)
            SheetNames.Add(s.Name?.Value ?? "?");

        // Read all sheets, prefer non-_Fields sheet
        foreach (var sheet in allSheets)
        {
            string sheetName = sheet.Name?.Value ?? "";
            if (sheetName == "_Fields") continue;

            var wsPart = wbPart.GetPartById(sheet.Id!) as WorksheetPart;
            var ws = wsPart?.Worksheet;
            if (ws is null) continue;

            // Print Area
            PrintArea = ReadPrintArea(wbPart, sheetName);
            if (PrintArea is not null)
                PrintAreaParsed = ParsePrintArea(PrintArea);

            // PageSetup
            var ps = ws.Descendants<PageSetup>().FirstOrDefault();
            var pm = ws.Descendants<PageMargins>().FirstOrDefault();
            var po = ws.Descendants<PrintOptions>().FirstOrDefault();

            if (ps is not null && pm is not null)
            {
                DebugMargins(pm);

                double paperW = 612, paperH = 792;
                if (ps.PaperSize?.Value is not null)
                    (paperW, paperH) = GetPaperDimensions((int)ps.PaperSize.Value);

                string orient = "Portrait";
                if (ps.Orientation?.Value is not null)
                    orient = ps.Orientation.Value.ToString().Equals("landscape", StringComparison.OrdinalIgnoreCase)
                        ? "Landscape" : "Portrait";

                double pageW = orient == "Landscape" && paperW < paperH ? paperH : paperW;
                double pageH = orient == "Landscape" && paperW < paperH ? paperW : paperH;

                PageSetup = new PageSetupData(
                    LeftMargin: pm.Left?.Value is not null ? (double)pm.Left.Value : 0.75,
                    RightMargin: pm.Right?.Value is not null ? (double)pm.Right.Value : 0.75,
                    TopMargin: pm.Top?.Value is not null ? (double)pm.Top.Value : 0.75,
                    BottomMargin: pm.Bottom?.Value is not null ? (double)pm.Bottom.Value : 0.75,
                    HeaderMargin: pm.Header?.Value is not null ? (double)pm.Header.Value : 0.5,
                    FooterMargin: pm.Footer?.Value is not null ? (double)pm.Footer.Value : 0.5,
                    CenterHorizontally: po?.HorizontalCentered?.Value ?? false,
                    CenterVertically: po?.VerticalCentered?.Value ?? false,
                    PaperSize: (int)(ps.PaperSize?.Value ?? 1),
                    Orientation: orient,
                    PageWidth: pageW,
                    PageHeight: pageH,
                    FitToPagesWide: (double)(ps.FitToWidth?.Value ?? 1),
                    FitToPagesTall: (double)(ps.FitToHeight?.Value ?? 1)
                );
            }

            // Column Widths
            ReadColumnWidths(ws);

            // Row Heights
            ReadRowHeights(ws);

            // Merged Cells
            var merges = ws.Descendants<MergeCells>().FirstOrDefault();
            if (merges is not null)
                foreach (var mc in merges.Cast<MergeCell>())
                    MergeCells.Add(mc.Reference?.Value ?? "?");

            // Comments
            var commentsPart = wsPart?.WorksheetCommentsPart;
            if (commentsPart is not null)
            {
                var xml = XElement.Load(commentsPart.GetStream());
                var ns = xml.GetDefaultNamespace();
                foreach (var comment in xml.Descendants(ns + "comment"))
                {
                    string? cellRef = comment.Attribute("ref")?.Value;
                    string? text = comment.Descendants(ns + "t").Select(t => t.Value).FirstOrDefault();
                    if (cellRef is not null)
                        Comments.Add((cellRef, text ?? ""));
                }
            }

            // Images
            var drawingsPart = wsPart?.DrawingsPart;
            if (drawingsPart is not null)
            {
                foreach (var anchor in drawingsPart.WorksheetDrawing.Descendants()
                    .Where(e => e.LocalName == "twoCellAnchor" || e.LocalName == "oneCellAnchor" || e.LocalName == "absoluteAnchor"))
                {
                    var fromMarker = anchor.GetFirstChild<DDS.FromMarker>();
                    if (fromMarker is not null)
                    {
                        var colEl = fromMarker.ChildElements.FirstOrDefault(e => e.LocalName == "col");
                        var rowEl = fromMarker.ChildElements.FirstOrDefault(e => e.LocalName == "row");
                        ImageAnchors.Add($"({colEl?.InnerText ?? "?"},{rowEl?.InnerText ?? "?"})");
                    }
                }
            }

            // Only read the first data sheet
            break;
        }
    }

    string? ReadPrintArea(WorkbookPart wbPart, string sheetName)
    {
        var defNames = wbPart.Workbook.DefinedNames;
        if (defNames is null) return null;

        foreach (var dn in defNames.Cast<DefinedName>())
        {
            string? name = dn.Name?.Value;
            if (name is null || !name.Equals("_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                continue;

            string? text = dn.Text;
            if (string.IsNullOrEmpty(text)) continue;

            text = text.Trim('\'', '=');
            int excl = text.IndexOf('!');
            if (excl < 0) continue;

            string refSheet = text[..excl].Trim('\'');
            string range = text[(excl + 1)..];

            if (!string.Equals(refSheet, sheetName, StringComparison.OrdinalIgnoreCase))
                continue;

            return range;
        }
        return null;
    }

    static (int StartCol, int StartRow, int EndCol, int EndRow)? ParsePrintArea(string range)
    {
        var parts = range.Replace("$", "").Split(':');
        if (parts.Length == 0) return null;
        var start = ParseCellRef(parts[0]);
        if (parts.Length == 1) return (start.col, start.row, start.col, start.row);
        var end = ParseCellRef(parts[1]);
        return (start.col, start.row, end.col, end.row);
    }

    void DebugMargins(PageMargins? pm)
    {
        if (pm is null) return;
        Console.WriteLine($"    [DEBUG] Raw pm.Left={pm.Left?.Value} pm.Top={pm.Top?.Value} pm.Header={pm.Header?.Value}");
    }

    void ReadColumnWidths(Worksheet ws)
    {
        var sheetProps = ws.Descendants<SheetFormatProperties>().FirstOrDefault();
        double defaultCharWidth = (double)(sheetProps?.DefaultColumnWidth?.Value ?? 8.43);
        int maxCol = ParseMaxColumn(ws.SheetDimension?.Reference?.Value);
        if (maxCol == 0) return;

        var widths = new double[maxCol];
        double defaultWidthPoints = ColWidthToPoints(defaultCharWidth);
        for (int i = 0; i < maxCol; i++) widths[i] = defaultWidthPoints;

        var cols = ws.Descendants<Columns>().FirstOrDefault();
        if (cols is not null)
        {
            foreach (var col in cols.Cast<Column>())
            {
                uint min = col.Min?.Value ?? 1;
                uint max = col.Max?.Value ?? min;
                double? cw = col.Width?.Value;
                if (cw is null) continue;

                double w = ColWidthToPoints((double)cw);
                for (uint c = min; c <= max && c <= maxCol; c++)
                    widths[c - 1] = w;
            }
        }

        ColumnWidths.AddRange(widths);
    }

    void ReadRowHeights(Worksheet ws)
    {
        var sheetProps = ws.Descendants<SheetFormatProperties>().FirstOrDefault();
        double defaultHeight = (double)(sheetProps?.DefaultRowHeight?.Value ?? 15);

        var sd = ws.Descendants<SheetData>().FirstOrDefault();
        if (sd is null) return;

        foreach (var row in sd.Cast<Row>())
        {
            int rowIdx = (int)(row.RowIndex?.Value ?? 0);
            double height = row.Height?.Value is not null ? (double)row.Height.Value : defaultHeight;
            RowHeights.Add((rowIdx, height));
        }
    }

    static double ColWidthToPoints(double charWidth)
    {
        double maxDigitWidth = 7;
        double pixels = charWidth * maxDigitWidth + 5;
        return pixels * 72.0 / 96.0;
    }

    static int ParseMaxColumn(string? dimRef)
    {
        if (string.IsNullOrEmpty(dimRef)) return 0;
        var parts = dimRef.Split(':');
        if (parts.Length < 2) return 0;
        string endRef = parts[^1];
        int col = 0;
        string letters = new(endRef.TakeWhile(char.IsLetter).ToArray());
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        return col;
    }

    static (int col, int row) ParseCellRef(string ref_)
    {
        string letters = new(ref_.TakeWhile(char.IsLetter).ToArray());
        string digits = new(ref_.SkipWhile(char.IsLetter).ToArray());
        int col = 0;
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        int row = int.TryParse(digits, out var r) ? r : 1;
        return (col, row);
    }

    static (double w, double h) GetPaperDimensions(int paperSize) => paperSize switch
    {
        1 => (612, 792), 2 => (612, 792), 3 => (612, 1008),
        4 => (792, 1224), 5 => (612, 1008), 6 => (612, 792),
        7 => (612, 792), 8 => (595, 842), 9 => (842, 1191),
        10 => (595, 842), 11 => (420, 595), 12 => (595, 842),
        13 => (498, 708), 14 => (612, 792), 15 => (612, 1008),
        _ => (612, 792)
    };

    public void PrintAll()
    {
        Console.WriteLine($"  Sheets: {string.Join(", ", SheetNames)}");
        Console.WriteLine($"  PrintArea: {PrintArea ?? "(none)"}  parsed: {PrintAreaParsed}");
        if (PageSetup is not null)
        {
            var ps = PageSetup.Value;
            Console.WriteLine($"  PageSetup:");
            Console.WriteLine($"    PaperSize={ps.PaperSize} Orientation={ps.Orientation}");
            Console.WriteLine($"    Page={ps.PageWidth:F0}x{ps.PageHeight:F0}pt");
            Console.WriteLine($"    Margins(in): L={ps.LeftMargin:F4} R={ps.RightMargin:F4} T={ps.TopMargin:F4} B={ps.BottomMargin:F4} Hdr={ps.HeaderMargin:F4} Ftr={ps.FooterMargin:F4}");
            Console.WriteLine($"    Center: H={ps.CenterHorizontally} V={ps.CenterVertically}");
            Console.WriteLine($"    FitTo: W={ps.FitToPagesWide} H={ps.FitToPagesTall}");
        }
        else
        {
            Console.WriteLine($"  PageSetup: (not found in OpenXML)");
        }

        if (ColumnWidths.Count > 0)
        {
            Console.Write($"  ColumnWidths ({ColumnWidths.Count} cols):");
            for (int i = 0; i < ColumnWidths.Count; i++)
                Console.Write($" {(char)('A' + i)}={ColumnWidths[i]:F2}");
            Console.WriteLine();
        }

        Console.WriteLine($"  RowHeights ({RowHeights.Count} rows)");
        foreach (var g in RowHeights.GroupBy(r => r.Height).OrderBy(g => g.Key))
            Console.WriteLine($"    {g.Count()} rows @ {g.Key:F1}pt");
        Console.WriteLine($"  MergeCells ({MergeCells.Count}): {string.Join(", ", MergeCells)}");
        Console.WriteLine($"  Comments ({Comments.Count}): {string.Join(" | ", Comments.Select(c => $"{c.Address}='{c.Text}'"))}");
        Console.WriteLine($"  Images ({ImageAnchors.Count}): anchors {string.Join(", ", ImageAnchors)}");
    }
}
