using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class ArtifactGenerator
{
    private readonly string _outputDir;
    private readonly IProgress<string> _progress;

    public ArtifactGenerator(string outputDir, IProgress<string> progress)
    {
        _outputDir = outputDir;
        _progress = progress;
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(Path.Combine(outputDir, "debug"));
    }

    public async Task GenerateAllAsync(ExtractionResult extraction, DatabaseDump? database)
    {
        _progress.Report("[10/15] Generating ConMas-compatible artifacts...");

        var jsonOpts = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        // workbook.json
        await WriteJsonAsync("workbook.json", extraction.Workbook, jsonOpts);

        // styles.json
        await WriteJsonAsync("styles.json", extraction.Styles, jsonOpts);

        // Generate per-sheet artifacts
        foreach (var sheet in extraction.Workbook.Sheets)
        {
            var safeName = SanitizeFileName(sheet.Name);

            // page.json (page setup)
            var pageData = new
            {
                sheet.Name,
                sheet.SheetId,
                sheet.Visible,
                sheet.PageSetup,
                sheet.PageMargins,
                sheet.SheetFormatProperties,
                sheet.PrintArea,
                sheet.FreezePanes,
                sheet.PrintTitles,
                sheet.AutoFilter,
                sheet.Protection
            };
            await WriteJsonAsync($"page_{safeName}.json", pageData, jsonOpts);

            // cells.json
            await WriteJsonAsync($"cells_{safeName}.json", sheet.Cells, jsonOpts);

            // merged.json
            await WriteJsonAsync($"merged_{safeName}.json", sheet.MergedCells, jsonOpts);

            // sheet.json (comprehensive)
            await WriteJsonAsync($"sheet_{safeName}.json", sheet, jsonOpts);

            // comments.json
            await WriteJsonAsync($"comments_{safeName}.json", sheet.Comments, jsonOpts);

            // validation.json
            await WriteJsonAsync($"validation_{safeName}.json", sheet.DataValidations, jsonOpts);

            // objects.json (shapes + images + objects)
            var objectsData = new
            {
                images = sheet.Images,
                shapes = sheet.Shapes,
                objects = sheet.Objects
            };
            await WriteJsonAsync($"objects_{safeName}.json", objectsData, jsonOpts);

            // protection.json
            await WriteJsonAsync($"protection_{safeName}.json", sheet.Protection, jsonOpts);

            // named_ranges.json
            await WriteJsonAsync($"named_ranges.json", extraction.Workbook.DefinedNames, jsonOpts);

            // metadata.json
            await WriteJsonAsync("metadata.json", extraction.Workbook.Properties, jsonOpts);

            // Page breaks
            await WriteJsonAsync($"page_breaks_{safeName}.json", sheet.PageBreaks, jsonOpts);

            // Conditional formatting
            await WriteJsonAsync($"conditional_{safeName}.json", sheet.ConditionalFormats, jsonOpts);

            // Hyperlinks
            await WriteJsonAsync($"hyperlinks_{safeName}.json", sheet.Hyperlinks, jsonOpts);

            // Rows
            await WriteJsonAsync($"rows_{safeName}.json", sheet.Rows, jsonOpts);

            // Columns
            await WriteJsonAsync($"columns_{safeName}.json", sheet.Columns, jsonOpts);
        }

        // xml_data.xml - Generate XML in ConMas format
        await GenerateXmlDataAsync(extraction, database);

        // def_cluster.json
        await GenerateDefClusterAsync(extraction, database);

        // background_image.png - Extract from workbook
        await ExtractBackgroundImageAsync(extraction);

        _progress.Report("[11/15] Artifacts written to output directory.");
    }

    private async Task WriteJsonAsync(string filename, object data, JsonSerializerOptions opts)
    {
        var path = Path.Combine(_outputDir, filename);
        var json = JsonSerializer.Serialize(data, opts);
        await File.WriteAllTextAsync(path, json);
        _progress.Report($"  Wrote {filename}");
    }

    private async Task GenerateXmlDataAsync(ExtractionResult extraction, DatabaseDump? database)
    {
        _progress.Report("  Generating xml_data.xml...");

        var xmlns = XNamespace.Get("http://schemas.conmas.com/forms");
        var xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
        var conmasNs = XNamespace.Get("http://schemas.conmas.com/forms");

        var doc = new XDocument(
            new XDeclaration("1.0", "utf-8", "yes"),
            new XElement(conmasNs + "form",
                new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                new XElement("formInfo",
                    new XElement("width", 612),
                    new XElement("height", 792),
                    new XElement("marginLeft", extraction.Workbook.Sheets.FirstOrDefault()?.PageMargins.Left ?? 0.70866),
                    new XElement("marginRight", extraction.Workbook.Sheets.FirstOrDefault()?.PageMargins.Right ?? 0.70866),
                    new XElement("marginTop", extraction.Workbook.Sheets.FirstOrDefault()?.PageMargins.Top ?? 0.74803),
                    new XElement("marginBottom", extraction.Workbook.Sheets.FirstOrDefault()?.PageMargins.Bottom ?? 0.74803),
                    new XElement("centerH", true),
                    new XElement("centerV", true),
                    new XElement("isOriginalWhole", true),
                    new XElement("printArea", extraction.Workbook.Sheets.FirstOrDefault()?.PrintArea ?? "A1:D12")
                ),
                new XElement("clusters", GenerateClustersXml(extraction))
            )
        );

        var xml = doc.ToString();
        await File.WriteAllTextAsync(Path.Combine(_outputDir, "xml_data.xml"), xml);

        // Debug: write raw XML
        await File.WriteAllTextAsync(Path.Combine(_outputDir, "debug", "xml_data_debug.xml"), xml);
    }

    private XElement GenerateClustersXml(ExtractionResult extraction)
    {
        var clusters = new XElement("clusters");
        int clusterId = 0;

        foreach (var sheet in extraction.Workbook.Sheets)
        {
            foreach (var mc in sheet.MergedCells)
            {
                // Get the cell value from first cell of merged range
                var topLeftCell = sheet.Cells.FirstOrDefault(c =>
                    c.Row == mc.StartRow && c.Column == mc.StartColumn);

                clusters.Add(new XElement("cluster",
                    new XElement("cluster_id", clusterId),
                    new XElement("cell_address", mc.Reference),
                    new XElement("left", ComputeRatio(mc.StartColumn, sheet.Columns)),
                    new XElement("right", ComputeRatio(mc.EndColumn, sheet.Columns)),
                    new XElement("top", ComputeRowRatio(mc.StartRow, sheet.Rows)),
                    new XElement("bottom", ComputeRowRatio(mc.EndRow, sheet.Rows)),
                    new XElement("left_pt", ComputePointLeft(mc.StartColumn, sheet)),
                    new XElement("top_pt", ComputePointTop(mc.StartRow, sheet)),
                    new XElement("width_pt", ComputePointWidth(mc.StartColumn, mc.EndColumn, sheet)),
                    new XElement("height_pt", ComputePointHeight(mc.StartRow, mc.EndRow, sheet)),
                    new XElement("value", topLeftCell?.Value?.ToString() ?? ""),
                    new XElement("style_index", topLeftCell?.StyleIndex ?? 0)
                ));
                clusterId++;
            }

            // Also include non-merged cells that have values
            var nonMergedCells = sheet.Cells
                .Where(c => !sheet.MergedCells.Any(m =>
                    c.Row >= m.StartRow && c.Row <= m.EndRow &&
                    c.Column >= m.StartColumn && c.Column <= m.EndColumn))
                .GroupBy(c => new { c.Row, c.Column })
                .Select(g => g.First());

            foreach (var cell in nonMergedCells)
            {
                if (cell.Value == null) continue;

                clusters.Add(new XElement("cluster",
                    new XElement("cluster_id", clusterId),
                    new XElement("cell_address", cell.Reference),
                    new XElement("left", ComputeRatio(cell.Column, sheet.Columns)),
                    new XElement("right", ComputeRatio(cell.Column + 1, sheet.Columns)),
                    new XElement("top", ComputeRowRatio(cell.Row, sheet.Rows)),
                    new XElement("bottom", ComputeRowRatio(cell.Row + 1, sheet.Rows)),
                    new XElement("value", cell.Value?.ToString() ?? ""),
                    new XElement("style_index", cell.StyleIndex ?? 0)
                ));
                clusterId++;
            }
        }

        return clusters;
    }

    private double ComputeRatio(int column, List<ColumnData> columns)
    {
        double totalWidth = 0;
        for (int i = 1; i <= column; i++)
        {
            var col = columns.FirstOrDefault(c => i >= c.Min && i <= c.Max);
            totalWidth += col?.Width ?? 8.43; // default column width
        }
        // Page width in points = 612 (Letter)
        // Convert character width to points: 1 char = 7 pixels @ 96dpi, 1pt = 1.333px
        double pageWidth = 612;
        double points = totalWidth * 7 / 1.333; // rough conversion
        return Math.Round(points / pageWidth, 7);
    }

    private double ComputeRowRatio(int rowIndex, List<RowData> rows)
    {
        var row = rows.FirstOrDefault(r => r.RowIndex == rowIndex);
        double rowHeight = row?.Height ?? 14.4; // default row height
        double pageHeight = 792; // Letter portrait
        double cumulative = 0;
        foreach (var r in rows.Where(r => r.RowIndex < rowIndex).OrderBy(r => r.RowIndex))
        {
            cumulative += r.Height ?? 14.4;
        }
        return Math.Round((cumulative + rowHeight / 2) / pageHeight, 7);
    }

    private double ComputePointLeft(int column, SheetData sheet)
    {
        double totalWidth = 0;
        for (int i = 1; i < column; i++)
        {
            var col = sheet.Columns.FirstOrDefault(c => i >= c.Min && i <= c.Max);
            totalWidth += col?.Width ?? 8.43;
        }
        return totalWidth * 7 / 1.333;
    }

    private double ComputePointTop(int row, SheetData sheet)
    {
        double total = 0;
        foreach (var r in sheet.Rows.Where(r => r.RowIndex < row).OrderBy(r => r.RowIndex))
        {
            total += r.Height ?? 14.4;
        }
        return total;
    }

    private double ComputePointWidth(int startCol, int endCol, SheetData sheet)
    {
        double total = 0;
        for (int i = startCol; i <= endCol; i++)
        {
            var col = sheet.Columns.FirstOrDefault(c => i >= c.Min && i <= c.Max);
            total += col?.Width ?? 8.43;
        }
        return total * 7 / 1.333;
    }

    private double ComputePointHeight(int startRow, int endRow, SheetData sheet)
    {
        double total = 0;
        for (int i = startRow; i <= endRow; i++)
        {
            var row = sheet.Rows.FirstOrDefault(r => r.RowIndex == i);
            total += row?.Height ?? 14.4;
        }
        return total;
    }

    private async Task GenerateDefClusterAsync(ExtractionResult extraction, DatabaseDump? database)
    {
        _progress.Report("  Generating def_cluster.json...");
        var clusters = new List<Dictionary<string, object?>>();
        int clusterId = 0;

        foreach (var sheet in extraction.Workbook.Sheets)
        {
            foreach (var mc in sheet.MergedCells)
            {
                var left = ComputeRatio(mc.StartColumn, sheet.Columns);
                var right = ComputeRatio(mc.EndColumn, sheet.Columns);
                var top = ComputeRowRatio(mc.StartRow, sheet.Rows);
                var bottom = ComputeRowRatio(mc.EndRow, sheet.Rows);

                clusters.Add(new Dictionary<string, object?>
                {
                    ["cluster_id"] = clusterId,
                    ["cell_addr"] = mc.Reference,
                    ["left_position"] = left.ToString("F7"),
                    ["right_position"] = right.ToString("F7"),
                    ["top_position"] = top.ToString("F7"),
                    ["bottom_position"] = bottom.ToString("F7"),
                    ["left_pt"] = Math.Round(ComputePointLeft(mc.StartColumn, sheet), 2),
                    ["top_pt"] = Math.Round(ComputePointTop(mc.StartRow, sheet), 2),
                    ["width_pt"] = Math.Round(ComputePointWidth(mc.StartColumn, mc.EndColumn, sheet), 2),
                    ["height_pt"] = Math.Round(ComputePointHeight(mc.StartRow, mc.EndRow, sheet), 2),
                    ["sheet_name"] = sheet.Name
                });
                clusterId++;
            }

            // Non-merged cells with values
            var nonMergedCells = sheet.Cells
                .Where(c => !sheet.MergedCells.Any(m =>
                    c.Row >= m.StartRow && c.Row <= m.EndRow &&
                    c.Column >= m.StartColumn && c.Column <= m.EndColumn))
                .GroupBy(c => new { c.Row, c.Column })
                .Select(g => g.First());

            foreach (var cell in nonMergedCells)
            {
                if (cell.Value == null) continue;

                var left = ComputeRatio(cell.Column, sheet.Columns);
                var right = ComputeRatio(cell.Column + 1, sheet.Columns);
                var top = ComputeRowRatio(cell.Row, sheet.Rows);
                var bottom = ComputeRowRatio(cell.Row + 1, sheet.Rows);

                clusters.Add(new Dictionary<string, object?>
                {
                    ["cluster_id"] = clusterId,
                    ["cell_addr"] = cell.Reference,
                    ["left_position"] = left.ToString("F7"),
                    ["right_position"] = right.ToString("F7"),
                    ["top_position"] = top.ToString("F7"),
                    ["bottom_position"] = bottom.ToString("F7"),
                    ["value"] = cell.Value?.ToString(),
                    ["sheet_name"] = sheet.Name
                });
                clusterId++;
            }
        }

        var jsonOpts = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };
        await WriteJsonAsync("def_cluster.json", new { clusters, cluster_count = clusters.Count }, jsonOpts);
    }

    private async Task ExtractBackgroundImageAsync(ExtractionResult extraction)
    {
        _progress.Report("  Extracting background image...");
        foreach (var sheet in extraction.Workbook.Sheets)
        {
            foreach (var img in sheet.Images)
            {
                if (img.DataBase64 != null)
                {
                    var bytes = Convert.FromBase64String(img.DataBase64);
                    var ext = img.ContentType?.Contains("png") == true ? "png"
                        : img.ContentType?.Contains("jpeg") == true ? "jpg"
                        : img.ContentType?.Contains("gif") == true ? "gif"
                        : "bin";
                    await File.WriteAllBytesAsync(Path.Combine(_outputDir, $"background_image.{ext}"), bytes);
                    _progress.Report($"  Extracted background image ({bytes.Length} bytes)");
                    return;
                }
            }
        }
        _progress.Report("  No background image found in workbook.");
    }

    private string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Join("_", name.Split(invalid, StringSplitOptions.RemoveEmptyEntries));
    }
}
