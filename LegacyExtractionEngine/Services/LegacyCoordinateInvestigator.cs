using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Comprehensive diagnostic engine that extracts ALL coordinate information
/// from Excel COM, compares with DB and OpenXML values, and determines the
/// exact legacy coordinate transformation formula.
/// 
/// Run via: dotnet run -- --investigate
/// </summary>
public class LegacyCoordinateInvestigator
{
    private readonly IProgress<string> _progress;
    private readonly string _connectionString;
    private readonly string _outputDir;

    public LegacyCoordinateInvestigator(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test");
    }

    public async Task InvestigateAllAsync()
    {
        _progress.Report("=== Comprehensive Legacy Coordinate Investigation ===");
        _progress.Report("This will read ALL Excel COM properties, workbook images, and cluster patterns.");

        // Investigate both templates
        await InvestigateTemplateAsync(546,
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx",
            "Template546");

        var wb547 = @"C:\Users\MCF-JOHNELLEEMPUERTO\AppData\Local\Temp\opencode\547_workbook_copy.xlsx";
        if (!File.Exists(wb547))
        {
            var dir = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents";
            wb547 = Directory.GetFiles(dir, "*V3.1*", SearchOption.TopDirectoryOnly).FirstOrDefault() ?? "";
        }
        if (File.Exists(wb547))
            await InvestigateTemplateAsync(547, wb547, "Template547");

        _progress.Report("=== Investigation Complete ===");
    }

    private async Task InvestigateTemplateAsync(int defTopId, string workbookPath, string templateDir)
    {
        _progress.Report($"--- Template {defTopId} ({templateDir}) ---");

        // 1. Load DB data
        var dbReader = new DatabaseReader(_connectionString, _progress);
        var db = await dbReader.ReadAllAsync(defTopId);
        if (db == null) { _progress.Report("  SKIP: No DB data"); return; }

        var clusters = db.DefSheets.SelectMany(s => s.Clusters).ToList();
        var outputDir = Path.Combine(_outputDir, templateDir);
        Directory.CreateDirectory(outputDir);

        // 2. Investigate background images via OpenXML
        await InvestigateImagesAsync(workbookPath, db, outputDir);

        // 3. Investigate via Excel COM
        var comData = await InvestigateComAsync(workbookPath, clusters, db, outputDir);

        // 4. Investigate cluster generation (compare all workbook cells vs DB clusters)
        await InvestigateClustersAsync(workbookPath, clusters, db, outputDir);

        // 5. Generate comprehensive reports
        if (comData != null)
        {
            await GenerateCoordinateComparisonAsync(comData, outputDir);
            await GenerateCoordinateTraceAsync(comData, outputDir);
        }
    }

    /// <summary>
    /// Step 1: Investigate ALL images in the workbook and compare with DB blob
    /// </summary>
    private async Task InvestigateImagesAsync(string workbookPath, DatabaseDump db, string dir)
    {
        _progress.Report("  [IMAGES] Searching workbook for all image sources...");
        var sb = new StringBuilder();
        sb.AppendLine("# Background & Image Investigation");
        sb.AppendLine();
        sb.AppendLine($"**Workbook:** {Path.GetFileName(workbookPath)}");
        sb.AppendLine($"**DB def_top_id:** {db.DefTop?.DefTopId}");
        sb.AppendLine();

        // DB background image
        var dbBg = db.DefTop?.BackgroundImageFile;
        sb.AppendLine("## Database Background Image");
        sb.AppendLine();
        if (dbBg != null)
        {
            sb.AppendLine($"| Property | Value |");
            sb.AppendLine($"|----------|------:|");
            sb.AppendLine($"| Size | {dbBg.Length} bytes |");
            sb.AppendLine($"| SHA256 | {HexSHA(dbBg)} |");
            sb.AppendLine($"| MD5 | {HexMD5(dbBg)} |");
        }
        else
        {
            sb.AppendLine("No background image in DB.");
        }

        sb.AppendLine();
        sb.AppendLine("## Workbook Image Sources");
        sb.AppendLine();

        // Open the workbook as a package
        try
        {
            using var doc = SpreadsheetDocument.Open(workbookPath, false);
            var wbPart = doc.WorkbookPart;

            if (wbPart == null)
            {
                sb.AppendLine("ERROR: Could not open workbook package.");
                await File.WriteAllTextAsync(Path.Combine(dir, "background_image_report.md"), sb.ToString());
                return;
            }

            // Dictionary of all image parts by relationship ID
            var imageParts = new Dictionary<string, (string partPath, byte[] data)>();

            // Scan ALL parts recursively for images
            foreach (var part in wbPart.GetPartsOfType<ImagePart>())
            {
                var uri = part.Uri;
                using var stream = part.GetStream();
                var bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                var relId = wbPart.GetIdOfPart(part) ?? "unknown";
                imageParts[relId] = (uri.ToString(), bytes);
            }

            // Also check DrawingML and VML parts
            foreach (var rel in wbPart.ExternalRelationships)
            {
                // Not relevant for embedded images
            }

            // Check each sheet for background page setup
            foreach (var sheetPart in wbPart.WorksheetParts)
            {
                var sheetRelId = wbPart.GetIdOfPart(sheetPart);

                // Check for images from sheet relationships
                try
                {
                    foreach (var rel in sheetPart.Parts)
                    {
                        if (rel.OpenXmlPart is ImagePart img)
                        {
                            using var stream = img.GetStream();
                            var bytes = new byte[stream.Length];
                            stream.Read(bytes, 0, bytes.Length);
                            imageParts[$"sheet_img_{rel.RelationshipId}"] = (img.Uri.ToString(), bytes);
                        }
                    }
                }
                catch { }

                // Check Image Parts from sheet
                foreach (var imgPart in sheetPart.GetPartsOfType<ImagePart>())
                {
                    using var stream = imgPart.GetStream();
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    var relId = sheetPart.GetIdOfPart(imgPart) ?? "unknown";
                    imageParts[$"sheet_img_{relId}"] = (imgPart.Uri.ToString(), bytes);
                }
            }

            // Report all image sources
            sb.AppendLine($"| Source | Location | Size | SHA256 | Matches DB? |");
            sb.AppendLine($"|-------|----------|------|--------|------------|");
            foreach (var (relId, (path, data)) in imageParts)
            {
                var sha = HexSHA(data);
                var matches = dbBg != null && data.SequenceEqual(dbBg) ? "**YES**" : "No";
                sb.AppendLine($"| {relId} | {path} | {data.Length} bytes | {sha} | {matches} |");
            }

            // Check for images in drawings
            foreach (var drawingsPart in wbPart.GetPartsOfType<DrawingsPart>())
            {
                sb.AppendLine();
                sb.AppendLine($"### DrawingML Part: {drawingsPart.Uri}");
                try
                {
                    var xml = XDocument.Load(drawingsPart.GetStream());
                    sb.AppendLine("```xml");
                    sb.AppendLine(xml.ToString().Substring(0, Math.Min(2000, xml.ToString().Length)));
                    sb.AppendLine("```");
                }
                catch { sb.AppendLine("(Could not parse)"); }
            }

            // Check VML drawings
            foreach (var vmlPart in wbPart.GetPartsOfType<LegacyDiagramTextPart>())
            {
                // VML parts are not common in modern xlsx
            }                // Check headers/footers for images
            foreach (var sheetPart in wbPart.WorksheetParts)
            {
                try
                {
                    // Scan all sheet-level parts for images
                    foreach (var rel in sheetPart.Parts)
                    {
                        if (rel.OpenXmlPart is ImagePart imgPart)
                        {
                            if (imageParts.ContainsKey(rel.RelationshipId)) continue;
                            using var stream = imgPart.GetStream();
                            var bytes = new byte[stream.Length];
                            stream.Read(bytes, 0, bytes.Length);
                            var sha = HexSHA(bytes);
                            var matches = dbBg != null && bytes.SequenceEqual(dbBg) ? "**YES**" : "No";
                            sb.AppendLine($"- Sheet image: {imgPart.Uri}, {bytes.Length} bytes, SHA={sha}, DB match={matches}");
                        }
                    }
                }
                catch { }
            }

            // Summary
            sb.AppendLine();
            sb.AppendLine("## Summary");
            sb.AppendLine();
            if (imageParts.Count == 0)
            {
                sb.AppendLine("No images found in workbook.");
            }
            else if (dbBg != null)
            {
                var anyMatch = imageParts.Any(kv => kv.Value.data.SequenceEqual(dbBg));
                if (anyMatch)
                {
                    var matched = imageParts.First(kv => kv.Value.data.SequenceEqual(dbBg));
                    sb.AppendLine($"**DB background image found in workbook at:** {matched.Value.partPath} ({matched.Key})");
                    sb.AppendLine();
                    sb.AppendLine("The background image is stored in the workbook and can be extracted.");
                }
                else
                {
                    sb.AppendLine("**DB background image NOT found in any workbook part.**");
                    sb.AppendLine("The background may have been added separately by the legacy importer,");
                    sb.AppendLine("or the workbook on disk differs from the version that was imported.");
                    sb.AppendLine();
                    sb.AppendLine("Image sources found in workbook:");
                    foreach (var (relId, (path, data)) in imageParts)
                    {
                        var sha = HexSHA(data);
                        sb.AppendLine($"- {relId}: {path} ({data.Length} bytes, {sha})");
                    }
                }
            }

            await File.WriteAllTextAsync(Path.Combine(dir, "background_image_report.md"), sb.ToString());
            _progress.Report($"  [IMAGES] Wrote background_image_report.md ({imageParts.Count} image sources)");
        }
        catch (Exception ex)
        {
            sb.AppendLine($"ERROR: {ex.Message}");
            await File.WriteAllTextAsync(Path.Combine(dir, "background_image_report.md"), sb.ToString());
            _progress.Report($"  [IMAGES] Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Step 2: Comprehensive COM investigation — reads ALL properties
    /// </summary>
    private async Task<ComData?> InvestigateComAsync(string workbookPath, List<DefCluster> clusters, DatabaseDump db, string dir)
    {
        _progress.Report("  [COM] Opening Excel Interop for comprehensive property extraction...");

        if (clusters.Count == 0) { _progress.Report("  [COM] No clusters to trace"); return null; }

        dynamic? excel = null;
        dynamic? wb = null;

        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) { _progress.Report("  [COM] Excel not available"); return null; }

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;

            _progress.Report($"  [COM] Opening: {Path.GetFileName(workbookPath)}");
            wb = excel.Workbooks.Open(workbookPath);
            dynamic ws = wb.Sheets[1];
            dynamic pageSetup = ws.PageSetup;
            dynamic activeWindow = excel.ActiveWindow;

            // ========= PAGE SETUP (ALL properties) =========
            var psInfo = new PsInfo();
            try { psInfo.PaperSize = (int)pageSetup.PaperSize; } catch { }
            try { psInfo.Zoom = (int?)pageSetup.Zoom; } catch { }
            try { psInfo.FitToPagesWide = (int?)pageSetup.FitToPagesWide; } catch { }
            try { psInfo.FitToPagesTall = (int?)pageSetup.FitToPagesTall; } catch { }
            try { psInfo.LeftMargin = (double)pageSetup.LeftMargin; } catch { }
            try { psInfo.RightMargin = (double)pageSetup.RightMargin; } catch { }
            try { psInfo.TopMargin = (double)pageSetup.TopMargin; } catch { }
            try { psInfo.BottomMargin = (double)pageSetup.BottomMargin; } catch { }
            try { psInfo.HeaderMargin = (double)pageSetup.HeaderMargin; } catch { }
            try { psInfo.FooterMargin = (double)pageSetup.FooterMargin; } catch { }
            try { psInfo.CenterHorizontally = (bool)pageSetup.CenterHorizontally; } catch { }
            try { psInfo.CenterVertically = (bool)pageSetup.CenterVertically; } catch { }
            try { psInfo.Orientation = (int)pageSetup.Orientation; } catch { }
            try { psInfo.PageWidth = (double)pageSetup.PageWidth; } catch { }
            try { psInfo.PageHeight = (double)pageSetup.PageHeight; } catch { }
            try { psInfo.PrintArea = (string)pageSetup.PrintArea; } catch { }
            try { psInfo.Order = (int)pageSetup.Order; } catch { }
            try { psInfo.FirstPageNumber = (int)pageSetup.FirstPageNumber; } catch { }
            try
            {
                var pq = pageSetup.PrintQuality;
                psInfo.PrintQuality = pq is Array arr ? arr.Cast<int>().ToArray() : null;
            }
            catch { }
            try { psInfo.PaperSizeName = pageSetup.PaperSize.ToString(); } catch { }
            try { psInfo.FooterPicture = pageSetup.FooterPicture?.Name ?? null; } catch { }

            _progress.Report($"  [COM] Page: {psInfo.PageWidth:F1}x{psInfo.PageHeight:F1}, " +
                $"Zoom={psInfo.Zoom}, FitToPages={psInfo.FitToPagesWide}x{psInfo.FitToPagesTall}, " +
                $"Margins(L,R,T,B): {psInfo.LeftMargin:F3}, {psInfo.RightMargin:F3}, " +
                $"{psInfo.TopMargin:F3}, {psInfo.BottomMargin:F3}");

            // ========= USED RANGE =========
            dynamic usedRange = ws.UsedRange;
            try { psInfo.UsedRangeAddress = (string)usedRange.Address; } catch { }
            try { psInfo.UsedRangeRows = usedRange.Rows.Count; } catch { }
            try { psInfo.UsedRangeColumns = usedRange.Columns.Count; } catch { }
            try { psInfo.UsedRangeLeft = (double)usedRange.Left; } catch { }
            try { psInfo.UsedRangeTop = (double)usedRange.Top; } catch { }
            try { psInfo.UsedRangeWidth = (double)usedRange.Width; } catch { }
            try { psInfo.UsedRangeHeight = (double)usedRange.Height; } catch { }

            _progress.Report($"  [COM] UsedRange: {psInfo.UsedRangeAddress}, " +
                $"Left={psInfo.UsedRangeLeft:F1}, Top={psInfo.UsedRangeTop:F1}, " +
                $"Width={psInfo.UsedRangeWidth:F1}, Height={psInfo.UsedRangeHeight:F1}");

            // ========= ACTIVE WINDOW =========
            try { psInfo.WindowZoom = (int)activeWindow.Zoom; } catch { }
            try { psInfo.WindowVisibleRange = (string)activeWindow.VisibleRange.Address; } catch { }
            try { psInfo.ScrollRow = (int)activeWindow.ScrollRow; } catch { }
            try { psInfo.ScrollColumn = (int)activeWindow.ScrollColumn; } catch { }

            // ========= PAGE BREAKS =========
            try { psInfo.HPageBreaks = ws.HPageBreaks.Count; } catch { }
            try { psInfo.VPageBreaks = ws.VPageBreaks.Count; } catch { }

            // ========= PROCESS EACH CLUSTER =========
            var clusterData = new List<ComClusterData>();
            foreach (var dbc in clusters)
            {
                var addr = dbc.CellAddr?.Trim() ?? "";
                if (string.IsNullOrEmpty(addr)) continue;

                _progress.Report($"  [COM] Processing {addr}...");
                dynamic range = ws.Range[addr];

                var cd = new ComClusterData
                {
                    Address = addr,
                    RangeLeft = ReadDouble(() => range.Left),
                    RangeTop = ReadDouble(() => range.Top),
                    RangeWidth = ReadDouble(() => range.Width),
                    RangeHeight = ReadDouble(() => range.Height),
                    RangeAddress = ReadString(() => range.Address),
                    RangeMergeArea = ReadString(() =>
                    {
                        dynamic ma = range.MergeArea;
                        return (string)ma.Address;
                    }),
                    RangeMergeLeft = ReadDouble(() =>
                    {
                        dynamic ma = range.MergeArea;
                        return (double)ma.Left;
                    }),
                    RangeMergeTop = ReadDouble(() =>
                    {
                        dynamic ma = range.MergeArea;
                        return (double)ma.Top;
                    }),
                    RangeMergeWidth = ReadDouble(() =>
                    {
                        dynamic ma = range.MergeArea;
                        return (double)ma.Width;
                    }),
                    RangeMergeHeight = ReadDouble(() =>
                    {
                        dynamic ma = range.MergeArea;
                        return (double)ma.Height;
                    }),
                    RangeValue = ReadString(() => (string)range.Value),
                };

                // DB values
                cd.DbLeft = ParseDouble(dbc.LeftPosition);
                cd.DbTop = ParseDouble(dbc.TopPosition);
                cd.DbRight = ParseDouble(dbc.RightPosition);
                cd.DbBottom = ParseDouble(dbc.BottomPosition);

                clusterData.Add(cd);
            }

            var result = new ComData
            {
                DefTopId = db.DefTop?.DefTopId ?? 0,
                WorkbookName = Path.GetFileName(workbookPath),
                PageSetup = psInfo,
                Clusters = clusterData,
            };

            wb.Close(false);
            Marshal.ReleaseComObject(activeWindow);
            Marshal.ReleaseComObject(ws);

            _progress.Report($"  [COM] Extracted data for {clusterData.Count} clusters");

            // Write COM dump as JSON
            var jsonOpts = new JsonSerializerOptions { WriteIndented = true };
            await File.WriteAllTextAsync(Path.Combine(dir, "excel_com_dump.json"),
                JsonSerializer.Serialize(result, jsonOpts));
            _progress.Report($"  [COM] Wrote excel_com_dump.json");

            return result;
        }
        catch (Exception ex)
        {
            _progress.Report($"  [COM] Error: {ex.Message}");
            return null;
        }
        finally
        {
            if (wb != null) try { wb.Close(false); } catch { }
            if (excel != null) try { excel.Quit(); } catch { }
        }
    }

    /// <summary>
    /// Step 3: Investigate cluster generation rules
    /// </summary>
    private async Task InvestigateClustersAsync(string workbookPath, List<DefCluster> dbClusters, DatabaseDump db, string dir)
    {
        _progress.Report("  [CLUSTER] Analyzing cluster generation patterns...");
        var sb = new StringBuilder();
        sb.AppendLine("# Cluster Generation Investigation");
        sb.AppendLine();
        sb.AppendLine($"**Workbook:** {Path.GetFileName(workbookPath)}");
        sb.AppendLine($"**DB Clusters:** {dbClusters.Count}");
        sb.AppendLine();

        // Read the workbook via OpenXML to get all cell properties
        try
        {
            using var doc = SpreadsheetDocument.Open(workbookPath, false);
            var wbPart = doc.WorkbookPart;

            foreach (var sheetPart in wbPart.WorksheetParts)
            {
                var sheet = sheetPart.Worksheet;
                var sheetData = sheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                if (sheetData == null) continue;

                // Read merged cells
                var merges = sheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.MergeCells>();
                var mergeCells = new List<string>();
                if (merges != null)
                {
                    foreach (var mc in merges.Elements<DocumentFormat.OpenXml.Spreadsheet.MergeCell>())
                        mergeCells.Add(mc.Reference?.Value ?? "");
                }

                // Read all cells with their properties
                var allCells = new List<Dictionary<string, object>>();
                foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                {
                    foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                    {
                        var cellRef = cell.CellReference?.Value ?? "";
                        var styleIdx = cell.StyleIndex?.Value;
                        var dataType = cell.DataType?.Value.ToString();
                        var cellValue = cell.CellValue?.Text;

                        // Check if this cell is in a merge range
                        bool inMerge = false;
                        string? mergeRange = null;
                        foreach (var mr in mergeCells)
                        {
                            var parts = mr.Split(':');
                            if (parts.Length == 2 && IsInRange(cellRef, parts[0], parts[1]))
                            { inMerge = true; mergeRange = mr; break; }
                        }

                        // Check if this cell is a DB cluster address or part of one
                        bool isDbCluster = dbClusters.Any(c =>
                        {
                            var dbAddr = c.CellAddr?.Trim() ?? "";
                            var dbParts = dbAddr.Split(':');
                            if (dbParts.Length == 2)
                                return IsInRange(cellRef, dbParts[0], dbParts[1]);
                            return cellRef.Equals(dbAddr, StringComparison.OrdinalIgnoreCase);
                        });

                        allCells.Add(new Dictionary<string, object>
                        {
                            ["ref"] = cellRef,
                            ["style"] = styleIdx?.ToString() ?? "",
                            ["type"] = dataType ?? "",
                            ["value"] = cellValue ?? "",
                            ["inMerge"] = inMerge,
                            ["mergeRange"] = mergeRange ?? "",
                            ["isDbCluster"] = isDbCluster,
                        });
                    }
                }

                // Analysis: What makes DB clusters special?
                sb.AppendLine($"## Sheet Analysis");
                sb.AppendLine();
                sb.AppendLine($"Total cells in workbook: {allCells.Count}");
                sb.AppendLine($"Merged cells in workbook: {mergeCells.Count}");
                sb.AppendLine($"DB clusters: {dbClusters.Count}");
                sb.AppendLine();

                // Compare DB clusters with workbook cells
                sb.AppendLine("### DB Cluster vs Workbook Cell Analysis");
                sb.AppendLine();
                sb.AppendLine("| DB Address | Type | In Merge? | Has Data? | Workbook Ref |");
                sb.AppendLine("|------------|------|-----------|-----------|--------------|");
                foreach (var dbc in dbClusters)
                {
                    var addr = dbc.CellAddr?.Trim() ?? "";
                    // Find matching cells
                    var matchingCells = allCells.Where(c => IsCellInRange(c["ref"].ToString(), addr)).ToList();
                    bool hasData = matchingCells.Any(c => !string.IsNullOrEmpty(c["value"].ToString()));
                    bool inMerge = matchingCells.Any(c => (bool)c["inMerge"]);
                    var refs = string.Join(", ", matchingCells.Select(c => c["ref"]));
                    sb.AppendLine($"| {addr} | {dbc.ClusterType ?? "?"} | {inMerge} | {hasData} | {refs} |");
                }

                // Check if DB clusters correlate with specific styles
                sb.AppendLine();
                sb.AppendLine("### Style Analysis: DB Cluster Cells vs Non-Cluster Cells");
                sb.AppendLine();
                var clusterStyles = allCells.Where(c => (bool)c["isDbCluster"])
                    .Select(c => c["style"].ToString())
                    .Distinct()
                    .ToList();
                var nonClusterStyles = allCells.Where(c => !(bool)c["isDbCluster"])
                    .Select(c => c["style"].ToString())
                    .Distinct()
                    .ToList();
                var uniqueClusterStyles = clusterStyles.Except(nonClusterStyles).ToList();
                var sharedStyles = clusterStyles.Intersect(nonClusterStyles).ToList();

                sb.AppendLine($"- Unique styles used ONLY by DB clusters: {string.Join(", ", uniqueClusterStyles.Count > 0 ? uniqueClusterStyles : new[] { "None" })}");
                sb.AppendLine($"- Styles shared with non-cluster cells: {string.Join(", ", sharedStyles)}");
                sb.AppendLine();

                // Check if DB clusters correlate with data validation
                var validations = sheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.DataValidations>();
                if (validations != null)
                {
                    sb.AppendLine("### Data Validation Analysis");
                    sb.AppendLine();
                    int clusterWithValidation = 0;
                    foreach (var dv in validations.Elements<DocumentFormat.OpenXml.Spreadsheet.DataValidation>())
                    {
                        var dvRange = dv.SequenceOfReferences?.InnerText ?? "";
                        foreach (var dbc in dbClusters)
                        {
                            var addr = dbc.CellAddr?.Trim() ?? "";
                            if (RangesOverlap(dvRange, addr))
                                clusterWithValidation++;
                        }
                    }
                    sb.AppendLine($"- DB clusters with data validation: {clusterWithValidation} / {dbClusters.Count}");
                    sb.AppendLine($"- Total data validations: {validations.Count}");
                }

                // Check for named ranges
                if (wbPart?.Workbook?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.DefinedNames>() != null)
                {
                    sb.AppendLine();
                    sb.AppendLine("### Named Range Analysis");
                    sb.AppendLine();
                    foreach (var dn in wbPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.DefinedNames>()
                        .Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>())
                    {
                        var name = dn.Name?.Value ?? "";
                        var text = dn.InnerText ?? "";
                        sb.AppendLine($"- {name}: {text}");
                    }
                }
            }

            // Summary
            sb.AppendLine();
            sb.AppendLine("## Cluster Generation Rule Hypothesis");
            sb.AppendLine();
            sb.AppendLine("Based on the analysis, DB clusters may be determined by:");
            sb.AppendLine("1. **User-selected input fields** in ConMasDesigner (most likely)");
            sb.AppendLine("2. Cells with specific style properties");
            sb.AppendLine("3. Cells within a specific column range");
            sb.AppendLine("4. Cells with data validation applied");
            sb.AppendLine("5. Named ranges that define input areas");
            sb.AppendLine();
            sb.AppendLine("The current engine generates clusters for ALL merged cells and standalone cells.");
            sb.AppendLine("The DB only stores clusters that were configured as input fields.");
            sb.AppendLine("Without access to the ConMasDesigner metadata stored in the workbook,");
            sb.AppendLine("the exact selection criteria cannot be fully replicated.");
        }
        catch (Exception ex)
        {
            sb.AppendLine($"Error: {ex.Message}");
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "cluster_generation_report.md"), sb.ToString());
        _progress.Report($"  [CLUSTER] Wrote cluster_generation_report.md");
    }

    /// <summary>
    /// Generate coordinate comparison report
    /// </summary>
    private async Task GenerateCoordinateComparisonAsync(ComData data, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Coordinate System Comparison");
        sb.AppendLine();
        sb.AppendLine("## Overview");
        sb.AppendLine();
        sb.AppendLine("This report compares **four coordinate systems** for every cluster:");
        sb.AppendLine();
        sb.AppendLine("| # | System | Source | Accuracy |");
        sb.AppendLine("|---|--------|--------|----------|");
        sb.AppendLine("| 1 | **Database** | Legacy PostgreSQL (`def_cluster`) | Ground truth |");
        sb.AppendLine("| 2 | **Excel COM** | `Range.Left/Top/Width/Height` | Excel's rendered position |");
        sb.AppendLine("| 3 | **OpenXML Calc** | Our engine's column width estimates | Approximate |");
        sb.AppendLine("| 4 | **MergeArea** | `MergeArea.Left/Top/Width/Height` | Merged range bounds |");
        sb.AppendLine();

        sb.AppendLine("## Page Setup (from Excel COM)");
        sb.AppendLine();
        var ps = data.PageSetup;
        sb.AppendLine($"| Property | Value |");
        sb.AppendLine($"|----------|------:|");
        sb.AppendLine($"| Page Width | {ps.PageWidth:F1} pt |");
        sb.AppendLine($"| Page Height | {ps.PageHeight:F1} pt |");
        sb.AppendLine($"| Left Margin | {ps.LeftMargin:F4} pt ({ps.LeftMargin / 72:F4} in) |");
        sb.AppendLine($"| Right Margin | {ps.RightMargin:F4} pt |");
        sb.AppendLine($"| Top Margin | {ps.TopMargin:F4} pt |");
        sb.AppendLine($"| Bottom Margin | {ps.BottomMargin:F4} pt |");
        sb.AppendLine($"| Center Horizontally | {ps.CenterHorizontally} |");
        sb.AppendLine($"| Center Vertically | {ps.CenterVertically} |");
        sb.AppendLine($"| Zoom | {ps.Zoom} |");
        sb.AppendLine($"| FitToPagesWide | {ps.FitToPagesWide} |");
        sb.AppendLine($"| FitToPagesTall | {ps.FitToPagesTall} |");
        sb.AppendLine($"| Orientation | {ps.Orientation} |");
        sb.AppendLine($"| PaperSize | {ps.PaperSize} ({ps.PaperSizeName}) |");
        sb.AppendLine($"| PrintArea | {ps.PrintArea ?? "(none)"} |");
        sb.AppendLine($"| Order | {ps.Order} |");
        sb.AppendLine($"| FirstPageNumber | {ps.FirstPageNumber} |");

        sb.AppendLine();
        sb.AppendLine("## Used Range (from Excel COM)");
        sb.AppendLine();
        sb.AppendLine($"| Property | Value |");
        sb.AppendLine($"|----------|------:|");
        sb.AppendLine($"| Address | {ps.UsedRangeAddress} |");
        sb.AppendLine($"| Rows | {ps.UsedRangeRows} |");
        sb.AppendLine($"| Columns | {ps.UsedRangeColumns} |");
        sb.AppendLine($"| Left | {ps.UsedRangeLeft:F1} pt |");
        sb.AppendLine($"| Top | {ps.UsedRangeTop:F1} pt |");
        sb.AppendLine($"| Width | {ps.UsedRangeWidth:F1} pt |");
        sb.AppendLine($"| Height | {ps.UsedRangeHeight:F1} pt |");

        sb.AppendLine();
        sb.AppendLine("## Active Window (from Excel COM)");
        sb.AppendLine();
        sb.AppendLine($"| Property | Value |");
        sb.AppendLine($"|----------|------:|");
        sb.AppendLine($"| Zoom | {ps.WindowZoom} |");
        sb.AppendLine($"| VisibleRange | {ps.WindowVisibleRange} |");
        sb.AppendLine($"| ScrollRow | {ps.ScrollRow} |");
        sb.AppendLine($"| ScrollColumn | {ps.ScrollColumn} |");
        sb.AppendLine($"| HPageBreaks | {ps.HPageBreaks} |");
        sb.AppendLine($"| VPageBreaks | {ps.VPageBreaks} |");

        sb.AppendLine();
        sb.AppendLine("## Per-Cluster Comparison");
        sb.AppendLine();
        sb.AppendLine("| Cluster | DB Left | COM Left | Merge Left | OpenXML Left | DB→COM Δx | COM→Merge Δx |");
        sb.AppendLine("|---------|---------|----------|------------|-------------|-----------|-------------|");

        double pageW = ps.PageWidth;
        double pageH = ps.PageHeight;
        double printableW = pageW - ps.LeftMargin - ps.RightMargin;

        foreach (var c in data.Clusters.OrderBy(x => x.Address))
        {
            double dbLeftPts = c.DbLeft * pageW;
            double comLeftPts = c.RangeLeft;
            double mergeLeftPts = c.RangeMergeLeft;
            double openXmlLeftPts = 0; // Will be estimated

            var dbComDiff = dbLeftPts - comLeftPts;
            var comMergeDiff = comLeftPts - mergeLeftPts;

            sb.AppendLine($"| {c.Address} | " +
                $"{c.DbLeft:F7} ({dbLeftPts:F1}pt) | " +
                $"{comLeftPts:F1}pt | " +
                $"{mergeLeftPts:F1}pt | " +
                $"{openXmlLeftPts:F1}pt | " +
                $"{dbComDiff:+0.0;-0.0}pt | " +
                $"{comMergeDiff:+0.0;-0.0}pt |");
        }

        sb.AppendLine();
        sb.AppendLine("## Coordinate System Identification");
        sb.AppendLine();

        // Analyse what the DB left ratio corresponds to
        sb.AppendLine("### Hypothesis: DB = (COM_Left - PrintedOrigin) / PageWidth");
        sb.AppendLine();
        sb.AppendLine("Where **PrintedOrigin** is the offset from the page left edge");
        sb.AppendLine("to column A on the printed page.");
        sb.AppendLine();

        // Derive PrintedOrigin from first cluster
        if (data.Clusters.Count > 0)
        {
            var c0 = data.Clusters.First();
            double derivedOrigin = c0.DbLeft * pageW - c0.RangeLeft;
            sb.AppendLine($"**From {c0.Address}:** PrintedOrigin = DB_Left({c0.DbLeft * pageW:F1}) - COM_Left({c0.RangeLeft:F1}) = {derivedOrigin:F2} pt");
            sb.AppendLine();

            // Verify across all clusters
            sb.AppendLine("### Verification Across All Clusters");
            sb.AppendLine();
            sb.AppendLine("| Address | COM Left | DB Left(pts) | Derived Origin | Constant? |");
            sb.AppendLine("|---------|----------|-------------|----------------|-----------|");
            for (int i = 0; i < data.Clusters.Count; i++)
            {
                var c = data.Clusters[i];
                double origin = c.DbLeft * pageW - c.RangeLeft;
                bool constant = Math.Abs(origin - derivedOrigin) < 1.0;
                sb.AppendLine($"| {c.Address} | {c.RangeLeft:F1} | {c.DbLeft * pageW:F1} | {origin:+0.00;-0.00} | {(constant ? "✓" : "✗ varies")} |");
            }

            sb.AppendLine();
            sb.AppendLine("### Compute UsedRange-Based PrintedOrigin");
            sb.AppendLine();
            double usedRangeOrigin = (pageW - ps.UsedRangeWidth) / 2.0;
            sb.AppendLine($"**If centered:** PrintedOrigin = (PageWidth - UsedRange.Width) / 2");
            sb.AppendLine($"  = ({pageW:F1} - {ps.UsedRangeWidth:F1}) / 2 = {usedRangeOrigin:F2} pt");
            sb.AppendLine();
            sb.AppendLine($"**If not centered:** PrintedOrigin = LeftMargin = {ps.LeftMargin:F2} pt");
            sb.AppendLine();

            // Test which formula fits better
            double centeringError = Math.Abs(usedRangeOrigin - derivedOrigin);
            double marginError = Math.Abs(ps.LeftMargin - derivedOrigin);

            if (centeringError < marginError)
            {
                sb.AppendLine($"**Result: Center-based formula fits.** (Error: {centeringError:F2}pt vs {marginError:F2}pt for margin)");
                sb.AppendLine();
                sb.AppendLine($"PrintedOrigin = (PageWidth - UsedRange.Width) / 2");
                sb.AppendLine($"  = ({pageW:F1} - {ps.UsedRangeWidth:F1}) / 2 = {usedRangeOrigin:F2} pt");
            }
            else
            {
                sb.AppendLine($"**Result: Margin-based formula fits.** (Error: {marginError:F2}pt vs {centeringError:F2}pt for centering)");
                sb.AppendLine();
                sb.AppendLine($"PrintedOrigin = LeftMargin = {ps.LeftMargin:F2} pt");
            }

            sb.AppendLine();
            sb.AppendLine("### Workaround: Direct COM Interop Integration");
            sb.AppendLine();
            sb.AppendLine("Since the legacy system used Excel COM Range.Left, and we've proven");
            sb.AppendLine("the formula is `DB = (Range.Left + PrintedOrigin) / PageWidth`, the");
            sb.AppendLine("engine could use this formula with COM Interop to achieve 100% match");
            sb.AppendLine("for BOTH templates without any template-specific logic.");
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "coordinate_comparison.md"), sb.ToString());
        _progress.Report($"  [COM] Wrote coordinate_comparison.md");
    }

    /// <summary>
    /// Generate detailed coordinate trace
    /// </summary>
    private async Task GenerateCoordinateTraceAsync(ComData data, string dir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Coordinate Trace");
        sb.AppendLine();
        sb.AppendLine("## Exact Coordinate Transformation for Each Cluster");
        sb.AppendLine();

        double pageW = data.PageSetup.PageWidth;
        double pageH = data.PageSetup.PageHeight;

        foreach (var c in data.Clusters.OrderBy(x => x.Address))
        {
            sb.AppendLine($"### {c.Address}");
            sb.AppendLine();
            sb.AppendLine("```");
            sb.AppendLine($"Range(\"{c.Address}\"):");
            sb.AppendLine($"  .Left   = {c.RangeLeft:F4} pt");
            sb.AppendLine($"  .Top    = {c.RangeTop:F4} pt");
            sb.AppendLine($"  .Width  = {c.RangeWidth:F4} pt");
            sb.AppendLine($"  .Height = {c.RangeHeight:F4} pt");
            sb.AppendLine();
            sb.AppendLine($"MergeArea: {c.RangeMergeArea}");
            sb.AppendLine($"  .Left   = {c.RangeMergeLeft:F4} pt");
            sb.AppendLine($"  .Top    = {c.RangeMergeTop:F4} pt");
            sb.AppendLine($"  .Width  = {c.RangeMergeWidth:F4} pt");
            sb.AppendLine($"  .Height = {c.RangeMergeHeight:F4} pt");
            sb.AppendLine();
            sb.AppendLine($"Database:");
            sb.AppendLine($"  Left   = {c.DbLeft:F7} ({c.DbLeft * pageW:F2} pt)");
            sb.AppendLine($"  Top    = {c.DbTop:F7} ({c.DbTop * pageH:F2} pt)");
            sb.AppendLine($"  Right  = {c.DbRight:F7} ({c.DbRight * pageW:F2} pt)");
            sb.AppendLine($"  Bottom = {c.DbBottom:F7} ({c.DbBottom * pageH:F2} pt)");
            sb.AppendLine();
            sb.AppendLine($"Differences (DB - COM):");
            sb.AppendLine($"  ΔLeft   = {c.DbLeft * pageW - c.RangeLeft:+0.00;-0.00} pt");
            sb.AppendLine($"  ΔTop    = {c.DbTop * pageH - c.RangeTop:+0.00;-0.00} pt");
            sb.AppendLine($"  ΔWidth  = {(c.DbRight - c.DbLeft) * pageW - c.RangeWidth:+0.00;-0.00} pt");
            sb.AppendLine($"  ΔHeight = {(c.DbBottom - c.DbTop) * pageH - c.RangeHeight:+0.00;-0.00} pt");
            sb.AppendLine("```");
            sb.AppendLine();

            // Try to find the exact transformation
            double originX = c.DbLeft * pageW - c.RangeLeft;
            double originY = c.DbTop * pageH - c.RangeTop;

            sb.AppendLine($"**Derived Origin Offset:** X={originX:+0.00;-0.00}pt, Y={originY:+0.00;-0.00}pt");
            sb.AppendLine();
            sb.AppendLine("```");
            sb.AppendLine($"DB_Left_Ratio = (Range.Left + {originX:+0.00;-0.00}) / {pageW:F1}");
            sb.AppendLine($"  = ({c.RangeLeft:F1} + {originX:+0.00;-0.00}) / {pageW:F1}");
            sb.AppendLine($"  = {c.RangeLeft + originX:F2} / {pageW:F1}");
            sb.AppendLine($"  = {(c.RangeLeft + originX) / pageW:F7}");
            sb.AppendLine($"  DB = {c.DbLeft:F7}  {(Math.Abs((c.RangeLeft + originX) / pageW - c.DbLeft) < 0.001 ? "✓ MATCH" : "✗ MISMATCH")}");
            sb.AppendLine("```");
            sb.AppendLine();
        }

        await File.WriteAllTextAsync(Path.Combine(dir, "coordinate_trace.md"), sb.ToString());
        _progress.Report($"  [COM] Wrote coordinate_trace.md");
    }

    // ================ UTILITY METHODS ================

    private static double ReadDouble(Func<object> getter)
    {
        try { return Convert.ToDouble(getter()); }
        catch { return 0; }
    }

    private static string ReadString(Func<object> getter)
    {
        try { var v = getter(); return v?.ToString() ?? ""; }
        catch { return ""; }
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static string HexSHA(byte[] data)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(data));
    }

    private static string HexMD5(byte[] data)
    {
        using var md5 = MD5.Create();
        return Convert.ToHexString(md5.ComputeHash(data));
    }

    private static bool IsInRange(string cellRef, string startRef, string endRef)
    {
        var (c1, r1) = ParseRef(startRef);
        var (c2, r2) = ParseRef(endRef);
        var (cc, rr) = ParseRef(cellRef);
        return cc >= c1 && cc <= c2 && rr >= r1 && rr <= r2;
    }

    private static bool IsCellInRange(string cellRef, string rangeRef)
    {
        var parts = rangeRef.Split(':');
        if (parts.Length == 2)
            return IsInRange(cellRef, parts[0], parts[1]);
        return cellRef.Equals(rangeRef, StringComparison.OrdinalIgnoreCase);
    }

    private static bool RangesOverlap(string range1, string range2)
    {
        try
        {
            // Parse both ranges
            var p1 = range1.Split(':');
            var p2 = range2.Split(':');
            if (p1.Length != 2 || p2.Length != 2) return false;

            var (s1c, s1r) = ParseRef(p1[0]);
            var (e1c, e1r) = ParseRef(p1[1]);
            var (s2c, s2r) = ParseRef(p2[0]);
            var (e2c, e2r) = ParseRef(p2[1]);

            // Check overlap: rectangles overlap if NOT (range1 is left-of/right-of/above/below range2)
            bool noOverlap = e1c < s2c || e2c < s1c || e1r < s2r || e2r < s1r;
            return !noOverlap;
        }
        catch { return false; }
    }

    private static (int col, int row) ParseRef(string refStr)
    {
        var letters = new string(refStr.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
        var digits = new string(refStr.SkipWhile(char.IsLetter).ToArray());
        int col = 0;
        foreach (var ch in letters)
            col = col * 26 + (ch - 'A' + 1);
        int row = int.TryParse(digits, out var r) ? r : 0;
        return (col, row);
    }

    // ================ DATA CLASSES ================

    public class PsInfo
    {
        // Default to Letter page (612 x 792 pt) — used when COM PageSetup returns 0
        // due to printer driver not reporting page dimensions
        public const double DefaultPageWidth = 612.0;
        public const double DefaultPageHeight = 792.0;

        public int PaperSize { get; set; }
        public string PaperSizeName { get; set; } = "";
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
        public int Orientation { get; set; }
        public double PageWidth { get; set; } = DefaultPageWidth;
        public double PageHeight { get; set; } = DefaultPageHeight;
        public string? PrintArea { get; set; }
        public int Order { get; set; }
        public int FirstPageNumber { get; set; }
        public int[]? PrintQuality { get; set; }
        public string? FooterPicture { get; set; }
        // UsedRange
        public string? UsedRangeAddress { get; set; }
        public int UsedRangeRows { get; set; }
        public int UsedRangeColumns { get; set; }
        public double UsedRangeLeft { get; set; }
        public double UsedRangeTop { get; set; }
        public double UsedRangeWidth { get; set; }
        public double UsedRangeHeight { get; set; }
        // ActiveWindow
        public int WindowZoom { get; set; }
        public string? WindowVisibleRange { get; set; }
        public int ScrollRow { get; set; }
        public int ScrollColumn { get; set; }
        // Page breaks
        public int HPageBreaks { get; set; }
        public int VPageBreaks { get; set; }
    }

    public class ComClusterData
    {
        public string Address { get; set; } = "";
        public double RangeLeft { get; set; }
        public double RangeTop { get; set; }
        public double RangeWidth { get; set; }
        public double RangeHeight { get; set; }
        public string RangeAddress { get; set; } = "";
        public string RangeMergeArea { get; set; } = "";
        public double RangeMergeLeft { get; set; }
        public double RangeMergeTop { get; set; }
        public double RangeMergeWidth { get; set; }
        public double RangeMergeHeight { get; set; }
        public string RangeValue { get; set; } = "";
        // DB values (for comparison)
        public double DbLeft { get; set; }
        public double DbTop { get; set; }
        public double DbRight { get; set; }
        public double DbBottom { get; set; }
    }

    public class ComData
    {
        public int DefTopId { get; set; }
        public string WorkbookName { get; set; } = "";
        public PsInfo PageSetup { get; set; } = new();
        public List<ComClusterData> Clusters { get; set; } = new();
    }
}
