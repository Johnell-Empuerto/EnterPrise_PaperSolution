using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;
using LegacyExtractionEngine.Services;
using LegacyExtractionEngine.Services.Importer;

Console.OutputEncoding = System.Text.Encoding.UTF8;

var connectionString = $"Host=localhost;Port=5432;Database=irepodb;Username=postgres;Password=cimtops";
var testFolder = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\Test Folder Final Test";

IProgress<string> progress = new Progress<string>(msg =>
{
    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {msg}");
});

var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "--phase8";

if (mode == "--trace" || mode == "trace")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Coordinate System Investigation");
    Console.WriteLine(" Using Excel COM Interop");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var tracer = new CoordinateTracer(connectionString, progress);
    await tracer.TraceAllAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" OUTPUT: " + testFolder);
    Console.WriteLine(" STATUS: Trace Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--gap" || mode == "gap")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 5: Coordinate Gap Investigation");
    Console.WriteLine(" Tracing column widths and row heights via COM");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var gapInvestigator = new CoordinateGapInvestigator(progress);
    await gapInvestigator.InvestigateAllAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" OUTPUT: " + testFolder);
    Console.WriteLine(" STATUS: Investigation Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--list-templates" || mode == "list-templates")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Template Discovery");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var discovery = new TemplateDiscoveryService(connectionString, progress);
    var templates = await discovery.DiscoverAllTemplatesAsync();

    Console.WriteLine();
    Console.WriteLine($"Found {templates.Count} templates:");
    Console.WriteLine();
    Console.WriteLine("| ID | Name | Clusters | Workbook |");
    Console.WriteLine("|----|------|----------|----------|");
    foreach (var t in templates)
    {
        var wb = t.HasWorkbook ? "✓" : "✗";
        Console.WriteLine($"| {t.DefTopId} | {t.DefTopName ?? ""} | {t.ClusterCount} | {wb} |");
    }
    return;
}
else if (mode == "--phase6-engine" || mode == "phase6-engine")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 6 — Universal Coordinate Engine");
    Console.WriteLine(" Engine Validation Against All Templates");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var validator = new Phase6EngineValidator(connectionString, progress);
    await validator.ValidateAllAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Engine Validation Complete");
    Console.WriteLine("============================================");
}else if (mode == "--phase8" || mode == "phase8")
    {
        Console.WriteLine("============================================");
        Console.WriteLine(" Phase 8 — Universal Engine Validation");
        Console.WriteLine(" Production Integration & 100-Template Test");
        Console.WriteLine("============================================");
        Console.WriteLine();

        var phase8 = new Phase8Validator(connectionString, progress);
        await phase8.RunValidationAsync();

        Console.WriteLine();
        Console.WriteLine("============================================");
        Console.WriteLine(" STATUS: Phase 8 Validation Complete");
        Console.WriteLine("============================================");
    }
    else if (mode == "--phase7" || mode == "phase7")
    {
        Console.WriteLine("============================================");
        Console.WriteLine(" Phase 7 — Missing Coordinate Transform");
        Console.WriteLine(" Per-Cluster Error Analysis & Regression");
        Console.WriteLine("============================================");
        Console.WriteLine();

        var phase7 = new Phase7ErrorInvestigator(connectionString, progress);
        await phase7.RunInvestigationAsync();

        Console.WriteLine();
        Console.WriteLine("============================================");
        Console.WriteLine(" STATUS: Phase 7 Investigation Complete");
        Console.WriteLine("============================================");
    }
else if (mode == "--phase6" || mode == "phase6")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 6 — Universal Coordinate Engine");
    Console.WriteLine(" Classification & Strategy Analysis");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var classifier = new Phase6Classifier(connectionString, progress);
    await classifier.RunClassificationAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Classification Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--investigate" || mode == "investigate")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Comprehensive Legacy Coordinate Investigation");
    Console.WriteLine(" COM + Images + Cluster Patterns");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var investigator = new LegacyCoordinateInvestigator(connectionString, progress);
    await investigator.InvestigateAllAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" OUTPUT: " + testFolder);
    Console.WriteLine(" STATUS: Investigation Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--phase17" || mode == "phase17")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 17 — Legacy Coordinate Reconstruction");
    Console.WriteLine(" Column-Width/Row-Height Summation Algorithm");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var engine = new Phase17CoordinateReconstructor(progress);
    await engine.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 17 Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--inspect" || mode == "inspect")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Legacy Assembly Inspection");
    Console.WriteLine(" Decompiling installed PaperLess applications");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var inspector = new AssemblyInspector(progress);
    await inspector.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Inspection Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--phase14" || mode == "phase14")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 14 — Gap Origin Investigation");
    Console.WriteLine(" Templates: 546, 547, 548");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var investigator = new Phase14GapInvestigator(connectionString, progress);
    await investigator.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 14 Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--phase13" || mode == "phase13")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 13 — Legacy Transform Derivation");
    Console.WriteLine(" Reverse engineering the exact transform");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var engine = new Phase13TransformEngine(connectionString, progress);
    await engine.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 13 Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--phase12" || mode == "phase12")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 12 — Coordinate Source Investigation");
    Console.WriteLine(" Dumping every Excel coordinate source");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var scanner = new Phase12CoordinateScanner(connectionString, progress);
    await scanner.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 12 Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--phase11" || mode == "phase11")
{
    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 11 — Legacy Coordinate Reverse Engineering");
    Console.WriteLine(" Templates: 546, 547, 548");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var engineer = new Phase11ReverseEngineer(connectionString, progress);
    await engineer.RunAsync();

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 11 Complete");
    Console.WriteLine("============================================");
}
else if (mode == "--validate" || mode == "validate")
{
    // Parse template IDs from remaining args, default to 546 547 548
    var templateIds = args.Skip(1).Select(a => int.TryParse(a, out var id) ? id : 0)
        .Where(id => id > 0).Distinct().ToArray();
    if (templateIds.Length == 0)
        templateIds = new[] { 546, 547, 548 };

    Console.WriteLine("============================================");
    Console.WriteLine(" Phase 10 — Final Template Validation");
    Console.WriteLine($" Validating: {string.Join(", ", templateIds)}");
    Console.WriteLine("============================================");
    Console.WriteLine();

    var validator = new Phase10FinalValidator(connectionString, progress);
    await validator.RunValidationAsync(templateIds);

    Console.WriteLine();
    Console.WriteLine("============================================");
    Console.WriteLine(" STATUS: Phase 10 Complete");
    Console.WriteLine("============================================");
}
else
{
    Console.WriteLine("============================================");
    Console.WriteLine(" PaperLess Enterprise - Phase 4 Validation");
        Console.WriteLine(" Full Legacy Import Validation (546 & 547)");
        Console.WriteLine(" Using Excel COM for coordinate generation");
        Console.WriteLine("============================================");
        Console.WriteLine();

        try
        {
            Directory.CreateDirectory(testFolder);

            // Phase 1: Extract COM data for both templates (outside validator to open/close Excel only once per workbook)
            Console.WriteLine("=== Extracting COM coordinate data ===");
            Console.WriteLine();

            var comData = new Dictionary<int, ComExtraction?>();

            // Template 546
            try
            {
                var dbReader = new DatabaseReader(connectionString, progress);
                var db546 = await dbReader.ReadAllAsync(546);
                var clusters546 = db546.DefSheets.SelectMany(s => s.Clusters).ToList();

                // Extract PrintArea from OpenXML defined names (most reliable source)
                var openXmlPrintArea546 = ReadPrintAreaFromOpenXml(@"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx");

                using var com546 = new ComCoordinateService(progress);
                var extract546 = com546.Extract(@"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx",
                    clusters546, openXmlPrintArea546);
                comData[546] = extract546;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WARN] COM extraction for 546 failed: {ex.Message}");
                Console.WriteLine("Will fall back to OpenXML coordinate estimation.");
                comData[546] = null;
            }

            // Template 547
            try
            {
                var dbReader = new DatabaseReader(connectionString, progress);
                var db547 = await dbReader.ReadAllAsync(547);
                var clusters547 = db547.DefSheets.SelectMany(s => s.Clusters).ToList();
                var discovery = new TemplateDiscoveryService(connectionString, progress);
                var templates = await discovery.DiscoverAllTemplatesAsync();
                var tpl547 = templates.FirstOrDefault(t => t.DefTopId == 547);
                var wb547 = tpl547?.WorkbookPath ?? "";
                if (!File.Exists(wb547))
                {
                    var dir = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents";
                    wb547 = Directory.GetFiles(dir, "*V3.1*", SearchOption.TopDirectoryOnly).FirstOrDefault() ?? "";
                }
                if (File.Exists(wb547))
                {
                    // Extract PrintArea from OpenXML defined names
                    var openXmlPrintArea547 = ReadPrintAreaFromOpenXml(wb547);

                    using var com547 = new ComCoordinateService(progress);
                    var extract547 = com547.Extract(wb547, clusters547, openXmlPrintArea547);
                    comData[547] = extract547;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WARN] COM extraction for 547 failed: {ex.Message}");
                Console.WriteLine("Will fall back to OpenXML coordinate estimation.");
                comData[547] = null;
            }

            Console.WriteLine();
            Console.WriteLine("=== Running validation ===");
            Console.WriteLine();

            var validator = new Phase4Validator(connectionString, progress);
            await validator.RunAllAsync(comData);

            Console.WriteLine();
            Console.WriteLine("============================================");
            Console.WriteLine($" OUTPUT: {testFolder}");
            Console.WriteLine(" STATUS: Complete");
            Console.WriteLine("============================================");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FATAL ERROR: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

/// <summary>
/// Read the PrintArea range address from OpenXML defined names.
/// The print area is stored as a defined name with name "_xlnm.Print_Area"
/// and refers_to like "Sheet1!$B$1:$M$44".
/// This is more reliable than COM PageSetup.PrintArea which may return
/// printer defaults rather than workbook settings.
/// </summary>
static string? ReadPrintAreaFromOpenXml(string workbookPath)
{
    try
    {
        using var doc = SpreadsheetDocument.Open(workbookPath, false);
        var wbPart = doc.WorkbookPart;
        if (wbPart == null) return null;

        var wbXml = XDocument.Load(wbPart.GetStream());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        // Find the definedName with name="_xlnm.Print_Area"
        var printAreaDef = wbXml.Descendants(ns + "definedName")
            .FirstOrDefault(d =>
            {
                var name = (string?)d.Attribute("name");
                return name != null &&
                    (name == "_xlnm.Print_Area" ||
                     name.EndsWith("Print_Area", StringComparison.OrdinalIgnoreCase));
            });

        if (printAreaDef == null) return null;

        // The value is like "Sheet1!$B$1:$M$44" or "$B$1:$M$44"
        var value = printAreaDef.Value?.Trim();
        if (string.IsNullOrEmpty(value)) return null;

        // Strip sheet name prefix (everything before the !)
        var bangIdx = value.IndexOf('!');
        if (bangIdx >= 0)
            value = value.Substring(bangIdx + 1);

        // Strip single quotes if present (for sheet names with spaces)
        value = value.Replace("'", "");

        Console.WriteLine($"[OpenXML] Read PrintArea from defined name: '{value}'");
        return value;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[OpenXML] Failed to read PrintArea: {ex.Message}");
        return null;
    }
}
