using System.Reflection;
using System.Runtime.Loader;
using System.Text;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Inspects legacy .NET assemblies from the installed PaperLess applications
/// to discover the coordinate generation algorithm.
/// </summary>
public class AssemblyInspector
{
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;

    public AssemblyInspector(IProgress<string> progress)
    {
        _progress = progress;
        _outputDir = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\Test Folder Final Test\Phase15_LegacyAlgorithmReconstruction";
        Directory.CreateDirectory(_outputDir);
    }

    public async Task RunAsync()
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Legacy Assembly Decompilation Report");
        sb.AppendLine();
        sb.AppendLine("**Generated:** " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        sb.AppendLine();
        sb.AppendLine("## Overview");
        sb.AppendLine();
        sb.AppendLine("This report documents the legacy PaperLess application assemblies found at:");
        sb.AppendLine();
        sb.AppendLine("- `C:\\Program Files (x86)\\CIMTOPS CORPORATION\\iReporterExcelAddIn\\`");
        sb.AppendLine("- `C:\\Program Files (x86)\\CIMTOPS CORPORATION\\ConMas Generator\\`");
        sb.AppendLine("- `C:\\Program Files (x86)\\CIMTOPS CORPORATION\\ConMas Designer\\`");
        sb.AppendLine("- `C:\\Program Files (x86)\\CIMTOPS CORPORATION\\ConMas i-Reporter for Windows\\`");
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();

        var assemblies = new (string Path, string Label)[]
        {
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\Cimtops.Excel.dll", "iReporterExcelAddIn — Cimtops.Excel"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\Cimtops.R2Cluster.dll", "iReporterExcelAddIn — Cimtops.R2Cluster"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\iReporterExcelAddIn.dll", "iReporterExcelAddIn — Main AddIn"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\iReporterExcelAddInCommon.dll", "iReporterExcelAddIn — Common"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasGeneratorLib.dll", "ConMas Generator — GeneratorLib"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasJob.exe", "ConMas Generator — Job Executable"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasTool.exe", "ConMas Generator — Tool Executable"),
            (@"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\ConMasClient.exe", "ConMas Designer — Main Client"),
        };

        foreach (var (path, label) in assemblies)
        {
            if (!File.Exists(path))
            {
                sb.AppendLine($"## {label}");
                sb.AppendLine();
                sb.AppendLine($"File not found: `{path}`");
                sb.AppendLine();
                continue;
            }
            await InspectAssembly(path, label, sb);
        }

        var reportPath = Path.Combine(_outputDir, "Phase15_DecompilationReport.md");
        await File.WriteAllTextAsync(reportPath, sb.ToString());
        _progress.Report($"Report written: {reportPath}");
    }

    private async Task InspectAssembly(string path, string label, StringBuilder sb)
    {
        _progress.Report($"Inspecting: {label} ({path})");

        try
        {
            var fi = new FileInfo(path);
            sb.AppendLine($"## {label}");
            sb.AppendLine();
            sb.AppendLine($"**File:** `{path}`");
            sb.AppendLine($"**Size:** {fi.Length:N0} bytes");
            sb.AppendLine();

            // PE header analysis
            using var stream = File.OpenRead(path);
            using var reader = new BinaryReader(stream);
            if (reader.ReadUInt16() != 0x5A4D)
            {
                sb.AppendLine("**Type:** Not a valid PE executable (no MZ header)");
                sb.AppendLine();
                return;
            }

            reader.BaseStream.Position = 0x3C;
            var peOffset = reader.ReadInt32();
            reader.BaseStream.Position = peOffset;

            if (reader.ReadUInt32() != 0x00004550)
            {
                sb.AppendLine("**Type:** Not a valid PE executable (no PE signature)");
                sb.AppendLine();
                return;
            }

            var machine = reader.ReadUInt16();
            reader.ReadUInt16(); // sections
            reader.ReadUInt32(); // Timestamp
            reader.ReadUInt32(); // PointerToSymbolTable
            reader.ReadUInt32(); // NumberOfSymbols
            reader.ReadUInt16(); // OptionalHeaderSize
            reader.ReadUInt16(); // Characteristics

            var arch = machine switch
            {
                0x014C => "x86 (32-bit)",
                0x8664 => "x64 (64-bit)",
                0x01C4 => "ARM",
                _ => $"Unknown (0x{machine:X4})"
            };

            var magic = reader.ReadUInt16();
            var isPE32Plus = magic == 0x020B;
            reader.BaseStream.Position = (long)peOffset + (isPE32Plus ? 248L : 232L);
            var clrRva = reader.ReadUInt32();
            var clrSize = reader.ReadUInt32();
            var isManaged = clrRva != 0 && clrSize != 0;

            sb.AppendLine($"**Architecture:** {arch}");
            sb.AppendLine($"**Managed (.NET) assembly:** {(isManaged ? "YES" : "NO — Native DLL")}");
            sb.AppendLine($"**CLR Header:** RVA=0x{clrRva:X8}, Size=0x{clrSize:X8}");
            sb.AppendLine();

            if (isManaged)
            {
                await LoadAndReflectAssembly(path, sb);
            }
            else
            {
                sb.AppendLine("*(Native DLL — no managed code to decompile)*");
                sb.AppendLine();
            }
        }
        catch (Exception ex)
        {
            sb.AppendLine($"**Error:** {ex.GetType().Name}: {ex.Message}");
            sb.AppendLine();
        }
    }

    private async Task LoadAndReflectAssembly(string path, StringBuilder sb)
    {
        try
        {
            var resolver = new CustomAssemblyResolver(path);
            var assembly = resolver.LoadAssembly();

            if (assembly == null)
            {
                sb.AppendLine("**Reflection:** Could not load assembly. Falling back to string analysis.");
                sb.AppendLine();
                await StringAnalysis(path, sb);
                return;
            }

            sb.AppendLine("### Types");
            sb.AppendLine();
            sb.AppendLine("| Type | Kind | Public Methods | Public Fields |");
            sb.AppendLine("|------|------|---------------|--------------|");

            var types = assembly.GetTypes()
                .OrderBy(t => t.FullName)
                .ToArray();

            foreach (var type in types)
            {
                if (type.FullName == null || type.FullName.StartsWith("<")) continue;

                var kind = type.IsInterface ? "Interface" :
                           type.IsEnum ? "Enum" :
                           type.IsValueType ? "Struct" :
                           type.IsAbstract ? "Abstract Class" : "Class";

                var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance |
                    BindingFlags.Static | BindingFlags.DeclaredOnly)
                    .Where(m => !m.IsSpecialName)
                    .Select(m => m.Name).Distinct().ToArray();

                var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance |
                    BindingFlags.Static | BindingFlags.DeclaredOnly)
                    .Select(f => f.Name).ToArray();

                sb.AppendLine($"| `{type.FullName}` | {kind} | {methods.Length} | {fields.Length} |");
            }

            sb.AppendLine();

            // Detailed method listing for coordinate-related types
            var keywords = new[] { "Coordinate", "Cluster", "Template", "Worksheet", "Workbook",
                "Excel", "Import", "Export", "Publish", "Designer", "Field", "Range",
                "PrintArea", "Background", "Image", "XmlData", "DefCluster", "DefTop",
                "Normalize", "Scale", "Origin", "Cimtops", "R2", "Sheet" };

            var relevant = types.Where(t => t.FullName != null && !t.FullName.StartsWith("<") &&
                keywords.Any(k => t.FullName.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                .ToArray();

            if (relevant.Length > 0)
            {
                sb.AppendLine("### Detailed Methods (Coordinate-Related Types)");
                sb.AppendLine();

                foreach (var type in relevant)
                {
                    sb.AppendLine($"#### `{type.FullName}`");
                    sb.AppendLine();
                    sb.AppendLine("| Method | Access | Return | Parameters |");
                    sb.AppendLine("|--------|--------|--------|------------|");

                    var allMethods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance |
                        BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.DeclaredOnly)
                        .Where(m => !m.IsSpecialName)
                        .OrderBy(m => m.Name)
                        .ToArray();

                    foreach (var method in allMethods)
                    {
                        var access = method.IsPublic ? "public" :
                                     method.IsFamily ? "protected" :
                                     method.IsPrivate ? "private" : "internal";
                        var stat = method.IsStatic ? " static" : "";
                        var ret = method.ReturnType.Name;
                        var pars = string.Join(", ", method.GetParameters()
                            .Select(p => $"{p.ParameterType.Name} {p.Name}"));

                        sb.AppendLine($"| `{method.Name}` | `{access}{stat}` | `{ret}` | `{pars}` |");
                    }
                    sb.AppendLine();
                }
            }

            resolver.Dispose();
        }
        catch (Exception ex)
        {
            sb.AppendLine($"**Reflection error:** {ex.GetType().Name}: {ex.Message}");
            sb.AppendLine();
            await StringAnalysis(path, sb);
        }
    }

    private async Task StringAnalysis(string path, StringBuilder sb)
    {
        sb.AppendLine("### String-Based Analysis (Fallback)");
        sb.AppendLine();

        try
        {
            var bytes = await File.ReadAllBytesAsync(path);
            var found = new HashSet<string>();

            for (int i = 0; i < bytes.Length - 1; i++)
            {
                if (bytes[i] >= 32 && bytes[i] < 127 && bytes[i + 1] == 0)
                {
                    var sbs = new StringBuilder();
                    int j = i;
                    while (j < bytes.Length - 1 && bytes[j] >= 32 && bytes[j] < 127 && bytes[j + 1] == 0)
                    {
                        sbs.Append((char)bytes[j]);
                        j += 2;
                    }
                    var s = sbs.ToString();
                    if (s.Length >= 4 && !found.Contains(s))
                        found.Add(s);
                }
            }

            var relevant = found
                .Where(s => s.Any(char.IsUpper) && !s.All(char.IsUpper) && s.Length >= 4)
                .OrderBy(s => s)
                .ToArray();

            sb.AppendLine("```");
            foreach (var s in relevant.Take(200))
                sb.AppendLine(s);
            if (relevant.Length > 200)
                sb.AppendLine($"... and {relevant.Length - 200} more");
            sb.AppendLine("```");
        }
        catch (Exception ex)
        {
            sb.AppendLine($"String analysis failed: {ex.Message}");
        }
    }

    private class CustomAssemblyResolver : IDisposable
    {
        private readonly string _assemblyPath;
        private readonly string _assemblyDir;
        private readonly AssemblyLoadContext _context;

        public CustomAssemblyResolver(string assemblyPath)
        {
            _assemblyPath = assemblyPath;
            _assemblyDir = Path.GetDirectoryName(assemblyPath) ?? ".";
            _context = new AssemblyLoadContext("LegacyInspector", isCollectible: true);
            _context.Resolving += OnResolving;
        }

        public Assembly? LoadAssembly()
        {
            try
            {
                var bytes = File.ReadAllBytes(_assemblyPath);
                return _context.LoadFromStream(new MemoryStream(bytes));
            }
            catch
            {
                return null;
            }
        }

        private Assembly? OnResolving(AssemblyLoadContext context, AssemblyName assemblyName)
        {
            var searchDirs = new[]
            {
                _assemblyDir,
                @"C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn",
                @"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator",
                @"C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin",
            };

            foreach (var dir in searchDirs)
            {
                var dllPath = Path.Combine(dir, assemblyName.Name + ".dll");
                if (File.Exists(dllPath))
                {
                    try { return context.LoadFromAssemblyPath(dllPath); }
                    catch { }
                }
            }
            return null;
        }

        public void Dispose()
        {
            _context.Resolving -= OnResolving;
            _context.Unload();
        }
    }
}
