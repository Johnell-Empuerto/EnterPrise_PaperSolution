using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace LegacyExtractionEngine.Services;

public class ForensicComparer
{
    private readonly IProgress<string> _progress;

    public ForensicComparer(IProgress<string> progress)
    {
        _progress = progress;
    }

    public async Task CompareAllAsync(string dbXmlPath, string genXmlPath, string outputDir)
    {
        var dbBytes = await File.ReadAllBytesAsync(dbXmlPath);
        var genBytes = await File.ReadAllBytesAsync(genXmlPath);

        var report = new StringBuilder();

        _progress.Report("[Forensic] Level 1: Raw bytes");
        report.AppendLine("# Forensic Byte Comparison Report");
        report.AppendLine();
        AppendLevel1(report, dbBytes, genBytes);

        var dbText = Encoding.UTF8.GetString(dbBytes);
        var genText = Encoding.UTF8.GetString(genBytes);

        _progress.Report("[Forensic] Level 2: Whitespace");
        AppendLevel2(report, dbText, genText);

        _progress.Report("[Forensic] Level 3: XML declaration");
        AppendLevel3(report, dbText, genText);

        _progress.Report("[Forensic] Level 4-5: Node structure & attributes");
        AppendLevel4And5(report, dbText, genText);

        _progress.Report("[Forensic] Level 6-7: Numeric values");
        AppendLevel6And7(report, dbText, genText);

        _progress.Report("[Forensic] Hex diff");
        var hexDiffPath = Path.Combine(outputDir, "xml_hex_diff.md");
        await GenerateHexDiffAsync(dbBytes, genBytes, hexDiffPath);
        report.AppendLine($"## Hex Diff: {hexDiffPath}");

        var lineDiffPath = Path.Combine(outputDir, "xml_line_diff.md");
        _progress.Report("[Forensic] Line diff");
        await GenerateLineDiffAsync(dbText, genText, lineDiffPath);
        report.AppendLine($"## Line Diff: {lineDiffPath}");

        await File.WriteAllTextAsync(Path.Combine(outputDir, "xml_byte_diff.md"), report.ToString());
        _progress.Report("  Wrote xml_byte_diff.md");
    }

    private static void AppendLevel1(StringBuilder sb, byte[] db, byte[] gen)
    {
        sb.AppendLine("## Level 1: Raw Bytes");
        sb.AppendLine($"| Metric | Legacy (DB) | Generated | Match |");
        sb.AppendLine($"|--------|-------------|-----------|-------|");
        sb.AppendLine($"| **Byte count** | {db.Length} | {gen.Length} | {db.Length == gen.Length} |");

        using var sha256 = SHA256.Create();
        var dbHash = Convert.ToHexString(sha256.ComputeHash(db));
        var genHash = Convert.ToHexString(sha256.ComputeHash(gen));
        sb.AppendLine($"| **SHA256** | {dbHash} | {genHash} | {dbHash == genHash} |");

        using var md5 = MD5.Create();
        var dbMd5 = Convert.ToHexString(md5.ComputeHash(db));
        var genMd5 = Convert.ToHexString(md5.ComputeHash(gen));
        sb.AppendLine($"| **MD5** | {dbMd5} | {genMd5} | {dbMd5 == genMd5} |");

        sb.AppendLine($"> **Overall: {(db.Length == gen.Length && dbHash == genHash ? "IDENTICAL" : "DIFFERENT")}**");
        sb.AppendLine();
    }

    private static void AppendLevel2(StringBuilder sb, string db, string gen)
    {
        sb.AppendLine("## Level 2: Whitespace");
        var dbLines = db.Split('\n');
        var genLines = gen.Split('\n');

        sb.AppendLine($"| Metric | Legacy | Generated | Match |");
        sb.AppendLine($"|--------|--------|-----------|-------|");
        sb.AppendLine($"| **Line count** | {dbLines.Length} | {genLines.Length} | {dbLines.Length == genLines.Length} |");

        bool crlfDb = db.Contains("\r\n");
        bool crlfGen = gen.Contains("\r\n");
        sb.AppendLine($"| **Line endings** | {(crlfDb ? "CRLF" : "LF")} | {(crlfGen ? "CRLF" : "LF")} | {crlfDb == crlfGen} |");

        bool trailingNewlineDb = db.EndsWith("\n");
        bool trailingNewlineGen = gen.EndsWith("\n");
        sb.AppendLine($"| **Trailing newline** | {trailingNewlineDb} | {trailingNewlineGen} | {trailingNewlineDb == trailingNewlineGen} |");

        int blankLinesDb = dbLines.Count(l => l.Trim().Length == 0);
        int blankLinesGen = genLines.Count(l => l.Trim().Length == 0);
        sb.AppendLine($"| **Blank lines** | {blankLinesDb} | {blankLinesGen} | {blankLinesDb == blankLinesGen} |");
        sb.AppendLine();
    }

    private static void AppendLevel3(StringBuilder sb, string db, string gen)
    {
        sb.AppendLine("## Level 3: XML Declaration");
        string ExtractDecl(string xml)
        {
            if (xml.StartsWith("<?"))
            {
                var end = xml.IndexOf("?>");
                return end > 0 ? xml[..(end + 2)] : "";
            }
            return "(none)";
        }
        var dbDecl = ExtractDecl(db);
        var genDecl = ExtractDecl(gen);
        sb.AppendLine($"| | Legacy | Generated |");
        sb.AppendLine($"|---|--------|-----------|");
        sb.AppendLine($"| **Declaration** | `{EscapeMd(dbDecl)}` | `{EscapeMd(genDecl)}` |");
        sb.AppendLine();
    }

    private static void AppendLevel4And5(StringBuilder sb, string db, string gen)
    {
        sb.AppendLine("## Level 4-5: Node Structure & Attributes");
        try
        {
            var dbDoc = XDocument.Parse(db);
            var genDoc = XDocument.Parse(gen);

            var dbNodes = dbDoc.Descendants().ToList();
            var genNodes = genDoc.Descendants().ToList();

            int minCount = Math.Min(dbNodes.Count, genNodes.Count);
            int firstDiff = -1;
            for (int i = 0; i < minCount; i++)
            {
                if (dbNodes[i].Name != genNodes[i].Name)
                {
                    firstDiff = i;
                    break;
                }
            }

            sb.AppendLine($"| Metric | Legacy | Generated |");
            sb.AppendLine($"|--------|--------|-----------|");
            sb.AppendLine($"| **Node count** | {dbNodes.Count} | {genNodes.Count} |");
            sb.AppendLine($"| **Node order match** | {firstDiff == -1} | |");

            if (firstDiff >= 0)
            {
                sb.AppendLine($"| **First divergence** | Node {firstDiff}: `{dbNodes[firstDiff].Name}` vs `{genNodes[firstDiff].Name}` | |");
                var path = new List<string>();
                var n = dbNodes[firstDiff];
                while (n.Parent != null) { path.Add(n.Name.LocalName); n = n.Parent; }
                path.Reverse();
                sb.AppendLine($"| **XPath (expected)** | `/{string.Join("/", path)}` | |");
            }
            sb.AppendLine();

            sb.AppendLine("### Attribute Comparison");
            bool allAttrsMatch = true;
            for (int i = 0; i < minCount; i++)
            {
                var dbAttrs = dbNodes[i].Attributes().OrderBy(a => a.Name.ToString()).ToList();
                var genAttrs = genNodes[i].Attributes().OrderBy(a => a.Name.ToString()).ToList();
                var dbAttrStr = string.Join(" ", dbAttrs.Select(a => $"{a.Name}=\"{a.Value}\""));
                var genAttrStr = string.Join(" ", genAttrs.Select(a => $"{a.Name}=\"{a.Value}\""));
                if (dbAttrStr != genAttrStr)
                {
                    sb.AppendLine($"- Node `{dbNodes[i].Name}` attrs differ: `{dbAttrStr}` vs `{genAttrStr}`");
                    allAttrsMatch = false;
                }
            }
            if (allAttrsMatch)
                sb.AppendLine("- All attributes match.");
            sb.AppendLine();
        }
        catch (Exception ex)
        {
            sb.AppendLine($"Error parsing XML: {ex.Message}");
            sb.AppendLine();
        }
    }

    private static void AppendLevel6And7(StringBuilder sb, string db, string gen)
    {
        sb.AppendLine("## Level 6-7: Numeric Values");
        var dbLines = db.Split('\n');
        var genLines = gen.Split('\n');

        var numericDiffs = new List<(int Line, string Tag, string DbVal, string GenVal)>();

        for (int i = 0; i < Math.Min(dbLines.Length, genLines.Length); i++)
        {
            if (dbLines[i] == genLines[i]) continue;

            var dbTag = ExtractTagValue(dbLines[i]);
            var genTag = ExtractTagValue(genLines[i]);
            if (dbTag.Tag != genTag.Tag) continue;

            if (double.TryParse(dbTag.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var dbNum) &&
                double.TryParse(genTag.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var genNum))
            {
                numericDiffs.Add((i + 1, dbTag.Tag, dbTag.Value, genTag.Value));
            }
        }

        if (numericDiffs.Count == 0)
        {
            sb.AppendLine("No numeric differences found.");
            sb.AppendLine();
            return;
        }

        sb.AppendLine("| Line | Tag | Legacy | Generated | Diff (ULP) |");
        sb.AppendLine("|------|-----|--------|-----------|------------|");
        foreach (var (line, tag, dbVal, genVal) in numericDiffs)
        {
            var dbNum = double.Parse(dbVal, System.Globalization.CultureInfo.InvariantCulture);
            var genNum = double.Parse(genVal, System.Globalization.CultureInfo.InvariantCulture);
            var ulp = Math.Abs(dbNum - genNum) * 1e7;
            sb.AppendLine($"| {line} | `{tag}` | {dbVal} | {genVal} | {ulp:F2} ULP |");
        }
        sb.AppendLine();
    }

    private static (string Tag, string Value) ExtractTagValue(string line)
    {
        line = line.Trim();
        if (!line.StartsWith("<") || !line.Contains(">")) return ("", "");
        var endTag = line.IndexOf('>');
        var tag = line[1..endTag].Split(' ')[0].TrimStart('/');
        var startClose = line.LastIndexOf("</");
        if (startClose < 0)
        {
            var val = line[(endTag + 1)..];
            if (val.EndsWith("/>")) val = val[..^2];
            return (tag, val.Trim());
        }
        var value = line[(endTag + 1)..startClose];
        return (tag, value);
    }

    private static async Task GenerateHexDiffAsync(byte[] db, byte[] gen, string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Hex Byte Diff");
        sb.AppendLine($"| Offset | Legacy (hex) | Generated (hex) | Legacy (ASCII) | Generated (ASCII) |");
        sb.AppendLine($"|--------|-------------|----------------|----------------|-------------------|");

        int maxLen = Math.Max(db.Length, gen.Length);
        int diffsFound = 0;
        const int maxDiffs = 500;

        for (int i = 0; i < maxLen && diffsFound < maxDiffs; i++)
        {
            byte dbB = i < db.Length ? db[i] : (byte)0;
            byte genB = i < gen.Length ? gen[i] : (byte)0;
            if (dbB == genB) continue;

            diffsFound++;
            var dbAscii = dbB >= 32 && dbB < 127 ? ((char)dbB).ToString() : ".";
            var genAscii = genB >= 32 && genB < 127 ? ((char)genB).ToString() : ".";
            sb.AppendLine($"| {i} | 0x{dbB:X2} | 0x{genB:X2} | {dbAscii} | {genAscii} |");
        }

        if (diffsFound >= maxDiffs)
            sb.AppendLine($"| ... | ({maxLen - diffsFound} more diffs) | | |");

        sb.AppendLine($"**Total differing bytes: {diffsFound}**");
        await File.WriteAllTextAsync(path, sb.ToString());
    }

    private static async Task GenerateLineDiffAsync(string db, string gen, string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Line-by-Line Diff");
        sb.AppendLine("| Line | Legacy | Generated | Length Diff |");
        sb.AppendLine("|------|--------|-----------|-------------|");

        var dbLines = db.Split('\n');
        var genLines = gen.Split('\n');
        int maxLines = Math.Max(dbLines.Length, genLines.Length);
        int diffsFound = 0;
        const int maxDiffs = 200;

        for (int i = 0; i < maxLines && diffsFound < maxDiffs; i++)
        {
            var dbL = i < dbLines.Length ? dbLines[i].TrimEnd('\r') : "";
            var genL = i < genLines.Length ? genLines[i].TrimEnd('\r') : "";
            if (dbL == genL) continue;

            diffsFound++;
            var lenDiff = dbL.Length - genL.Length;
            var dbDisplay = dbL.Length > 80 ? dbL[..77] + "..." : dbL;
            var genDisplay = genL.Length > 80 ? genL[..77] + "..." : genL;
            sb.AppendLine($"| {i + 1} | `{EscapeMd(dbDisplay)}` | `{EscapeMd(genDisplay)}` | {lenDiff} |");
        }

        if (diffsFound >= maxDiffs)
            sb.AppendLine($"| ... | ({maxLines - diffsFound} more diffs) | | |");

        sb.AppendLine($"**Total differing lines: {diffsFound}**");
        await File.WriteAllTextAsync(path, sb.ToString());
    }

    private static string EscapeMd(string s)
    {
        return s.Replace("|", "\\|").Replace("`", "'");
    }
}
