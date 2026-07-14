using System.Text;
using System.Text.Json;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class DeepDiffEngine
{
    private readonly IProgress<string> _progress;

    public DeepDiffEngine(IProgress<string> progress)
    {
        _progress = progress;
    }

    public async Task RunDeepDiffAsync(DatabaseDump database, ExtractionResult generated, string outputDir)
    {
        _progress.Report("[13/15] Running deep diff...");

        // Serialize both to JSON for comparison
        var opts = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        var dbJson = JsonSerializer.Serialize(database, opts);
        var genJson = JsonSerializer.Serialize(generated, opts);

        // Write dumps
        await File.WriteAllTextAsync(Path.Combine(outputDir, "database_dump.json"), dbJson);
        await File.WriteAllTextAsync(Path.Combine(outputDir, "generated_dump.json"), genJson);
        _progress.Report("  Wrote database_dump.json and generated_dump.json");

        // Parse for comparison
        using var dbDoc = JsonDocument.Parse(dbJson);
        using var genDoc = JsonDocument.Parse(genJson);

        var diffs = new List<DeepDiffResult>();
        CompareElements(dbDoc.RootElement, genDoc.RootElement, "$", diffs);

        var deepDiffResult = new
        {
            def_top_id = database.DefTop?.DefTopId ?? 0,
            generated_at = DateTime.UtcNow,
            total_diffs = diffs.Count,
            equal_count = diffs.Count(d => d.DiffType == "EQUAL"),
            modified_count = diffs.Count(d => d.DiffType == "MODIFIED"),
            added_count = diffs.Count(d => d.DiffType == "ADDED"),
            removed_count = diffs.Count(d => d.DiffType == "REMOVED"),
            differences = diffs
        };

        await File.WriteAllTextAsync(Path.Combine(outputDir, "deep_diff.json"),
            JsonSerializer.Serialize(deepDiffResult, opts));
        _progress.Report($"  Wrote deep_diff.json ({diffs.Count} diffs found)");

        // Write summary
        await WriteDeepDiffSummaryAsync(deepDiffResult, outputDir);
    }

    private void CompareElements(JsonElement db, JsonElement gen, string path, List<DeepDiffResult> results)
    {
        if (db.ValueKind != gen.ValueKind)
        {
            results.Add(new DeepDiffResult
            {
                DiffType = "TYPE_MISMATCH",
                Path = path,
                DatabaseValue = GetValueString(db),
                GeneratedValue = GetValueString(gen),
                DatabaseType = db.ValueKind.ToString(),
                GeneratedType = gen.ValueKind.ToString(),
                Reason = $"Type mismatch: {db.ValueKind} vs {gen.ValueKind}"
            });
            return;
        }

        switch (db.ValueKind)
        {
            case JsonValueKind.Object:
                CompareObjects(db, gen, path, results);
                break;
            case JsonValueKind.Array:
                CompareArrays(db, gen, path, results);
                break;
            case JsonValueKind.String:
            case JsonValueKind.Number:
            case JsonValueKind.True:
            case JsonValueKind.False:
            case JsonValueKind.Null:
                CompareValues(db, gen, path, results);
                break;
        }
    }

    private void CompareObjects(JsonElement db, JsonElement gen, string path, List<DeepDiffResult> results)
    {
        var dbProps = new HashSet<string>();
        var genProps = new HashSet<string>();

        foreach (var prop in db.EnumerateObject())
        {
            dbProps.Add(prop.Name);
            if (gen.TryGetProperty(prop.Name, out var genProp))
            {
                CompareElements(prop.Value, genProp, $"{path}.{prop.Name}", results);
            }
            else
            {
                results.Add(new DeepDiffResult
                {
                    DiffType = "REMOVED",
                    Path = $"{path}.{prop.Name}",
                    DatabaseValue = GetValueString(prop.Value),
                    Reason = $"Property '{prop.Name}' exists in database but not in generated output"
                });
            }
        }

        foreach (var prop in gen.EnumerateObject())
        {
            if (!dbProps.Contains(prop.Name))
            {
                results.Add(new DeepDiffResult
                {
                    DiffType = "ADDED",
                    Path = $"{path}.{prop.Name}",
                    GeneratedValue = GetValueString(prop.Value),
                    Reason = $"Property '{prop.Name}' exists in generated output but not in database"
                });
            }
        }
    }

    private void CompareArrays(JsonElement db, JsonElement gen, string path, List<DeepDiffResult> results)
    {
        var dbLen = db.GetArrayLength();
        var genLen = gen.GetArrayLength();

        if (dbLen != genLen)
        {
            results.Add(new DeepDiffResult
            {
                DiffType = "MODIFIED",
                Path = path,
                DatabaseValue = $"[array length: {dbLen}]",
                GeneratedValue = $"[array length: {genLen}]",
                Reason = $"Array length differs: {dbLen} vs {genLen}"
            });
        }

        int minLen = Math.Min(dbLen, genLen);
        for (int i = 0; i < minLen; i++)
        {
            CompareElements(db[i], gen[i], $"{path}[{i}]", results);
        }

        if (dbLen > genLen)
        {
            for (int i = genLen; i < dbLen; i++)
            {
                results.Add(new DeepDiffResult
                {
                    DiffType = "REMOVED",
                    Path = $"{path}[{i}]",
                    DatabaseValue = GetValueString(db[i]),
                    Reason = $"Array element at index {i} removed in generated output"
                });
            }
        }
        else if (genLen > dbLen)
        {
            for (int i = dbLen; i < genLen; i++)
            {
                results.Add(new DeepDiffResult
                {
                    DiffType = "ADDED",
                    Path = $"{path}[{i}]",
                    GeneratedValue = GetValueString(gen[i]),
                    Reason = $"Array element at index {i} added in generated output"
                });
            }
        }
    }

    private void CompareValues(JsonElement db, JsonElement gen, string path, List<DeepDiffResult> results)
    {
        var dbStr = db.ValueKind == JsonValueKind.Number
            ? db.GetRawText()
            : db.ToString();
        var genStr = gen.ValueKind == JsonValueKind.Number
            ? gen.GetRawText()
            : gen.ToString();

        if (dbStr == genStr)
        {
            results.Add(new DeepDiffResult
            {
                DiffType = "EQUAL",
                Path = path,
                DatabaseValue = Truncate(dbStr, 100),
                GeneratedValue = Truncate(genStr, 100)
            });
        }
        else
        {
            results.Add(new DeepDiffResult
            {
                DiffType = "MODIFIED",
                Path = path,
                DatabaseValue = Truncate(dbStr, 100),
                GeneratedValue = Truncate(genStr, 100),
                Reason = db.ValueKind == JsonValueKind.Number && gen.ValueKind == JsonValueKind.Number
                    ? "Numeric value differs"
                    : "String value differs"
            });
        }
    }

    private string GetValueString(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? "",
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "null",
            JsonValueKind.Object => "{object}",
            JsonValueKind.Array => "[array]",
            _ => ""
        };
    }

    private string Truncate(string s, int maxLen)
    {
        return s?.Length > maxLen ? s[..maxLen] + "..." : s ?? "";
    }

    private async Task WriteDeepDiffSummaryAsync(object deepDiffResult, string outputDir)
    {
        var result = (dynamic)deepDiffResult;
        var md = new StringBuilder();
        md.AppendLine("# Deep Diff Summary");
        md.AppendLine();
        md.AppendLine("## Overview");
        md.AppendLine();
        md.AppendLine($"| Metric | Value |");
        md.AppendLine($"|--------|-------|");
        md.AppendLine($"| Total Comparisons | {result.total_diffs} |");
        md.AppendLine($"| Equal | {result.equal_count} |");
        md.AppendLine($"| Modified | {result.modified_count} |");
        md.AppendLine($"| Added | {result.added_count} |");
        md.AppendLine($"| Removed | {result.removed_count} |");

        await File.WriteAllTextAsync(Path.Combine(outputDir, "debug", "deep_diff_summary.md"), md.ToString());
        _progress.Report("  Wrote deep_diff_summary.md");
    }
}
