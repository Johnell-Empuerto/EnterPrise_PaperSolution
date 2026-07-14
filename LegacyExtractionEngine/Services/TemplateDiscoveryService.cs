using Npgsql;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Discovers all templates from the database and maps them to workbook file paths.
/// Used by Phase 6 to classify every available template's coordinate strategy.
/// </summary>
public class TemplateDiscoveryService
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;

    public TemplateDiscoveryService(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
    }

    public async Task<List<TemplateInfo>> DiscoverAllTemplatesAsync()
    {
        var templates = new List<TemplateInfo>();

        _progress.Report("=== Discovering all templates from database ===");

        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        // Get all def_top IDs and names
        var sql = @"
            SELECT dt.def_top_id, dt.def_top_name, dt.def_top_org,
                   dts.cluster_count, dts.sheet_count,
                   dts.excel_size
            FROM def_top dt
            LEFT JOIN def_top_size dts ON dt.def_top_id = dts.def_top_id
            WHERE dt.def_file IS NOT NULL
               OR EXISTS (SELECT 1 FROM def_sheet ds WHERE ds.def_top_id = dt.def_top_id)
            ORDER BY dt.def_top_id";

        await using var cmd = new NpgsqlCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var info = new TemplateInfo
            {
                DefTopId = reader.GetInt32(0),
                DefTopName = reader.IsDBNull(1) ? null : reader.GetString(1),
                DefTopOrg = reader.IsDBNull(2) ? null : reader.GetInt32(2),
                ClusterCount = reader.IsDBNull(3) ? 0 : Convert.ToInt64(reader[3]),
                SheetCount = reader.IsDBNull(4) ? 1 : Convert.ToInt64(reader[4]),
                ExcelSize = reader.IsDBNull(5) ? 0 : Convert.ToInt64(reader[5]),
            };

            templates.Add(info);
        }

        _progress.Report($"Found {templates.Count} templates in database");

        // Now try to find workbook files for each template
        foreach (var t in templates)
        {
            t.WorkbookPath = FindWorkbookForTemplate(t.DefTopId, t.DefTopName);
        }

        // Also check if there are templates with clusters but no def_file
        var extraSql = @"
            SELECT dt.def_top_id, dt.def_top_name
            FROM def_top dt
            JOIN def_sheet ds ON dt.def_top_id = ds.def_top_id
            JOIN def_cluster dc ON ds.def_sheet_id = dc.def_sheet_id
            WHERE dt.def_file IS NULL
            GROUP BY dt.def_top_id, dt.def_top_name
            HAVING COUNT(dc.cluster_id) > 0
            ORDER BY dt.def_top_id";

        try
        {
            await using var extraCmd = new NpgsqlCommand(extraSql, conn);
            await using var extraReader = await extraCmd.ExecuteReaderAsync();
            while (await extraReader.ReadAsync())
            {
                var id = extraReader.GetInt32(0);
                var name = extraReader.IsDBNull(1) ? null : extraReader.GetString(1);
                if (!templates.Any(t => t.DefTopId == id))
                {
                    templates.Add(new TemplateInfo
                    {
                        DefTopId = id,
                        DefTopName = name,
                        WorkbookPath = FindWorkbookForTemplate(id, name)
                    });
                }
            }
        }
        catch { }

        return templates.OrderBy(t => t.DefTopId).ToList();
    }

    /// <summary>
    /// Find the workbook file for a given template ID.
    /// Checks known locations for templates 546 and 547.
    /// Does NOT search broadly — workbooks come from def_file in DB, not disk.
    /// </summary>
    public string? FindWorkbookForTemplate(int defTopId, string? defTopName)
    {
        var docs = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents";

        // Template 546: FormTest workbook
        if (defTopId == 546)
        {
            var p1 = Path.Combine(docs, "FormTest - Copy.xlsx");
            if (File.Exists(p1)) return p1;
            var p2 = Path.Combine(docs, "old_form.xlsx");
            if (File.Exists(p2)) return p2;
            return null;
        }

        // Template 547: questionnaire workbook
        if (defTopId == 547)
        {
            var p1 = @"C:\Users\MCF-JOHNELLEEMPUERTO\AppData\Local\Temp\opencode\547_workbook_copy.xlsx";
            if (File.Exists(p1)) return p1;
            try
            {
                var files = Directory.GetFiles(docs, "*V3.1*", SearchOption.TopDirectoryOnly);
                if (files.Length > 0) return files[0];
            }
            catch { }
            return null;
        }

        return null;
    }

    /// <summary>
    /// Get the def_top XML data to extract formInfo/printArea.
    /// </summary>
    public string? GetDefTopXml(int defTopId)
    {
        try
        {
            using var conn = new NpgsqlConnection(_connectionString);
            conn.Open();
            var cmd = new NpgsqlCommand("SELECT xml_data FROM def_top WHERE def_top_id = @id", conn);
            cmd.Parameters.AddWithValue("id", defTopId);
            return cmd.ExecuteScalar()?.ToString();
        }
        catch { return null; }
    }
}

public class TemplateInfo
{
    public int DefTopId { get; set; }
    public string? DefTopName { get; set; }
    public int? DefTopOrg { get; set; }
    public long ClusterCount { get; set; }
    public long SheetCount { get; set; }
    public long ExcelSize { get; set; }
    public string? WorkbookPath { get; set; }

    public bool HasWorkbook => !string.IsNullOrEmpty(WorkbookPath) && File.Exists(WorkbookPath);
    public bool HasClusters => ClusterCount > 0;

    public override string ToString()
    {
        var wb = HasWorkbook ? $"✓ {Path.GetFileName(WorkbookPath)}" : "✗ No workbook";
        return $"#{DefTopId} \"{DefTopName ?? "?"}\" {ClusterCount} clusters, {wb}";
    }
}
