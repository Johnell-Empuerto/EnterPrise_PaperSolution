using Npgsql;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class DatabaseReader
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;

    public DatabaseReader(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
    }

    public async Task<DatabaseDump> ReadAllAsync(int defTopId)
    {
        _progress.Report("[1/15] Connecting PostgreSQL...");
        var dump = new DatabaseDump { ExtractedAt = DateTime.UtcNow, Source = "database" };

        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        _progress.Report("[2/15] Loading def_top...");
        dump.DefTop = await ReadDefTopAsync(conn, defTopId);

        _progress.Report("[3/15] Loading def_sheet...");
        dump.DefSheets = await ReadDefSheetsAsync(conn, defTopId);

        _progress.Report("[4/15] Loading def_cluster...");
        foreach (var sheet in dump.DefSheets)
        {
            sheet.Clusters = await ReadDefClustersAsync(conn, sheet.DefSheetId);
        }

        _progress.Report("[5/15] Loading def_top_size...");
        dump.DefTopSize = await ReadDefTopSizeAsync(conn, defTopId);

        _progress.Report("[6/15] Loading related tables...");
        dump.DefTopOptions = await ReadTableAsync<DefTopOption>(conn, "def_top_option", "def_top_id", defTopId);
        dump.DefLabels = await ReadTableAsync<DefLabel>(conn, "def_label", "def_top_id", defTopId);

        _progress.Report("[7/15] Loading all foreign key related tables...");
        var relatedTables = await DiscoverAndReadForeignKeysAsync(conn, "def_top", defTopId);
        foreach (var (tableName, rows) in relatedTables)
        {
            if (!dump.AllRelatedTables.ContainsKey(tableName))
                dump.AllRelatedTables[tableName] = rows;
        }

        foreach (var sheet in dump.DefSheets)
        {
            var sheetRelated = await DiscoverAndReadForeignKeysAsync(conn, "def_sheet", sheet.DefSheetId);
            foreach (var (tableName, rows) in sheetRelated)
            {
                if (!dump.AllRelatedTables.ContainsKey(tableName))
                    dump.AllRelatedTables[tableName] = rows;
            }
        }

        return dump;
    }

    private async Task<DefTop?> ReadDefTopAsync(NpgsqlConnection conn, int defTopId)
    {
        var columns = await GetColumnNamesAsync(conn, "def_top");
        var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM def_top WHERE def_top_id = @id";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("id", defTopId);
        await using var reader = await cmd.ExecuteReaderAsync();

        if (!await reader.ReadAsync()) return null;

        var defTop = new DefTop { DefTopId = defTopId };
        MapReaderToObject(reader, columns, defTop);
        return defTop;
    }

    private async Task<List<DefSheet>> ReadDefSheetsAsync(NpgsqlConnection conn, int defTopId)
    {
        var sheets = new List<DefSheet>();
        var columns = await GetColumnNamesAsync(conn, "def_sheet");
        var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM def_sheet WHERE def_top_id = @id ORDER BY def_sheet_no";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("id", defTopId);
        await using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var sheet = new DefSheet();
            MapReaderToObject(reader, columns, sheet);
            sheets.Add(sheet);
        }
        return sheets;
    }

    private async Task<List<DefCluster>> ReadDefClustersAsync(NpgsqlConnection conn, int defSheetId)
    {
        var clusters = new List<DefCluster>();
        var columns = await GetColumnNamesAsync(conn, "def_cluster");
        var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM def_cluster WHERE def_sheet_id = @id ORDER BY cluster_id";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("id", defSheetId);
        await using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var cluster = new DefCluster { DefSheetId = defSheetId };
            MapReaderToObject(reader, columns, cluster);
            clusters.Add(cluster);
        }
        return clusters;
    }

    private async Task<DefTopSize?> ReadDefTopSizeAsync(NpgsqlConnection conn, int defTopId)
    {
        var columns = await GetColumnNamesAsync(conn, "def_top_size");
        var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM def_top_size WHERE def_top_id = @id";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("id", defTopId);
        await using var reader = await cmd.ExecuteReaderAsync();

        if (!await reader.ReadAsync()) return null;

        var size = new DefTopSize { DefTopId = defTopId };
        for (int i = 0; i < columns.Count; i++)
        {
            var val = reader.IsDBNull(i) ? null : reader[i];
            switch (columns[i])
            {
                case "excel_size": size.ExcelSize = val as long?; break;
                case "pdf_size": size.PdfSize = val as long?; break;
                case "sheet_count": size.SheetCount = val as long?; break;
                case "cluster_count": size.ClusterCount = val as long?; break;
                case "updated_at": size.UpdatedAt = val as DateTime?; break;
            }
        }
        return size;
    }

    private async Task<List<T>> ReadTableAsync<T>(NpgsqlConnection conn, string table, string fkColumn, int fkValue) where T : new()
    {
        var result = new List<T>();
        try
        {
            var columns = await GetColumnNamesAsync(conn, table);
            var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM {table} WHERE \"{fkColumn}\" = @id";
            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("id", fkValue);
            await using var reader = await cmd.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var item = new T();
                if (item is DefTopOption opt)
                {
                    for (int i = 0; i < columns.Count; i++)
                    {
                        var val = reader.IsDBNull(i) ? null : reader[i];
                        if (columns[i] == "def_top_id") opt.DefTopId = Convert.ToInt32(val);
                        else opt.AllColumns[columns[i]] = val;
                    }
                }
                else if (item is DefLabel label)
                {
                    for (int i = 0; i < columns.Count; i++)
                    {
                        var val = reader.IsDBNull(i) ? null : reader[i];
                        if (columns[i] == "def_label_id") label.DefLabelId = Convert.ToInt32(val);
                        else if (columns[i] == "def_top_id") label.DefTopId = Convert.ToInt32(val);
                        else label.AllColumns[columns[i]] = val;
                    }
                }
                result.Add(item);
            }
        }
        catch (Exception ex)
        {
            _progress.Report($"  Warning: Could not read {table}: {ex.Message}");
        }
        return result;
    }

    private async Task<Dictionary<string, List<Dictionary<string, object?>>>> DiscoverAndReadForeignKeysAsync(NpgsqlConnection conn, string sourceTable, int sourceId)
    {
        var result = new Dictionary<string, List<Dictionary<string, object?>>>();

        try
        {
            var fkSql = @"
                SELECT
                    tc.table_name,
                    kcu.column_name
                FROM information_schema.table_constraints tc
                JOIN information_schema.key_column_usage kcu
                    ON tc.constraint_name = kcu.constraint_name
                    AND tc.table_schema = kcu.table_schema
                JOIN information_schema.constraint_column_usage ccu
                    ON ccu.constraint_name = tc.constraint_name
                    AND ccu.table_schema = tc.table_schema
                WHERE tc.constraint_type = 'FOREIGN KEY'
                    AND ccu.table_name = @tableName
                    AND ccu.column_name = 'def_top_id'";

            await using var fkCmd = new NpgsqlCommand(fkSql, conn);
            fkCmd.Parameters.AddWithValue("tableName", sourceTable);
            await using var fkReader = await fkCmd.ExecuteReaderAsync();

            var foreignTables = new List<(string table, string column)>();
            while (await fkReader.ReadAsync())
            {
                foreignTables.Add((fkReader.GetString(0), fkReader.GetString(1)));
            }
            await fkReader.CloseAsync();

            foreach (var (table, column) in foreignTables)
            {
                try
                {
                    var columns = await GetColumnNamesAsync(conn, table);
                    var sql = $"SELECT {string.Join(", ", columns.Select(c => $"\"{c}\""))} FROM {table} WHERE \"{column}\" = @id";
                    await using var dataCmd = new NpgsqlCommand(sql, conn);
                    dataCmd.Parameters.AddWithValue("id", sourceId);
                    await using var dataReader = await dataCmd.ExecuteReaderAsync();

                    var rows = new List<Dictionary<string, object?>>();
                    while (await dataReader.ReadAsync())
                    {
                        var row = new Dictionary<string, object?>();
                        for (int i = 0; i < columns.Count; i++)
                        {
                            row[columns[i]] = dataReader.IsDBNull(i) ? null : dataReader[i];
                        }
                        rows.Add(row);
                    }

                    if (rows.Count > 0)
                        result[table] = rows;

                    await dataReader.CloseAsync();
                }
                catch (Exception ex)
                {
                    _progress.Report($"  Warning: Could not read {table}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            _progress.Report($"  Warning: FK discovery failed for {sourceTable}: {ex.Message}");
        }

        return result;
    }

    private async Task<List<string>> GetColumnNamesAsync(NpgsqlConnection conn, string table)
    {
        var columns = new List<string>();
        var sql = "SELECT column_name FROM information_schema.columns WHERE table_name = @table ORDER BY ordinal_position";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("table", table);
        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            columns.Add(reader.GetString(0));
        }
        return columns;
    }

    private void MapReaderToObject(NpgsqlDataReader reader, List<string> columns, object target)
    {
        var targetType = target.GetType();

        for (int i = 0; i < columns.Count; i++)
        {
            var val = reader.IsDBNull(i) ? null : reader[i];
            var propName = columns[i] switch
            {
                "def_top_id" when target is DefTop => nameof(DefTop.DefTopId),
                "def_top_name" when target is DefTop => nameof(DefTop.DefTopName),
                "xml_data" when target is DefTop => nameof(DefTop.XmlData),
                "thumbnail" when target is DefTop => nameof(DefTop.Thumbnail),
                "background_image_file" when target is DefTop => nameof(DefTop.BackgroundImageFile),
                "designer_version" when target is DefTop => nameof(DefTop.DesignerVersion),
                "server_version" when target is DefTop => nameof(DefTop.ServerVersion),
                "def_sheet_id" when target is DefSheet => nameof(DefSheet.DefSheetId),
                "def_sheet_no" when target is DefSheet => nameof(DefSheet.DefSheetNo),
                "def_sheet_name" when target is DefSheet => nameof(DefSheet.DefSheetName),
                "def_top_id" when target is DefSheet => nameof(DefSheet.DefTopId),
                "def_sheet_id" when target is DefSheet => nameof(DefSheet.DefSheetId),
                "cluster_id" when target is DefCluster => nameof(DefCluster.ClusterId),
                "cluster_name" when target is DefCluster => nameof(DefCluster.ClusterName),
                "cluster_type" when target is DefCluster => nameof(DefCluster.ClusterType),
                "left_position" when target is DefCluster => nameof(DefCluster.LeftPosition),
                "right_position" when target is DefCluster => nameof(DefCluster.RightPosition),
                "top_position" when target is DefCluster => nameof(DefCluster.TopPosition),
                "bottom_position" when target is DefCluster => nameof(DefCluster.BottomPosition),
                "cell_addr" when target is DefCluster => nameof(DefCluster.CellAddr),
                "input_parameter" when target is DefCluster => nameof(DefCluster.InputParameter),
                _ => null
            };

            if (propName != null)
            {
                var prop = targetType.GetProperty(propName);
                if (prop != null && val != null)
                {
                    try
                    {
                        if (prop.PropertyType == typeof(int) || prop.PropertyType == typeof(int?))
                            prop.SetValue(target, Convert.ToInt32(val));
                        else if (prop.PropertyType == typeof(decimal) || prop.PropertyType == typeof(decimal?))
                            prop.SetValue(target, Convert.ToDecimal(val));
                        else if (prop.PropertyType == typeof(DateTime) || prop.PropertyType == typeof(DateTime?))
                            prop.SetValue(target, Convert.ToDateTime(val));
                        else if (prop.PropertyType == typeof(byte[]))
                            prop.SetValue(target, val is byte[] ? val : null);
                        else
                            prop.SetValue(target, val is string s ? s : val?.ToString());
                    }
                    catch { /* type conversion issue, skip */ }
                }
            }

            if (target is DefTop dt)
                dt.AllColumns[columns[i]] = val;
            else if (target is DefSheet ds)
                ds.AllColumns[columns[i]] = val;
            else if (target is DefCluster dc)
                dc.AllColumns[columns[i]] = val;
        }
    }
}
