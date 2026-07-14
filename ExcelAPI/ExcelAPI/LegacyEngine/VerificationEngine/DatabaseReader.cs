using Npgsql;

namespace ExcelAPI.LegacyEngine.VerificationEngine;

public class VerificationDatabaseReader : IVerificationDatabaseReader
{
    private readonly string _connectionString;
    private readonly ILogger<VerificationDatabaseReader> _logger;

    public VerificationDatabaseReader(IConfiguration configuration, ILogger<VerificationDatabaseReader> logger)
    {
        _connectionString = configuration.GetConnectionString("IrepoDb")
            ?? "Host=localhost;Port=5432;Database=irepodb;Username=postgres;Password=cimtops";
        _logger = logger;
    }

    public async Task<DefTopData?> ReadDefTopAsync(int defTopId)
    {
        try
        {
            await using var conn = new NpgsqlConnection(_connectionString);
            await conn.OpenAsync();

            await using var cmd = new NpgsqlCommand(
                "SELECT def_top_id, xml_data, background_image_file FROM def_top WHERE def_top_id = @id", conn);
            cmd.Parameters.AddWithValue("id", defTopId);

            await using var reader = await cmd.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            return new DefTopData
            {
                DefTopId = defTopId,
                XmlData = reader.IsDBNull(1) ? null : reader.GetString(1),
                BackgroundImageData = reader.IsDBNull(2) ? null : (byte[])reader[2]
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to read def_top for {DefTopId}", defTopId);
            return null;
        }
    }

    public async Task<List<DefClusterData>> ReadDefClustersAsync(int defTopId)
    {
        var result = new List<DefClusterData>();
        try
        {
            await using var conn = new NpgsqlConnection(_connectionString);
            await conn.OpenAsync();

            await using var cmd = new NpgsqlCommand(@"
                SELECT c.def_sheet_id, c.cluster_id, c.cluster_name, c.cluster_type,
                       c.left_position, c.right_position, c.top_position, c.bottom_position,
                       c.cell_addr, c.input_parameter
                FROM def_cluster c
                JOIN def_sheet s ON c.def_sheet_id = s.def_sheet_id
                WHERE s.def_top_id = @id
                ORDER BY c.def_sheet_id, c.cluster_id", conn);
            cmd.Parameters.AddWithValue("id", defTopId);

            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                result.Add(new DefClusterData
                {
                    DefSheetId = reader.GetInt32(0),
                    ClusterId = reader.GetInt32(1),
                    ClusterName = reader.IsDBNull(2) ? null : reader.GetString(2),
                    ClusterType = reader.IsDBNull(3) ? null : reader.GetString(3),
                    LeftPosition = reader.IsDBNull(4) ? null : reader.GetString(4),
                    RightPosition = reader.IsDBNull(5) ? null : reader.GetString(5),
                    TopPosition = reader.IsDBNull(6) ? null : reader.GetString(6),
                    BottomPosition = reader.IsDBNull(7) ? null : reader.GetString(7),
                    CellAddr = reader.IsDBNull(8) ? null : reader.GetString(8),
                    InputParameter = reader.IsDBNull(9) ? null : reader.GetString(9)
                });
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to read def_cluster for {DefTopId}", defTopId);
        }
        return result;
    }

    public async Task<byte[]?> ReadBackgroundImageAsync(int defTopId)
    {
        try
        {
            await using var conn = new NpgsqlConnection(_connectionString);
            await conn.OpenAsync();

            await using var cmd = new NpgsqlCommand(
                "SELECT background_image_file FROM def_top WHERE def_top_id = @id", conn);
            cmd.Parameters.AddWithValue("id", defTopId);

            var result = await cmd.ExecuteScalarAsync();
            return result as byte[];
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to read background image for {DefTopId}", defTopId);
            return null;
        }
    }
}
