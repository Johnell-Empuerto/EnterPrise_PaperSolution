using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Serializes RuntimeForm to JSON for consumption by the Next.js frontend.
    /// Provides optimized formatting with camelCase naming and null-value skipping.
    /// </summary>
    public class RuntimeSerializer
    {
        private static readonly JsonSerializerOptions _options = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// Serialize a RuntimeForm to a JSON string.
        /// </summary>
        public string Serialize(RuntimeForm form)
        {
            return JsonSerializer.Serialize(form, _options);
        }

        /// <summary>
        /// Serialize a RuntimeForm to a JSON file.
        /// </summary>
        public void SerializeToFile(RuntimeForm form, string outputPath)
        {
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var json = Serialize(form);
            System.IO.File.WriteAllText(outputPath, json);
        }

        /// <summary>
        /// Deserialize a RuntimeForm from a JSON string.
        /// </summary>
        public RuntimeForm? Deserialize(string json)
        {
            return JsonSerializer.Deserialize<RuntimeForm>(json, _options);
        }

        /// <summary>
        /// Get the JSON serializer options for external use.
        /// </summary>
        public static JsonSerializerOptions Options => _options;

        /// <summary>
        /// Serialize to a minimal JSON string (no indentation).
        /// </summary>
        public string SerializeMinified(RuntimeForm form)
        {
            var minOptions = new JsonSerializerOptions
            {
                WriteIndented = false,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };

            return JsonSerializer.Serialize(form, minOptions);
        }
    }
}
