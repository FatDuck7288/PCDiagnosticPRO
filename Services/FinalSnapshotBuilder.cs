using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using VirtualIT.HardwareProbe;

namespace PCDiagnosticPro.Services
{
    public class FinalSnapshotBuilder
    {
        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public async Task<string> BuildAsync(
            string sourceJsonPath,
            HardwareProbeSnapshot hardwareSnapshot,
            string outputDirectory,
            string? runId = null)
        {
            var sourceJson = await File.ReadAllTextAsync(sourceJsonPath, Encoding.UTF8);
            JsonNode? rootNode = JsonNode.Parse(sourceJson);

            if (rootNode is not JsonObject rootObject)
            {
                rootObject = new JsonObject();
            }

            rootObject["hardwareProbe"] = JsonNode.Parse(JsonSerializer.Serialize(hardwareSnapshot, SerializerOptions));

            var fileName = string.IsNullOrWhiteSpace(runId)
                ? $"Snapshot_Final_{DateTime.UtcNow:yyyyMMdd_HHmmss}.json"
                : $"Snapshot_Final_{runId}.json";

            var outputPath = Path.Combine(outputDirectory, fileName);
            var json = rootObject.ToJsonString(SerializerOptions);
            await File.WriteAllTextAsync(outputPath, json, new UTF8Encoding(false));

            return outputPath;
        }
    }
}
