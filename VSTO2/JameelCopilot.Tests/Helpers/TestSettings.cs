using System.Collections.Generic;
using JameelCopilot.Services;

namespace JameelCopilot.Tests.Helpers
{
    /// <summary>Machine settings for tests, built the way the add-in reads them.</summary>
    internal static class TestSettings
    {
        public const string BaseUrl = "https://llm.example.internal/v1";
        public const string Key = "fake-key-for-tests";

        public static MachineSettings Configured(
            string model = "m1", string effort = null, string writeTools = "", int? maxRows = null,
            string temperature = null, string temperatureMax = null, string maxOutputTokens = null)
        {
            var values = new Dictionary<string, object>
            {
                ["ApiBaseUrl"] = BaseUrl,
                ["ApiKeyProtected"] = "BLOB",
                ["Model"] = model,
                ["ReasoningEffort"] = effort,
                ["EnabledWriteTools"] = writeTools,
                ["Temperature"] = temperature,
                ["TemperatureMax"] = temperatureMax,
                ["MaxOutputTokens"] = maxOutputTokens,
            };
            if (maxRows.HasValue) values["MaxBulkExportRows"] = maxRows.Value;
            return MachineSettings.Read(
                name => values.TryGetValue(name, out var value) ? value : null,
                blob => blob == "BLOB" ? Key : null);
        }

        public static MachineSettings Unconfigured()
        {
            return MachineSettings.Read(name => null, blob => null);
        }
    }
}
