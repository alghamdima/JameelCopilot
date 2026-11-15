using System;
using System.Collections.Generic;
using JameelCopilot.Services;

namespace JameelCopilot
{
    /// <summary>
    /// The settings in effect. They come from HKLM\SOFTWARE\JameelCopilot only
    /// (see <see cref="MachineSettings"/>): an administrator sets them through
    /// the installer, and nothing in the add-in can change them.
    /// </summary>
    public static class Config
    {
        /// <summary>The four safe-write tools an administrator can switch on.</summary>
        public static readonly string[] AllWriteTools = MachineSettings.AllWriteTools;

        private static MachineSettings _current;

        // Read on first use, so a test can apply its own settings first.
        internal static MachineSettings Current => _current ?? (_current = MachineSettings.ReadFromRegistry());

        public static Uri ApiBaseUrl => Current.ApiBaseUrl;
        public static string ApiKey => Current.ApiKey;
        public static string Model => Current.Model;

        /// <summary>Sent as <c>reasoning_effort</c> when set; null means the model's default.</summary>
        public static string ReasoningEffort => Current.ReasoningEffort;

        public static HashSet<string> EnabledWriteTools => Current.EnabledWriteTools;
        public static bool WriteToolsEnabled => Current.EnabledWriteTools.Count > 0;
        public static int MaxBulkExportRows => Current.MaxBulkExportRows;

        /// <summary>The "Balanced" temperature; null means none is sent and users get no style choice.</summary>
        public static double? Temperature => Current.Temperature;

        /// <summary>The "Creative" temperature: the highest Jameel Copilot ever sends.</summary>
        public static double TemperatureMax => Current.TemperatureMax;

        /// <summary>Sent as <c>max_tokens</c> when set.</summary>
        public static int? MaxOutputTokens => Current.MaxOutputTokens;

        public static bool IsConfigured => Current.IsConfigured;

        /// <summary>Why Jameel Copilot is not configured, or null when it is. Never contains the key.</summary>
        public static string Problem => Current.Problem;

        /// <summary>Reads the settings from the registry again.</summary>
        public static void Load()
        {
            _current = MachineSettings.ReadFromRegistry();
        }

        // Test seam.
        internal static void Apply(MachineSettings settings)
        {
            _current = settings;
        }
    }
}
