using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Win32;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Jameel Copilot's settings, set by an administrator for every user of this PC
    /// (HKLM\SOFTWARE\JameelCopilot, written by the installer). Reading never
    /// throws: anything missing or invalid leaves Jameel Copilot "not configured", with
    /// <see cref="Problem"/> naming the setting.
    /// </summary>
    public sealed class MachineSettings
    {
        public const string RegistryPath = @"SOFTWARE\JameelCopilot";
        public const int DefaultMaxBulkExportRows = 2000;
        public const double DefaultTemperature = 0.5;
        public const double DefaultTemperatureMax = 0.9;
        public const double PreciseTemperature = 0.2;

        public static readonly string[] AllWriteTools =
        {
            "outlook_create_draft",
            "outlook_mark_as_read",
            "outlook_flag_message",
            "outlook_set_category"
        };

        private MachineSettings()
        {
            EnabledWriteTools = new HashSet<string>(StringComparer.Ordinal);
            MaxBulkExportRows = DefaultMaxBulkExportRows;
            Temperature = DefaultTemperature;
            TemperatureMax = DefaultTemperatureMax;
        }

        /// <summary>The Custom API base address (https, no trailing slash), or null.</summary>
        public Uri ApiBaseUrl { get; private set; }

        /// <summary>The decrypted API key, or null.</summary>
        public string ApiKey { get; private set; }

        public string Model { get; private set; }

        /// <summary>Sent as <c>reasoning_effort</c> when set; null otherwise.</summary>
        public string ReasoningEffort { get; private set; }

        /// <summary>The safe-write tools the administrator switched on. None by default.</summary>
        public HashSet<string> EnabledWriteTools { get; private set; }

        public int MaxBulkExportRows { get; private set; }

        /// <summary>
        /// The temperature of the "Balanced" response style, 0 to 2. Null
        /// (the value "off") sends none, for a model that rejects it.
        /// </summary>
        public double? Temperature { get; private set; }

        /// <summary>The "Creative" style's temperature, and the highest Jameel Copilot ever sends.</summary>
        public double TemperatureMax { get; private set; }

        /// <summary>Sent as <c>max_tokens</c> when set; null sends none.</summary>
        public int? MaxOutputTokens { get; private set; }

        public bool IsConfigured => Problem == null;

        /// <summary>Null when configured, else one sentence naming the setting. Never contains the key.</summary>
        public string Problem { get; private set; }

        /// <param name="read">Returns a setting's raw value by name, or null.</param>
        /// <param name="unprotect">Decrypts the stored key, or returns null.</param>
        public static MachineSettings Read(Func<string, object> read, Func<string, string> unprotect)
        {
            var settings = new MachineSettings();
            try
            {
                settings.ApiBaseUrl = ParseUrl(Text(read("ApiBaseUrl")));
                settings.ApiKey = Unprotect(unprotect, Text(read("ApiKeyProtected")));
                settings.Model = Text(read("Model"));
                settings.ReasoningEffort = Text(read("ReasoningEffort"));
                settings.EnabledWriteTools = ParseWriteTools(Text(read("EnabledWriteTools")));
                settings.MaxBulkExportRows = ParseRows(read("MaxBulkExportRows"));
                settings.Temperature = ParseTemperature(Text(read("Temperature")));
                settings.TemperatureMax = ParseTemperatureMax(Text(read("TemperatureMax")), settings.Temperature);
                settings.MaxOutputTokens = ParseMaxOutputTokens(read("MaxOutputTokens"));
            }
            catch (Exception ex)
            {
                // Whatever went wrong reading the registry, Outlook must still start.
                return new MachineSettings { Problem = "The Jameel Copilot settings could not be read (" + ex.GetType().Name + ")." };
            }

            settings.Problem =
                settings.ApiBaseUrl == null ? "ApiBaseUrl is missing or is not an https:// address."
                : settings.ApiKey == null ? "ApiKeyProtected is missing or cannot be read on this PC."
                : settings.Model == null ? "Model is missing."
                : null;
            return settings;
        }

        public static MachineSettings ReadFromRegistry()
        {
            return Read(
                name =>
                {
                    // The 64-bit view, where the installer writes, even from a 32-bit Outlook.
                    using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                    using (var key = hklm.OpenSubKey(RegistryPath))
                    {
                        return key?.GetValue(name);
                    }
                },
                ApiKeyProtector.Unprotect);
        }

        // A trimmed string value; null for blank or for any other registry type.
        private static string Text(object value)
        {
            var text = (value as string)?.Trim();
            return string.IsNullOrEmpty(text) ? null : text;
        }

        private static Uri ParseUrl(string text)
        {
            if (text == null) return null;
            return Uri.TryCreate(text.TrimEnd('/'), UriKind.Absolute, out var uri)
                && uri.Scheme == Uri.UriSchemeHttps
                && !string.IsNullOrEmpty(uri.Host)
                ? uri
                : null;
        }

        private static string Unprotect(Func<string, string> unprotect, string blob)
        {
            if (blob == null) return null;
            try
            {
                var key = unprotect(blob);
                return string.IsNullOrEmpty(key) ? null : key;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static HashSet<string> ParseWriteTools(string list)
        {
            return new HashSet<string>(
                (list ?? "").Split(',').Select(name => name.Trim()).Where(name => AllWriteTools.Contains(name, StringComparer.Ordinal)),
                StringComparer.Ordinal);
        }

        // Numbers are written with a dot whatever the PC's regional settings.
        private static double? ParseNumber(string text)
        {
            return text != null
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                && !double.IsNaN(number) && !double.IsInfinity(number)
                ? number
                : (double?)null;
        }

        private static double? ParseTemperature(string text)
        {
            if (string.Equals(text, "off", StringComparison.OrdinalIgnoreCase)) return null;
            return Math.Max(0, Math.Min(2, ParseNumber(text) ?? DefaultTemperature));
        }

        private static double ParseTemperatureMax(string text, double? temperature)
        {
            return Math.Max(temperature ?? 0, Math.Min(2, ParseNumber(text) ?? DefaultTemperatureMax));
        }

        // A whole number above zero, as a string or a registry number; anything else is "no limit".
        private static int? ParseMaxOutputTokens(object value)
        {
            long tokens;
            if (value is int) tokens = (int)value;
            else if (value is long) tokens = (long)value;
            else if (!(value is string) || !long.TryParse(((string)value).Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out tokens)) return null;
            return tokens <= 0 ? (int?)null : (int)Math.Min(100000, tokens);
        }

        // REG_DWORD, REG_QWORD or a numeric string; anything else is the default.
        private static int ParseRows(object value)
        {
            long rows;
            if (value is int) rows = (int)value;
            else if (value is long) rows = (long)value;
            else if (!(value is string) || !long.TryParse(((string)value).Trim(), out rows)) return DefaultMaxBulkExportRows;
            return (int)Math.Max(1, Math.Min(BulkExportRowCap.Max, rows));
        }
    }
}
