using System;
using Microsoft.Win32;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Keeps Outlook from disabling Jameel Copilot for a slow start. Outlook times every
    /// add-in while it starts and disables one that takes over a second; most
    /// of Jameel Copilot's time there is the .NET and VSTO runtimes loading, which Jameel Copilot
    /// cannot shorten. So Jameel Copilot does what the user's "Always enable this add-in"
    /// does: it names itself in the user's DoNotDisableAddinList. A company
    /// policy (the "List of managed add-ins") still decides over this list.
    /// </summary>
    internal static class AddinResiliency
    {
        public const string ProgId = "JameelCopilot";

        // The value "Always enable this add-in" writes: disabled once for a slow boot load.
        private const int BootLoad = 1;

        /// <summary>
        /// Names Jameel Copilot in the list when it is not there yet. Never throws: on any
        /// failure Jameel Copilot is simply not on the list.
        /// </summary>
        /// <param name="officeVersion">Outlook's version, such as "16.0.17928.20114".</param>
        /// <param name="root">The user's registry hive; tests pass a key of their own.</param>
        /// <returns>True when Jameel Copilot is on the list afterwards.</returns>
        public static bool KeepEnabled(string officeVersion, RegistryKey root = null)
        {
            try
            {
                var major = MajorVersion(officeVersion);
                if (major == null) return false;
                root = root ?? Registry.CurrentUser;
                using (var list = root.CreateSubKey(@"Software\Microsoft\Office\" + major + @"\Outlook\Resiliency\DoNotDisableAddinList"))
                {
                    if (list == null) return false;
                    if (list.GetValue(ProgId) == null) list.SetValue(ProgId, BootLoad, RegistryValueKind.DWord);
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>"16.0" from "16.0.17928.20114"; null for anything that is not a version.</summary>
        internal static string MajorVersion(string officeVersion)
        {
            var parts = (officeVersion ?? "").Split('.');
            return parts.Length >= 2 && int.TryParse(parts[0], out var major) && major >= 15 && int.TryParse(parts[1], out var minor)
                ? major + "." + minor
                : null;
        }
    }
}
