using System;
using System.IO;
using System.Text.RegularExpressions;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// Advice in the deployment guide that was wrong once. IT follows the
    /// guide to the letter, so a wrong line there is a failed rollout.
    /// </summary>
    public class DeployGuideTests
    {
        private static readonly string Guide =
            File.ReadAllText(RepoFiles.Find("Deploy", "README.txt")).Replace("\r\n", "\n");

        // A section runs from its title to the two blank lines before the next
        // one. Its text comes back on one line, so a sentence is found
        // wherever the guide happens to wrap it.
        private static string Section(string title)
        {
            var start = Guide.IndexOf("\n" + title + "\n", StringComparison.Ordinal);
            Assert.True(start >= 0, "No section: " + title);
            var end = Guide.IndexOf("\n\n\n", start + 1, StringComparison.Ordinal);
            return Regex.Replace(end < 0 ? Guide.Substring(start) : Guide.Substring(start, end - start), @"\s+", " ");
        }

        // Configuration Manager runs an Application's command only where
        // detection says "not installed": an edited command never reaches a
        // PC that has Jameel Copilot, so neither does a changed setting or a new key.
        [Fact]
        public void Sccm_ChangesAndTheKeyGoOutAsAPackage_NotByEditingTheApplication()
        {
            var sccm = Section("DEPLOYING WITH CONFIGURATION MANAGER (SCCM)");

            Assert.DoesNotContain("and redeploy it", sccm);
            Assert.DoesNotContain("redeploy the command", sccm);
            Assert.Contains("only on a PC where detection says Jameel Copilot is not installed", sccm);
            Assert.Contains("\"Always rerun program\"", sccm);
            Assert.Contains("msiexec /i JameelCopilot-<version>-x64.msi API_KEY=<new key> /qn", sccm);
        }

        // The upgrade's Application also lands on PCs without Jameel Copilot, and an
        // uninstall first would delete the settings the upgrade relies on.
        [Fact]
        public void Sccm_UpgradeKeepsTheSettingsInItsCommand_AndDoesNotUninstallFirst()
        {
            var sccm = Section("DEPLOYING WITH CONFIGURATION MANAGER (SCCM)");

            Assert.DoesNotContain("command needs no settings", sccm);
            Assert.Contains("\"Uninstall\" left unticked", sccm);
            Assert.Contains("Give it the full command, settings included", sccm);
        }

        // The MSI remembers a setting that is left out, so leaving
        // MAX_OUTPUT_TOKENS out does not remove the limit; 0 does.
        [Fact]
        public void RemovingTheAnswerLimit_IsMaxOutputTokensZero()
        {
            var troubleshooting = Section("TROUBLESHOOTING");
            Assert.DoesNotContain("remove MAX_OUTPUT_TOKENS", troubleshooting);
            Assert.Contains("MAX_OUTPUT_TOKENS=0", troubleshooting);

            var install = Section("INSTALL");
            Assert.Contains("install again with MAX_OUTPUT_TOKENS=0", install);
            Assert.Contains("ends with a line that says so", install);
        }
    }
}
