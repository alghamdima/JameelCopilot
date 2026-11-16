using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Installer
{
    /// <summary>
    /// What the MSI is authored to do, read from Installer\Package.wxs. The
    /// installed behaviour is checked by hand on a clean PC (Deploy\README.txt).
    /// </summary>
    public class PackageSourceTests
    {
        private static readonly XNamespace Wix = "http://wixtoolset.org/schemas/v4/wxs";
        private static readonly string PackagePath = RepoFiles.Find("Installer", "Package.wxs");
        private static readonly string PackageText = File.ReadAllText(PackagePath);
        private static readonly XDocument Package = XDocument.Parse(PackageText);

        private static IEnumerable<XElement> All(string name) => Package.Descendants(Wix + name);

        private static XElement ComponentWithKey(string key, string bitness) =>
            All("Component").Single(c => (string)c.Attribute("Bitness") == bitness
                && c.Descendants(Wix + "RegistryKey").Any(k => (string)k.Attribute("Key") == key));

        private static string Value(XElement component, string key, string name) =>
            (string)component.Descendants(Wix + "RegistryKey").Single(k => (string)k.Attribute("Key") == key)
                .Elements(Wix + "RegistryValue").Single(v => (string)v.Attribute("Name") == name).Attribute("Value");

        [Fact]
        public void IsAPerMachineUpgradeablePackageNamedHB()
        {
            var package = All("Package").Single();

            Assert.Equal("Jameel Copilot", (string)package.Attribute("Name"));
            Assert.Equal("perMachine", (string)package.Attribute("Scope"));
            // Never changes: it is how a new version finds and replaces the old one.
            Assert.Equal("D2979919-ADE1-4A9B-953B-970B96F42252", (string)package.Attribute("UpgradeCode"));
            Assert.Single(All("MajorUpgrade"));
        }

        // Outlook reads the 64-bit or the 32-bit registry depending on its own bitness.
        [Theory]
        [InlineData("always64")]
        [InlineData("always32")]
        public void AddInIsRegisteredForBothBitnesses(string bitness)
        {
            const string key = @"Software\Microsoft\Office\Outlook\Addins\JameelCopilot";
            var component = ComponentWithKey(key, bitness);

            Assert.Equal("3", Value(component, key, "LoadBehavior"));
            Assert.Equal("Jameel Copilot", Value(component, key, "FriendlyName"));
            Assert.Equal("file:///[INSTALLFOLDER]JameelCopilot.vsto|vstolocal", Value(component, key, "Manifest"));
        }

        [Fact]
        public void InstallsUnderProgramFilesJameelCopilot()
        {
            var folder = All("Directory").Single(d => (string)d.Attribute("Id") == "INSTALLFOLDER");

            Assert.Equal("JameelCopilot", (string)folder.Attribute("Name"));
            Assert.Equal("ProgramFiles64Folder", (string)folder.Parent.Attribute("Id"));
        }

        [Fact]
        public void SettingsGoToHklmJcOutlook_In64BitView()
        {
            const string key = @"SOFTWARE\JameelCopilot";
            var component = ComponentWithKey(key, "always64");

            Assert.Equal("[API_BASE_URL]", Value(component, key, "ApiBaseUrl"));
            Assert.Equal("[MODEL]", Value(component, key, "Model"));
            Assert.Equal("[REASONING_EFFORT]", Value(component, key, "ReasoningEffort"));
            Assert.Equal("[WRITE_TOOLS]", Value(component, key, "EnabledWriteTools"));
            Assert.Equal("[MAX_BULK_EXPORT_ROWS]", Value(component, key, "MaxBulkExportRows"));
            Assert.Equal("[TEMPERATURE]", Value(component, key, "Temperature"));
            Assert.Equal("[TEMPERATURE_MAX]", Value(component, key, "TemperatureMax"));
            Assert.Equal("[MAX_OUTPUT_TOKENS]", Value(component, key, "MaxOutputTokens"));
            // Only the encrypted key is ever written.
            Assert.Equal("[API_KEY_PROTECTED]", Value(component, key, "ApiKeyProtected"));
            Assert.DoesNotContain(All("RegistryValue"), v => ((string)v.Attribute("Value") ?? "").Contains("[API_KEY]"));
            Assert.All(All("RegistryKey").Where(k => (string)k.Attribute("Key") == key),
                k => Assert.Equal("HKLM", (string)k.Attribute("Root")));
        }

        [Theory]
        [InlineData("API_KEY")]
        [InlineData("API_KEY_PROTECTED")]
        [InlineData("API_KEY_PROTECTED_SAVED")]
        public void KeyPropertiesAreHiddenFromLogs(string property)
        {
            var declared = All("Property").Single(p => (string)p.Attribute("Id") == property);

            Assert.Equal("yes", (string)declared.Attribute("Hidden"));
        }

        // An upgrade or a repair keeps every setting that is not passed again.
        [Theory]
        [InlineData("API_BASE_URL", "ApiBaseUrl")]
        [InlineData("MODEL", "Model")]
        [InlineData("REASONING_EFFORT", "ReasoningEffort")]
        [InlineData("WRITE_TOOLS", "EnabledWriteTools")]
        [InlineData("MAX_BULK_EXPORT_ROWS", "MaxBulkExportRows")]
        [InlineData("API_KEY_PROTECTED", "ApiKeyProtected")]
        [InlineData("TEMPERATURE", "Temperature")]
        [InlineData("TEMPERATURE_MAX", "TemperatureMax")]
        [InlineData("MAX_OUTPUT_TOKENS", "MaxOutputTokens")]
        public void SettingsAreRememberedAcrossUpgrades(string property, string valueName)
        {
            var saved = All("Property").Single(p => (string)p.Attribute("Id") == property + "_SAVED");
            var search = saved.Elements(Wix + "RegistrySearch").Single();
            Assert.Equal(@"SOFTWARE\JameelCopilot", (string)search.Attribute("Key"));
            Assert.Equal(valueName, (string)search.Attribute("Name"));
            Assert.Equal("always64", (string)search.Attribute("Bitness"));

            // The saved value is used only when the command line gave none.
            var restore = All("SetProperty").Single(s => (string)s.Attribute("Id") == property);
            Assert.Equal("[" + property + "_SAVED]", (string)restore.Attribute("Value"));
            Assert.Contains(property + "_SAVED", (string)restore.Attribute("Condition"));
            Assert.Contains("NOT ", (string)restore.Attribute("Condition"));
        }

        // A standard user can start a repair of an installed per-machine
        // product, and Windows Installer runs it elevated. A property marked
        // Secure is taken from that user's command line. The settings must
        // not be, or any user of the PC could repoint Jameel Copilot (every user's mail
        // content, and the key) at another server.
        [Theory]
        [InlineData("API_BASE_URL")]
        [InlineData("API_KEY")]
        [InlineData("API_KEY_PROTECTED")]
        [InlineData("MODEL")]
        [InlineData("REASONING_EFFORT")]
        [InlineData("WRITE_TOOLS")]
        [InlineData("MAX_BULK_EXPORT_ROWS")]
        [InlineData("TEMPERATURE")]
        [InlineData("TEMPERATURE_MAX")]
        [InlineData("MAX_OUTPUT_TOKENS")]
        public void SettingCannotBeSuppliedByAStandardUser(string property)
        {
            Assert.DoesNotContain(All("Property"),
                p => (string)p.Attribute("Id") == property && (string)p.Attribute("Secure") == "yes");
        }

        // The saved values are restored on the client (for the checks shown
        // before elevation) and again on the elevated side, which never
        // trusts what the client sends.
        [Fact]
        public void SavedSettingsAreRestoredInBothSequences()
        {
            Assert.NotEmpty(All("SetProperty"));
            Assert.All(All("SetProperty"), s => Assert.Equal("both", (string)s.Attribute("Sequence")));
        }

        // Running the installed MSI again is a maintenance run, and Windows
        // Installer rewrites registry values in one only when a reinstall is
        // asked for, with "m" in its mode. Without this the documented
        // "run it again with MODEL=..." ends with exit code 0 and changes nothing.
        [Fact]
        public void RunningTheInstalledMsiAgain_RewritesTheSettings()
        {
            var reinstall = All("SetProperty").Single(s => (string)s.Attribute("Id") == "REINSTALL");
            var mode = All("SetProperty").Single(s => (string)s.Attribute("Id") == "REINSTALLMODE");

            Assert.Equal("ALL", (string)reinstall.Attribute("Value"));
            Assert.Contains("m", (string)mode.Attribute("Value"));
            foreach (var action in new[] { reinstall, mode })
            {
                // Only for an installed product that is not being removed, and before feature states are decided.
                Assert.StartsWith("Installed AND NOT REMOVE", (string)action.Attribute("Condition"));
                Assert.Equal("CostInitialize", (string)action.Attribute("Before"));
            }
        }

        // The old version is removed first, so no file stays behind because
        // its version number did not change. The settings were read into
        // properties before that, so they survive.
        [Fact]
        public void AnUpgradeRemovesTheOldVersionBeforeInstallingTheNew()
        {
            Assert.Equal("afterInstallInitialize", (string)All("MajorUpgrade").Single().Attribute("Schedule"));
        }

        // A run with a user interface started by a standard user reaches the
        // elevated side without the saved values. It must stop there and not
        // write empty settings: the three are checked on every run but an uninstall.
        [Fact]
        public void SettingsAreCheckedOnEveryRunButAnUninstall()
        {
            var settings = All("Launch").Select(l => (string)l.Attribute("Condition"))
                .Where(c => c.Contains("API_BASE_URL") || c.Contains("MODEL") || c.Contains("API_KEY")).ToList();

            Assert.Equal(3, settings.Count);
            Assert.All(settings, c => Assert.StartsWith("REMOVE~=\"ALL\" OR ", c));
            Assert.All(settings, c => Assert.DoesNotContain("Installed", c));
        }

        // One version number: the MSI's, the DLL's, the manifests' and the one About shows.
        [Fact]
        public void TheBuildStampsOneVersionEverywhere()
        {
            var project = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));
            var assemblyInfo = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "Properties", "AssemblyInfo.cs"));
            var buildMsi = File.ReadAllText(RepoFiles.Find("Installer", "Build-Msi.ps1"));

            Assert.Contains("<ApplicationVersion>$(JcVersion).0</ApplicationVersion>", project);
            Assert.Contains("System.Reflection.AssemblyVersion", project);
            Assert.Contains("System.Reflection.AssemblyFileVersion", project);
            Assert.DoesNotContain("AssemblyVersion(", assemblyInfo);
            Assert.DoesNotContain("AssemblyFileVersion(", assemblyInfo);
            Assert.Contains("Version = $Version", buildMsi);
            // The script refuses to package a DLL that carries another version.
            Assert.Contains("FileVersion", buildMsi);
        }

        // The MSI's signature does not cover a DLL once it is on disk: where
        // only signed code may load, each DLL needs its own. The add-in's is
        // signed before the manifests record its hash.
        [Fact]
        public void TheBuildSignsTheDllsAndRefusesToPackageAnUnsignedOne()
        {
            var project = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));
            var actions = File.ReadAllText(RepoFiles.Find("Installer", "Actions", "JameelCopilot.Installer.Actions.csproj"));
            var buildMsi = File.ReadAllText(RepoFiles.Find("Installer", "Build-Msi.ps1"));

            Assert.Matches("AfterTargets=\"CoreCompile\"[^>]*>\\s*<SignFile ", project);
            Assert.Contains("<SignFile ", actions);
            Assert.Contains("Get-AuthenticodeSignature", buildMsi);
        }

        // Settings > Apps and Programs and Features show the Jameel Copilot mark, not a blank box.
        [Fact]
        public void ProgramsAndFeatures_ShowsTheJcIcon()
        {
            var icon = All("Icon").Single();
            var shown = All("Property").Single(p => (string)p.Attribute("Id") == "ARPPRODUCTICON");

            Assert.EndsWith(@"VSTO2\JameelCopilot\Resources\jc.ico", (string)icon.Attribute("SourceFile"));
            Assert.Equal((string)icon.Attribute("Id"), (string)shown.Attribute("Value"));
        }

        [Fact]
        public void RefusesToInstallWithoutAnHttpsAddressAModelAndAKey()
        {
            var conditions = All("Launch").Select(l => (string)l.Attribute("Condition")).ToList();

            Assert.Contains(conditions, c => c.Contains("API_BASE_URL") && c.Contains("https://"));
            Assert.Contains(conditions, c => c.Contains("MODEL"));
            Assert.Contains(conditions, c => c.Contains("API_KEY_FILE") && c.Contains("API_KEY_PROTECTED"));
            Assert.All(All("Launch"), l => Assert.False(string.IsNullOrWhiteSpace((string)l.Attribute("Message"))));
        }

        [Fact]
        public void ChecksWindowsDotNetAndWebView2()
        {
            var conditions = string.Join("\n", All("Launch").Select(l => (string)l.Attribute("Condition")));

            Assert.Contains("VersionNT64", conditions);
            Assert.Contains("WIX_IS_NETFRAMEWORK_472_OR_LATER_INSTALLED", conditions);
            Assert.Contains("WEBVIEW2", conditions);
            Assert.Contains("{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}", PackageText);
        }

        // Click-to-Run Office keeps its copy of the VSTO runtime where an
        // installer cannot see it; a check would block every company PC.
        [Fact]
        public void DoesNotCheckTheVstoRuntime()
        {
            Assert.DoesNotContain("VSTO Runtime Setup", PackageText);
        }

        // The old script loosened ClickOnce trust for the whole machine.
        [Fact]
        public void DoesNotTouchMachineTrustSettings()
        {
            Assert.DoesNotContain("TrustManager", PackageText);
            Assert.DoesNotContain("PromptingLevel", PackageText);
            Assert.DoesNotContain(@"VSTO\Security\Inclusion", PackageText);
        }

        [Fact]
        public void NothingClosesOutlook()
        {
            Assert.Empty(Package.Descendants().Where(e => e.Name.LocalName == "CloseApplication"));
            Assert.DoesNotContain("taskkill", PackageText.ToLowerInvariant());
        }

        [Fact]
        public void EncryptsTheKeyBeforeAnythingIsWritten()
        {
            var action = All("CustomAction").Single(a => (string)a.Attribute("Id") == "ProtectApiKey");
            Assert.Equal("immediate", (string)action.Attribute("Execute"));
            Assert.Equal("ProtectApiKey", (string)action.Attribute("DllEntry"));

            var scheduled = All("InstallExecuteSequence").Elements(Wix + "Custom")
                .Single(c => (string)c.Attribute("Action") == "ProtectApiKey");
            Assert.Contains("API_KEY", (string)scheduled.Attribute("Condition"));
            Assert.Contains("API_KEY_FILE", (string)scheduled.Attribute("Condition"));
        }

        // One implementation of the key format, shared with the add-in.
        [Fact]
        public void InstallerActions_UseTheAddInsKeyProtector()
        {
            var project = File.ReadAllText(RepoFiles.Find("Installer", "Actions", "JameelCopilot.Installer.Actions.csproj"));
            var action = File.ReadAllText(RepoFiles.Find("Installer", "Actions", "ProtectApiKeyAction.cs"));

            Assert.Contains(@"..\..\VSTO2\JameelCopilot\Services\ApiKeyProtector.cs", project);
            Assert.Contains("ApiKeyProtector.Protect(", action);
            Assert.DoesNotContain("session.Log(key", action);
            // The key file is read next to the MSI, not in the action's working folder (System32).
            Assert.Contains("session[\"API_KEY_FILE\"]", action);
            Assert.Contains("session[\"OriginalDatabase\"]", action);
        }

        // The developer script writes the same values the MSI does.
        [Theory]
        [InlineData("Temperature")]
        [InlineData("TemperatureMax")]
        [InlineData("MaxOutputTokens")]
        public void DeveloperScript_WritesTheModelSettings(string valueName)
        {
            var script = File.ReadAllText(RepoFiles.Find("Installer", "Set-JameelCopilotConfig.ps1"));

            Assert.Contains("$settings.SetValue('" + valueName + "', $" + valueName + ")", script);
        }

        [Theory]
        [InlineData("Build-Msi.ps1")]
        [InlineData("Set-JameelCopilotConfig.ps1")]
        public void Scripts_AreAsciiOnly(string script)
        {
            // Windows PowerShell 5.1 reads a script without a byte-order mark as ANSI.
            Assert.DoesNotContain(File.ReadAllBytes(RepoFiles.Find("Installer", script)), b => b > 127);
        }
    }
}
