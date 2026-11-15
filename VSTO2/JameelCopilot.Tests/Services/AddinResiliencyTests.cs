using System;
using Microsoft.Win32;
using JameelCopilot.Services;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>
    /// Outlook disabled Jameel Copilot on the owner's PC for starting slowly, even after
    /// Jameel Copilot stopped doing work at start: the runtimes' own load is counted. Jameel Copilot
    /// now does what "Always enable this add-in" does.
    /// </summary>
    public class AddinResiliencyTests : IDisposable
    {
        private readonly string _path = @"Software\JameelCopilot.Tests\" + Guid.NewGuid().ToString("N");
        private readonly RegistryKey _root;

        public AddinResiliencyTests()
        {
            _root = Registry.CurrentUser.CreateSubKey(_path);
        }

        public void Dispose()
        {
            _root.Dispose();
            Registry.CurrentUser.DeleteSubKeyTree(_path, false);
        }

        private object Listed(string major = "16.0")
        {
            using (var list = _root.OpenSubKey(@"Software\Microsoft\Office\" + major + @"\Outlook\Resiliency\DoNotDisableAddinList"))
            {
                return list?.GetValue("JameelCopilot");
            }
        }

        [Fact]
        public void PutsJcOnTheUsersDoNotDisableList_AsAlwaysEnableDoes()
        {
            Assert.True(AddinResiliency.KeepEnabled("16.0.17928.20114", _root));

            Assert.Equal(1, Listed());
            using (var list = _root.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"))
            {
                Assert.Equal(RegistryValueKind.DWord, list.GetValueKind("JameelCopilot"));
            }
        }

        // A value already there (written by Outlook when the user chose "Always enable") is left as it is.
        [Fact]
        public void AValueAlreadyThere_IsLeftAlone()
        {
            using (var list = _root.CreateSubKey(@"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"))
            {
                list.SetValue("JameelCopilot", 3, RegistryValueKind.DWord);
            }

            Assert.True(AddinResiliency.KeepEnabled("16.0.1", _root));
            Assert.Equal(3, Listed());
        }

        [Fact]
        public void UsesOutlooksOwnVersionKey()
        {
            AddinResiliency.KeepEnabled("15.0.5589.1000", _root);

            Assert.Equal(1, Listed("15.0"));
            Assert.Null(Listed("16.0"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("abc")]
        [InlineData("16")]
        [InlineData("12.0.1")]
        public void NoUsableVersion_WritesNothing(string version)
        {
            Assert.False(AddinResiliency.KeepEnabled(version, _root));
            Assert.Null(Listed());
            Assert.Null(AddinResiliency.MajorVersion(version));
        }

        [Fact]
        public void StartupNamesJcOnTheList_WithoutBuildingAnything()
        {
            var addIn = System.IO.File.ReadAllText(JameelCopilot.Tests.Helpers.RepoFiles.Find("VSTO2", "JameelCopilot", "ThisAddIn.cs"));
            var startup = addIn.Substring(addIn.IndexOf("private void ThisAddIn_Startup(", StringComparison.Ordinal));
            startup = startup.Substring(0, startup.IndexOf("private void EnsureServices()", StringComparison.Ordinal));

            Assert.Contains("AddinResiliency.KeepEnabled(", startup);
        }
    }
}
