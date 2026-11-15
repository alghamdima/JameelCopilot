using System;
using System.IO;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// Outlook times every add-in while it starts and disables one whose start
    /// takes over a second (the median of five starts). Jameel Copilot was disabled that
    /// way on the owner's PC: it built its network client, read its settings
    /// and hooked the mailbox search before Outlook had drawn its window.
    /// Starting must do nothing that can wait for the first click.
    /// </summary>
    public class StartupTests
    {
        private static readonly string AddIn =
            File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "ThisAddIn.cs"));

        private static string Body(string from, string to)
        {
            var start = AddIn.IndexOf(from, StringComparison.Ordinal);
            var end = AddIn.IndexOf(to, start + 1, StringComparison.Ordinal);
            Assert.True(start >= 0 && end > start, "Not found: " + from);
            return AddIn.Substring(start, end - start);
        }

        [Theory]
        [InlineData("new CustomApiChatService")]      // builds the HTTP client: the slowest single step
        [InlineData("Config.Load")]                   // registry and the key's decryption
        [InlineData("new OutlookThreadMarshaller")]
        [InlineData("new MailboxUserStore")]
        [InlineData("LiveAdvancedSearchHost")]        // subscribes to an Outlook event
        [InlineData("new ChatArchive")]
        [InlineData("new UserMemoryService")]
        [InlineData("EnsureServices")]
        public void Startup_BuildsNothing(string work)
        {
            var startup = Body("private void ThisAddIn_Startup(", "private void EnsureServices()");

            Assert.DoesNotContain(work, startup);
        }

        [Fact]
        public void TheServices_AreBuiltOnce_ByTheFirstClick()
        {
            var ensure = Body("private void EnsureServices()", "private void ThisAddIn_Shutdown(");

            Assert.Contains("if (_servicesReady) return;", ensure);
            Assert.Contains("ChatService = new CustomApiChatService();", ensure);
            Assert.Contains("Config.Load();", ensure);
        }

        // Each button of Jameel Copilot's ends in one of these two methods.
        [Theory]
        [InlineData("public void ShowTaskPane()", "ShowInspectorTaskPane(insp)")]
        [InlineData("public void ShowSettings()", "SettingsForm.Open(ChatService)")]
        public void EveryEntryPoint_BuildsTheServicesBeforeUsingThem(string method, string firstUse)
        {
            var start = AddIn.IndexOf(method, StringComparison.Ordinal);
            var ensure = AddIn.IndexOf("EnsureServices();", start, StringComparison.Ordinal);
            var use = AddIn.IndexOf(firstUse, start, StringComparison.Ordinal);

            Assert.True(start >= 0 && ensure > start && use > ensure, method + " must call EnsureServices() before " + firstUse);
        }

        // The ribbon's callbacks run while Outlook starts: they must not reach for the services.
        [Fact]
        public void TheRibbon_UsesNoServiceToDrawItself()
        {
            var ribbon = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "Ribbon.cs"));
            var picture = ribbon.Substring(ribbon.IndexOf("public Bitmap GetJcImage(", StringComparison.Ordinal));
            picture = picture.Substring(0, picture.IndexOf("private static string GetResourceText", StringComparison.Ordinal));

            Assert.DoesNotContain("Globals.ThisAddIn", picture);
            Assert.DoesNotContain("Config.", picture);
        }
    }
}
