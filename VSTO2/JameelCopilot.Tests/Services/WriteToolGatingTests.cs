using System.IO;
using System.Linq;
using JameelCopilot.Services.Tools;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>
    /// The four mailbox-changing tools exist only when an administrator named
    /// them. Nothing may fall back to "all of them".
    /// </summary>
    [Collection("Config")]
    public class WriteToolGatingTests
    {
        [Fact]
        public void NoneEnabled_NoneAdvertised_EvenWhenACallerAsksForWriteTools()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                var names = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: true).Select(t => (string)t["name"]).ToList();

                Assert.Empty(names.Intersect(Config.AllWriteTools));
                Assert.Equal(20, names.Count);   // 17 that read the mailbox, the calendar, the address book and org hierarchy, and three that only propose (a note, a meeting, a rule)
            }
        }

        [Fact]
        public void OneEnabled_OnlyThatOneAdvertised()
        {
            using (new ConfigStateScope(TestSettings.Configured(writeTools: "outlook_flag_message")))
            {
                var names = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: true).Select(t => (string)t["name"]).ToList();

                Assert.Equal(new[] { "outlook_flag_message" }, names.Intersect(Config.AllWriteTools));
            }
        }

        // A missing setting must mean "no write tools", never "every write tool".
        [Theory]
        [InlineData("Services", "OutlookToolHost.cs")]
        [InlineData("Services", "Tools", "ToolCatalogSchema.cs")]
        public void NoFallbackToEveryWriteTool(params string[] path)
        {
            var source = File.ReadAllText(RepoFiles.Find(new[] { "VSTO2", "JameelCopilot" }.Concat(path).ToArray()));

            Assert.DoesNotContain("(Config.AllWriteTools)", source);
        }
    }
}
