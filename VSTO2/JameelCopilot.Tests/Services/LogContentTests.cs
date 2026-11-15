using System.IO;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>What the panes write to the trace log about a chat turn.</summary>
    public class LogContentTests
    {
        [Theory]
        [InlineData("TaskPane", "Chat", "ChatController.cs")]
        [InlineData("TaskPane", "InboxCopilot", "InboxCopilotController.cs")]
        public void Controllers_LogNoMessageTextToolArgumentsOrResults(string a, string b, string c)
        {
            var source = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", a, b, c));

            // The start of a web message is the user's prompt; arguments and
            // result summaries are mail content.
            Assert.DoesNotContain("json.Substring", source);
            Assert.DoesNotContain("argsJson.Substring", source);
            Assert.DoesNotContain("\" summary=\" + summary", source);
            Assert.Contains("TraceLog.Write(\"WebMessageReceived: \" + type", source);
        }
    }
}
