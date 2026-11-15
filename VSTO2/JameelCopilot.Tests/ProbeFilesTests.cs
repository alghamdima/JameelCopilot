using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// The files the owner copies to a PC on the company network to try the
    /// model: they are run by hand, far from this build, so what can be
    /// checked here is checked here.
    /// </summary>
    public class ProbeFilesTests
    {
        private static string Tool(string name) => RepoFiles.Find("tools", name);

        // Windows PowerShell 5.1 reads a script without a byte-order mark as ANSI.
        [Theory]
        [InlineData("Probe-CustomApi.ps1")]
        [InlineData("Probe-Language.ps1")]
        public void ProbeScripts_AreAsciiOnly(string script)
        {
            Assert.DoesNotContain(File.ReadAllBytes(Tool(script)), b => b > 127);
        }

        [Fact]
        public void LanguageCases_AreWellFormed()
        {
            var cases = JArray.Parse(File.ReadAllText(Tool("language-cases.json")));

            Assert.Equal(
                new[]
                {
                    "arabic-question", "arabic-dialect", "arabic-with-english-terms",
                    "english-question", "arabic-about-english-email", "english-after-arabic",
                },
                cases.Select(c => (string)c["name"]));
            foreach (var c in cases)
            {
                Assert.Equal("user", (string)c["messages"].Last()["role"]);
                Assert.Contains((string)c["expect"], new[] { "ar", "en" });
                if ((string)c["expect"] == "ar") Assert.InRange((double)c["minArabic"], c["context"] == null ? 0.5 : 0.3, 1.0);
            }
            // The Arabic survived the file's encoding.
            Assert.Contains(((string)cases[0]["messages"][0]["content"]), ch => ch >= '؀' && ch <= 'ۿ');
            // One case puts an English email in front of an Arabic question.
            Assert.Contains("sara@example.com", (string)cases[4]["context"]);
        }

        // A correct Arabic answer about an English email keeps the name, the
        // address and the subject as written and may quote a sentence. The
        // probe must not call that answer mixed.
        [Fact]
        public void LanguageCases_AnArabicAnswerThatQuotesTheEnglishEmail_Passes()
        {
            var probeCase = JArray.Parse(File.ReadAllText(Tool("language-cases.json")))
                .Single(c => (string)c["name"] == "arabic-about-english-email");
            const string answer =
                "المرسلة Sara Example (sara@example.com) تطلب في رسالتها بعنوان Q3 invoice تأكيد إجمالي الفاتورة " +
                "وهو 4,200 قبل يوم الخميس. نص الطلب: \"Please confirm the invoice total of 4,200 by Thursday.\"";

            // The probe's own measure (Get-ArabicShare): Arabic letters among all letters.
            var script = File.ReadAllText(Tool("Probe-Language.ps1"));
            Assert.Contains("[regex]::Matches($Text, '\\p{L}').Count", script);
            Assert.Contains("[regex]::Matches($Text, '[\\p{IsArabic}-[\\P{L}]]').Count", script);
            var share = (double)Regex.Matches(answer, "[\\p{IsArabic}-[\\P{L}]]").Count / Regex.Matches(answer, "\\p{L}").Count;

            Assert.InRange(share, 0.4, 0.6); // far from an English answer (under 0.1), short of 0.6
            Assert.True(share >= (double)probeCase["minArabic"] + 0.05,
                "A correct answer scores " + share.ToString("0.00") + "; the case asks for " + probeCase["minArabic"]);
        }

        // The probe shows the model the same prompt shape the add-in sends.
        [Fact]
        public void LanguageProbe_UsesTheBuildersContextHeader_AndThePromptFile()
        {
            var script = File.ReadAllText(Tool("Probe-Language.ps1"));

            Assert.Contains(SystemPromptBuilder.ContextHeader, script);
            Assert.Contains("core-agent.txt", script);
            Assert.True(File.Exists(RepoFiles.Find("VSTO2", "JameelCopilot", "Prompts", "core-agent.txt")));
        }

        // The key is typed hidden or read from a file; it is never printed.
        [Theory]
        [InlineData("Probe-CustomApi.ps1")]
        [InlineData("Probe-Language.ps1")]
        public void ProbeScripts_NeverPrintTheKey(string script)
        {
            var lines = File.ReadAllLines(Tool(script));

            Assert.DoesNotContain(lines, line => line.Contains("Write-") && line.Contains("$key"));
            Assert.Contains(lines, line => line.Contains("-AsSecureString"));
            Assert.Contains(lines, line => line.Contains("'^https://'"));
        }
    }
}
