using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// The add-in talks to one place: the Custom API address an administrator
    /// configured. No other address may be written into its code or pages.
    /// </summary>
    public class EgressGuardTests
    {
        // A scheme followed by a host character. "https://" on its own (built
        // up with the virtual host, or named in a message) is not an address.
        private static readonly Regex UrlLiteral = new Regex(@"\b(https?|wss?|ftp)://[A-Za-z0-9\[]", RegexOptions.IgnoreCase);

        private static readonly string[] Scanned = { ".cs", ".js", ".html", ".css", ".xml", ".config" };

        private static bool IsAllowed(string line)
        {
            // XML and MAPI names that are identifiers, not places anything is sent to.
            if (line.Contains("schemas.microsoft.com")) return true;
            var trimmed = line.TrimStart();
            var isComment = trimmed.StartsWith("//", StringComparison.Ordinal) || trimmed.StartsWith("*", StringComparison.Ordinal);
            return isComment && line.Contains("learn.microsoft.com");
        }

        [Fact]
        public void NoUrlLiteralsOutsideAllowList()
        {
            var root = Path.GetDirectoryName(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));
            var offenders = new List<string>();
            foreach (var file in Directory.GetFiles(root, "*.*", SearchOption.AllDirectories))
            {
                var relative = file.Substring(root.Length + 1);
                if (relative.StartsWith("bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    || relative.StartsWith("obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    || !Scanned.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                var lines = File.ReadAllLines(file);
                for (var i = 0; i < lines.Length; i++)
                {
                    if (UrlLiteral.IsMatch(lines[i]) && !IsAllowed(lines[i]))
                    {
                        offenders.Add(relative + ":" + (i + 1) + ": " + lines[i].Trim());
                    }
                }
            }

            Assert.True(offenders.Count == 0, "Address literals found:\n" + string.Join("\n", offenders));
        }

        [Theory]
        [InlineData("var x = \"https://evil.example/\";", false)]
        [InlineData("socket = new WebSocket('wss://api.example/v1');", false)]
        [InlineData("// see https://evil.example/docs", false)]
        [InlineData("// https://learn.microsoft.com/en-us/office/vba/outlook", true)]
        [InlineData("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">", true)]
        public void AllowList_IsNarrow(string line, bool allowed)
        {
            Assert.True(UrlLiteral.IsMatch(line));
            Assert.Equal(allowed, IsAllowed(line));
        }

        [Theory]
        [InlineData("Navigate(\"https://\" + WebView2Bootstrap.VirtualHost + \"/index.html\");")]
        [InlineData("\"ApiBaseUrl is missing or is not an https:// address.\"")]
        public void SchemeWithoutAHost_IsNotAnAddress(string line)
        {
            Assert.False(UrlLiteral.IsMatch(line));
        }
    }
}
