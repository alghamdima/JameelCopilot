using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// Guards on the chat and PDF page sources. The pages have no JavaScript
    /// test runner, so these read the files.
    /// shortcut: source checks only, add a JS test runner if markdown.js grows beyond rendering.
    /// </summary>
    public class WebUiSourceTests
    {
        private static string WebUi(string name) =>
            File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "WebUI", name));

        // Arabic text must align right and order correctly: each block takes
        // its direction from its own first letter.
        [Theory]
        [InlineData("'<' + listType + ' dir=\"auto\">'")]
        [InlineData("'<li dir=\"auto\">'")]
        [InlineData("'<p dir=\"auto\">'")]
        [InlineData("'<h' + level + ' dir=\"auto\">'")]
        [InlineData("'<blockquote dir=\"auto\">'")]
        [InlineData("'<th dir=\"auto\">'")]
        [InlineData("'<td dir=\"auto\">'")]
        public void Markdown_BlocksAreDirAuto(string emitted)
        {
            Assert.Contains(emitted, WebUi("markdown.js"));
        }

        [Theory]
        [InlineData("'<' + listType + '>'")]
        [InlineData("'<li>'")]
        [InlineData("'<p>'")]
        [InlineData("'<h' + level + '>'")]
        [InlineData("'<blockquote>'")]
        [InlineData("'<th>'")]
        [InlineData("'<td>'")]
        public void Markdown_NoBlockWithoutDirection(string emitted)
        {
            Assert.DoesNotContain(emitted, WebUi("markdown.js"));
        }

        [Fact]
        public void ChatInput_IsDirAuto()
        {
            Assert.Matches(new Regex("<textarea[^>]*\\bdir=\"auto\"", RegexOptions.Singleline), WebUi("index.html"));
        }

        [Fact]
        public void PdfTitleAndSubtitle_AreDirAuto()
        {
            var template = WebUi("print-template.html");

            Assert.Contains("class=\"doc-title\" dir=\"auto\"", template);
            Assert.Contains("class=\"doc-subtitle\" dir=\"auto\"", template);
        }

        // The header shows the mark alone (the owner's choice); the mark itself is named for a screen reader.
        [Fact]
        public void Header_ShowsTheMarkAlone()
        {
            var html = WebUi("index.html");

            Assert.DoesNotContain("app-name", html);
            Assert.Contains("<img class=\"app-mark\" src=\"jc-mark.svg\" alt=\"Jameel Copilot\">", html);
        }

        private static string AddIn(params string[] parts) =>
            File.ReadAllText(RepoFiles.Find(new[] { "VSTO2", "JameelCopilot" }.Concat(parts).ToArray()));

        private static readonly string[][] Controllers =
        {
            new[] { "TaskPane", "Chat", "ChatController.cs" },
            new[] { "TaskPane", "InboxCopilot", "InboxCopilotController.cs" },
        };

        // Model output is untrusted. A link is shown as text, never as an
        // anchor: a crafted email could make the model print a link that
        // carries mail content in its address.
        [Theory]
        [InlineData("markdown.js")]
        [InlineData("chat.js")]
        public void NoAnchorsInWebUi(string script)
        {
            var js = WebUi(script);

            Assert.DoesNotContain("href", js);
            Assert.DoesNotContain("<a ", js);
            Assert.DoesNotContain("'a'", js);
        }

        [Fact]
        public void IndexHtml_HasCsp()
        {
            Assert.Contains(
                "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; script-src 'self'; style-src 'self'; "
                + "img-src 'self' data:; font-src 'self'; connect-src 'none'; base-uri 'none'; form-action 'none'\">",
                WebUi("index.html"));
        }

        [Fact]
        public void PrintTemplate_HasNonceCsp()
        {
            Assert.Contains(
                "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; script-src 'nonce-__NONCE__'; "
                + "style-src 'nonce-__NONCE__'; img-src data:; form-action 'none'\">",
                WebUi("print-template.html"));
        }

        // Every WebView2 the add-in creates goes through the one lock-down helper.
        [Fact]
        public void EveryWebView_IsLockedDown()
        {
            var root = Path.GetDirectoryName(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));
            var creators = Directory.GetFiles(root, "*.cs", SearchOption.AllDirectories)
                .Where(f => !f.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
                         && !f.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar))
                .Where(f => File.ReadAllText(f).Contains("EnsureCoreWebView2Async("))
                .ToList();

            Assert.NotEmpty(creators);
            Assert.All(creators, f => Assert.Contains("LockDown(", File.ReadAllText(f)));
        }

        [Fact]
        public void WebMessages_AreAcceptedOnlyFromOwnPages()
        {
            foreach (var controller in Controllers)
            {
                var source = AddIn(controller);
                var handler = source.Substring(source.IndexOf("private void OnWebMessageReceived(", System.StringComparison.Ordinal));
                var check = handler.IndexOf("IsTrustedSource(e.Source)", System.StringComparison.Ordinal);
                var read = handler.IndexOf("TryGetWebMessageAsString", System.StringComparison.Ordinal);

                Assert.True(check >= 0 && check < read,
                    controller.Last() + " must check the message source before reading the message");
            }
        }

        [Theory]
        [InlineData("styles.css")]
        [InlineData("print-styles.css")]
        public void Stylesheets_AlignToTheTextDirection(string stylesheet)
        {
            var css = WebUi(stylesheet);

            Assert.DoesNotContain("text-align: left", css);
            Assert.DoesNotContain("padding-left", css);
        }

        [Fact]
        public void ChatLists_IndentOnTheTextSide()
        {
            Assert.Contains("padding-inline-start", WebUi("styles.css"));
        }

        // An overflow:hidden card (the draft) would shrink and clip instead of scrolling.
        [Fact]
        public void Messages_DoNotShrinkToFitThePane()
        {
            Assert.Contains(".messages > * { flex-shrink: 0; }", WebUi("styles.css"));
        }

        // Narrow pane tables must scroll cleanly rather than splitting short words mid-letter ("Tim e").
        [Fact]
        public void MarkdownTables_DoNotSplitWordsMidLetter_AndLiveExplorerFallsBackSafely()
        {
            var css = WebUi("styles.css");
            Assert.Contains(".msg-content table.md-table th:first-child", css);
            Assert.Contains("overflow-wrap: normal;", css);

            var surface = AddIn("Services", "Tools", "LiveOutlookSurface.cs");
            Assert.Contains("private Outlook.Explorer GetLiveExplorer()", surface);
        }
    }
}
