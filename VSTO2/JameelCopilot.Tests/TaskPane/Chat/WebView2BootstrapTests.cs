using System.IO;
using JameelCopilot.TaskPane.Chat;
using Xunit;

namespace JameelCopilot.Tests.TaskPane.Chat
{
    public class WebView2BootstrapTests
    {
        // Resource names use "." as both folder separator and extension
        // boundary. The helper has to disambiguate: last dot = extension,
        // earlier dots = folder separators, and ".min.js" is treated as a
        // single suffix on the filename (not a fake folder).

        [Fact]
        public void ResourceNameToRelativePath_SimpleFile()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("index.html");
            Assert.Equal("index.html", result);
        }

        [Fact]
        public void ResourceNameToRelativePath_SingleFolderPrefix()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.marked.js");
            Assert.Equal(Path.Combine("vendor", "marked.js"), result);
        }

        [Fact]
        public void ResourceNameToRelativePath_MinifiedFile_KeepsMinOnFilename()
        {
            // "vendor.marked.min.js" must become vendor\marked.min.js, NOT
            // vendor\marked\min.js.
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.marked.min.js");
            Assert.Equal(Path.Combine("vendor", "marked.min.js"), result);
        }

        [Fact]
        public void ResourceNameToRelativePath_NestedFolder()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.highlight.styles.github.css");
            Assert.Equal(
                Path.Combine("vendor", "highlight", "styles", "github.css"),
                result);
        }

        [Fact]
        public void ResourceNameToRelativePath_NoExtension_ReturnsUnchanged()
        {
            // Unusual case: no dot at all.
            var result = WebView2Bootstrap.ResourceNameToRelativePath("README");
            Assert.Equal("README", result);
        }

        [Theory]
        [InlineData("print-template.html", "print-template.html")]
        [InlineData("print-styles.css", "print-styles.css")]
        public void ResourceNameToRelativePath_HyphenatedPrintResources(string resourceTail, string expectedPath)
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath(resourceTail);
            Assert.Equal(expectedPath, result);
        }

        [Fact]
        public void VirtualHost_IsExpectedValue()
        {
            // The virtual host name is part of the public contract - the
            // chat HTML/JS will resolve resources via this URL.
            Assert.Equal("jameelcopilot.local", WebView2Bootstrap.VirtualHost);
        }

        // Only the add-in's own pages may talk to the host or be navigated to.
        [Theory]
        [InlineData("/index.html")]
        [InlineData("/")]
        public void IsTrustedSource_OwnPages(string path)
        {
            Assert.True(WebView2Bootstrap.IsTrustedSource("https://" + WebView2Bootstrap.VirtualHost + path));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("about:blank")]
        [InlineData("not a uri")]
        [InlineData("https://evil.example/")]
        [InlineData("data:text/html,<p>x</p>")]
        [InlineData("file:///C:/Windows/win.ini")]
        public void IsTrustedSource_RejectsEverythingElse(string uri)
        {
            Assert.False(WebView2Bootstrap.IsTrustedSource(uri));
        }

        [Fact]
        public void IsTrustedSource_RejectsLookalikesAndPlainHttp()
        {
            var host = WebView2Bootstrap.VirtualHost;

            Assert.False(WebView2Bootstrap.IsTrustedSource("http://" + host + "/index.html"));
            Assert.False(WebView2Bootstrap.IsTrustedSource("https://" + host + ".evil.example/"));
            Assert.False(WebView2Bootstrap.IsTrustedSource("https://evil.example/?x=https://" + host + "/"));
            Assert.False(WebView2Bootstrap.IsTrustedSource("https://" + host + "@evil.example/"));
        }

        // The PDF renderer loads its page with NavigateToString; chat views never do.
        [Theory]
        [InlineData("about:blank", false, false)]
        [InlineData("about:blank", true, true)]
        [InlineData("data:text/html;charset=utf-8;base64,PGh0bWw+", false, false)]
        [InlineData("data:text/html;charset=utf-8;base64,PGh0bWw+", true, true)]
        [InlineData("https://evil.example/", true, false)]
        [InlineData("data:application/x-msdownload;base64,AAAA", true, false)]
        public void IsAllowedNavigation_HostHtmlOnlyWhenAsked(string uri, bool allowHostHtml, bool expected)
        {
            Assert.Equal(expected, WebView2Bootstrap.IsAllowedNavigation(uri, allowHostHtml));
        }

        [Fact]
        public void IsAllowedNavigation_OwnPagesAlways()
        {
            Assert.True(WebView2Bootstrap.IsAllowedNavigation("https://" + WebView2Bootstrap.VirtualHost + "/index.html", false));
        }

        [Fact]
        public void PdfWebView2DataFolder_IsSeparateFromMainWebView2DataFolder()
        {
            Assert.NotEqual(WebView2Bootstrap.WebView2DataFolder, WebView2Bootstrap.PdfWebView2DataFolder);
            Assert.EndsWith(Path.Combine("JameelCopilot", "WebView2PdfData"), WebView2Bootstrap.PdfWebView2DataFolder);
        }

        [Fact]
        public void WebUi_LoadsSharedMarkdownBeforeChat()
        {
            var indexHtml = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "index.html"));

            var markdownIndex = indexHtml.IndexOf("<script src=\"markdown.js\"></script>");
            var chatIndex = indexHtml.IndexOf("<script src=\"chat.js\"></script>");

            Assert.True(markdownIndex >= 0, "index.html should load markdown.js.");
            Assert.True(chatIndex >= 0, "index.html should load chat.js.");
            Assert.True(markdownIndex < chatIndex, "markdown.js should load before chat.js.");
        }

        [Fact]
        public void Project_EmbedsSharedMarkdownResource()
        {
            var projectXml = File.ReadAllText(FindSourceFile("JameelCopilot", "JameelCopilot.csproj"));

            Assert.Contains("<EmbeddedResource Include=\"WebUI\\markdown.js\">", projectXml);
            Assert.Contains("<LogicalName>JameelCopilot.WebUI.markdown.js</LogicalName>", projectXml);
        }

        [Fact]
        public void Project_EmbedsPrintTemplateResources()
        {
            var projectXml = File.ReadAllText(FindSourceFile("JameelCopilot", "JameelCopilot.csproj"));

            Assert.Contains("<EmbeddedResource Include=\"WebUI\\print-template.html\">", projectXml);
            Assert.Contains("<LogicalName>JameelCopilot.WebUI.print-template.html</LogicalName>", projectXml);
            Assert.Contains("<EmbeddedResource Include=\"WebUI\\print-styles.css\">", projectXml);
            Assert.Contains("<LogicalName>JameelCopilot.WebUI.print-styles.css</LogicalName>", projectXml);
        }

        [Fact]
        public void PrintTemplate_ReferencesStylesMarkdownAndRendererTokens()
        {
            var templateHtml = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "print-template.html"));

            Assert.Contains("<link rel=\"stylesheet\" href=\"print-styles.css\" nonce=\"__NONCE__\">", templateHtml);
            Assert.Contains("<script src=\"markdown.js\" nonce=\"__NONCE__\"></script>", templateHtml);
            Assert.Contains("<title>__TITLE_TEXT__</title>", templateHtml);
            Assert.Contains("<h1 class=\"doc-title\" dir=\"auto\">__TITLE_TEXT__</h1>", templateHtml);
            Assert.Contains("<p class=\"doc-subtitle\" dir=\"auto\">__SUBTITLE_TEXT__</p>", templateHtml);
            Assert.Contains("Generated by Jameel Copilot &middot; __GENERATED_AT__", templateHtml);
            Assert.Contains("__MD_INJECT__", templateHtml);
            Assert.Contains("window.markdown.render(window.__JAMEELCOPILOT_MD__ || '')", templateHtml);
            Assert.Contains("data-render-state=\"loading\"", templateHtml);
            Assert.Contains("data-render-state", templateHtml);
            Assert.Contains("ready", templateHtml);
            Assert.Contains("error", templateHtml);
        }

        [Fact]
        public void WebUi_ChatRendersFileSavedToolResultsAsFileCard()
        {
            var chatJs = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "chat.js"));
            var stylesCss = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "styles.css"));

            Assert.Contains("appendFileCardToMessage", chatJs);
            Assert.Contains("onFileSaved", chatJs);
            Assert.Contains("onExportError", chatJs);
            Assert.Contains("Unknown export error.", chatJs);
            Assert.Contains("result_type === 'file_saved'", chatJs);
            Assert.Contains("open_file", chatJs);
            Assert.Contains("reveal_in_explorer", chatJs);
            Assert.Contains("renderInlineErrorCard", chatJs);

            Assert.Contains(".msg-attachments", stylesCss);
            Assert.Contains(".file-card", stylesCss);
            Assert.Contains(".file-card-btn", stylesCss);
        }

        [Fact]
        public void WebUi_ChatRendersExportErrorsAsActionableInlineCard()
        {
            var chatJs = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "chat.js"));
            var stylesCss = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "styles.css"));

            Assert.Contains("var retryableExportErrorCodes", chatJs);
            Assert.Contains("file_locked", chatJs);
            Assert.Contains("webview2_missing", chatJs);
            Assert.Contains("path_timeout", chatJs);
            Assert.Contains("pdf_render_timeout", chatJs);
            Assert.Contains("function renderInlineErrorCard(messageId, err)", chatJs);
            Assert.Contains("'Export failed'", chatJs);
            Assert.Contains("detail.textContent = exportErrorMessage(err);", chatJs);
            Assert.Contains("body.appendChild(detail)", chatJs);
            Assert.DoesNotContain("innerHTML = exportErrorMessage", chatJs);
            Assert.Contains("handleExportPdf(retryMessageId);", chatJs);
            Assert.Contains("Ask IT to install the WebView2 runtime.", chatJs);
            Assert.Contains("if (!renderInlineErrorCard(messageId, error))", chatJs);
            Assert.Contains("retryMessageId = msgEl.dataset.messageId;", chatJs);

            Assert.Contains(".error-card", stylesCss);
            Assert.Contains(".error-card-icon", stylesCss);
            Assert.Contains(".error-card-body", stylesCss);
            Assert.Contains(".error-card-title", stylesCss);
            Assert.Contains(".error-card-detail", stylesCss);
            Assert.Contains(".error-card-actions", stylesCss);
            Assert.Contains(".error-card-btn", stylesCss);
            Assert.Contains(".error-card-link", stylesCss);
        }

        [Fact]
        public void WebUi_ChatRoutesPdfExportToolErrorsToInlineErrorCards()
        {
            var chatJs = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "chat.js"));

            Assert.Contains("row.dataset.toolName = String(name || '');", chatJs);
            Assert.Contains("function normalizePdfExportToolError(resultJson)", chatJs);
            Assert.Contains("obj.error.code", chatJs);
            Assert.Contains("obj.error.message", chatJs);
            Assert.Contains("row.dataset.toolName === 'outlook_export_pdf'", chatJs);
            Assert.Contains("renderInlineErrorCard(null, pdfExportError)", chatJs);
            Assert.Contains("if (row.parentNode) row.parentNode.removeChild(row);", chatJs);
            Assert.Contains("row.classList.add('tool-status-err');", chatJs);
            Assert.Contains("row.textContent = '\\u26A0 ' + (row.dataset.verb || 'A step') + ' did not work; Jameel Copilot carried on without it.';", chatJs);
            Assert.DoesNotContain("(summary || 'tool error')", chatJs);   // a tool's raw error is for the model, not the user
        }

        [Fact]
        public void WebUi_ChatSupportsPerMessagePdfExport()
        {
            var chatJs = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "chat.js"));
            var stylesCss = File.ReadAllText(FindSourceFile("JameelCopilot", "WebUI", "styles.css"));

            Assert.Contains("handleExportPdf", chatJs);
            Assert.Contains("deriveFilenameHint", chatJs);
            Assert.Contains("export_pdf", chatJs);
            Assert.Contains("content_markdown", chatJs);
            Assert.Contains("filename_hint", chatJs);
            Assert.Contains("Save message as PDF", chatJs);
            Assert.Contains("entry.raw", chatJs);
            Assert.Contains("setExportButtonPending", chatJs);
            Assert.Contains("resetExportButton", chatJs);
            Assert.Contains("onFileSaved", chatJs);
            Assert.Contains("onExportError", chatJs);
            Assert.Contains("heading.replace(/\\s+#+\\s*$/, '')", chatJs);
            Assert.DoesNotContain("^#{1,3}\\s+(.+?)\\s*#*\\s*$", chatJs);

            Assert.Contains(".msg-action-pdf", stylesCss);
            Assert.Contains(".msg-assistant:hover .msg-action", stylesCss);
            Assert.Contains(".msg-assistant.is-streaming .msg-action", stylesCss);
            Assert.Contains(".msg-action:disabled", stylesCss);
        }

        private static string FindSourceFile(params string[] relativeParts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(relativeParts));
                if (File.Exists(candidate)) return candidate;

                current = current.Parent;
            }

            throw new FileNotFoundException("Could not find source file.", Path.Combine(relativeParts));
        }
    }
}
