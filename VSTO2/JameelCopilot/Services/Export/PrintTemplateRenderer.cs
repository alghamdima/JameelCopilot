using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace JameelCopilot.Services.Export
{
    public sealed class PrintTemplateRenderer
    {
        private static readonly Regex ImageRegex = new Regex(@"!\[([^\]]*)\]\([^)]+\)", RegexOptions.Compiled);
        private static readonly Regex TemplateTokenRegex = new Regex(@"__TITLE_TEXT__|__SUBTITLE_TEXT__|__GENERATED_AT__|__MD_INJECT__|__NONCE__", RegexOptions.Compiled);
        private readonly string templateHtml;

        public PrintTemplateRenderer(string templateHtml)
        {
            if (templateHtml == null) throw new ArgumentNullException(nameof(templateHtml));
            this.templateHtml = templateHtml;
        }

        public static PrintTemplateRenderer LoadFromFile(string templatePath)
        {
            return new PrintTemplateRenderer(File.ReadAllText(templatePath));
        }

        public string Render(string title, string subtitle, string markdown, DateTimeOffset generatedAt)
        {
            var safeMarkdown = StripInlineImagesOutsideCodeBlocks(markdown ?? "");
            var markdownLiteral = JsonConvert.ToString(safeMarkdown, '"', StringEscapeHandling.EscapeHtml);
            var markdownInjection = "window.__JAMEELCOPILOT_MD__ = " + markdownLiteral + ";";
            var titleText = WebUtility.HtmlEncode(title ?? "");
            var subtitleText = WebUtility.HtmlEncode(subtitle ?? "");
            // The page's policy allows only scripts and styles carrying this render's nonce.
            var nonce = NewNonce();
            var generatedAtText = generatedAt.ToLocalTime().ToString("MMMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);

            return TemplateTokenRegex.Replace(templateHtml, match =>
            {
                switch (match.Value)
                {
                    case "__TITLE_TEXT__":
                        return titleText;
                    case "__SUBTITLE_TEXT__":
                        return subtitleText;
                    case "__GENERATED_AT__":
                        return generatedAtText;
                    case "__MD_INJECT__":
                        return markdownInjection;
                    case "__NONCE__":
                        return nonce;
                    default:
                        return match.Value;
                }
            });
        }

        private static string NewNonce()
        {
            var bytes = new byte[16];
            using (var random = System.Security.Cryptography.RandomNumberGenerator.Create())
            {
                random.GetBytes(bytes);
            }
            return Convert.ToBase64String(bytes);
        }

        private static string StripInlineImagesOutsideCodeBlocks(string markdown)
        {
            var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var inFence = false;
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    inFence = !inFence;
                }
                else if (!inFence)
                {
                    lines[i] = ImageRegex.Replace(lines[i], "[image: $1]");
                }
            }

            return string.Join("\n", lines);
        }
    }
}
