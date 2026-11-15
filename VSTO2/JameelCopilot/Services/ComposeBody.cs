using System.Text;
using System.Text.RegularExpressions;

namespace JameelCopilot.Services
{
    /// <summary>
    /// The text of a compose window, split by who wrote it. Only
    /// <see cref="Own"/> is the user's email; <see cref="Quoted"/> is the
    /// message being replied to or forwarded, with its header.
    /// </summary>
    internal sealed class ComposeBody
    {
        // Where a quoted message starts in plain text: Outlook's
        // "-----Original Message-----" line, its rule of underscores, or a
        // "From:" line (من is the Arabic one) followed by another
        // "Label: value" line.
        private static readonly Regex QuotedStart = new Regex(
            @"^[ \t]*(-{3,}[^-\r\n]+-{3,}[ \t]*\r?$|_{10,}[ \t]*\r?$|(From|من):[^\r\n]*\r?\n[ \t]*[^\s:][^:\r\n]{0,30}:[ \t])",
            RegexOptions.Multiline | RegexOptions.CultureInvariant);

        public ComposeBody(string own, string quoted)
        {
            Own = own ?? "";
            Quoted = quoted ?? "";
        }

        public string Own { get; }
        public string Quoted { get; }

        /// <summary>Where the quoted message starts in <paramref name="text"/>, or -1.</summary>
        public static int QuotedIndex(string text)
        {
            var match = QuotedStart.Match(text ?? "");
            return match.Success ? match.Index : -1;
        }

        public static ComposeBody FromPlainText(string body)
        {
            body = body ?? "";
            int at = QuotedIndex(body);
            return at < 0
                ? new ComposeBody(body.Trim(), "")
                : new ComposeBody(body.Substring(0, at).Trim(), body.Substring(at).Trim());
        }

        /// <summary>The text of a Word range as plain lines.</summary>
        public static string FromWord(string rangeText)
        {
            var plain = new StringBuilder();
            foreach (var c in rangeText ?? "")
            {
                if (c == '\r' || c == '\v' || c == '\f') plain.Append('\n');
                else if (c == '\t' || c >= ' ') plain.Append(c);
                // The rest are Word's own marks: the end of a table cell, a picture, a field.
            }
            return plain.ToString().Trim();
        }

        /// <summary>Plain text as Word takes it: \r ends a paragraph, and \r\n would end two.</summary>
        public static string ToWord(string text)
        {
            return (text ?? "").Replace("\r\n", "\r").Replace('\n', '\r').Trim();
        }
    }
}
