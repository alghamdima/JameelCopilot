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
            if (at < 0)
            {
                return new ComposeBody(StripSignatureAndFooter(body), "");
            }
            return new ComposeBody(
                StripSignatureAndFooter(body.Substring(0, at)),
                StripThreadSignatures(body.Substring(at)));
        }

        // Common sign-offs that precede a manual or corporate email signature block.
        private static readonly Regex SignOffLine = new Regex(
            @"^[ \t]*(Best[ \t]+regards|Kind[ \t]+regards|Warm[ \t]+regards|Regards|Sincerely|Many[ \t]+thanks|Thanks[ \t]*&[ \t]*regards|Thank[ \t]+you|Thanks|BR|RGDS|مع[ \t]+خالص[ \t]+التحية|مع[ \t]+التحية|وتقبلوا[ \t]+فائق[ \t]+الاحترام|تحياتي|شكراً[ \t]+لكم|وشكراً)[ \t,!.:،]*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex ClassificationOrDisclaimerLine = new Regex(
            @"^[ \t]*(This email is classified as\b|Classification:|CONFIDENTIALITY NOTICE|DISCLAIMER:|This message and any attachments|هذا البريد الإلكتروني مصنف|للإستخدام الداخلي|للاستخدام الداخلي)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex SignatureContactMarker = new Regex(
            @"(\+\d{2,3}[\s\-]?\d|@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}|PO[ \t]*Box|P\.O\.[ \t]*Box|Postal[ \t]*Code|Kingdom of Saudi Arabia|Abdul Latif Jameel|aljfinance\.com|aljfs\.com|\bTel\b|\bMob\b|\bExt\b)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Strips trailing manual/corporate signature blocks and classification
        /// footers from a single message's own text so the model never sees,
        /// translates, or duplicates signatures.
        /// </summary>
        public static string StripSignatureAndFooter(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            int cutAt = lines.Length;

            // 1. Strip trailing classification / disclaimer footers first.
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                if (ClassificationOrDisclaimerLine.IsMatch(lines[i]))
                {
                    cutAt = i;
                }
            }

            // 2. Look for "-- " standard delimiter or a sign-off line followed by contact/signature markers,
            //    or a trailing corporate contact block when the user did not write "Regards".
            for (int i = 0; i < cutAt; i++)
            {
                var trimmed = lines[i].Trim();
                if (lines[i] == "-- " || trimmed == "--")
                {
                    cutAt = i;
                    break;
                }
                if (SignOffLine.IsMatch(trimmed))
                {
                    // Check if what follows is a real signature block (name + title/contact markers), not just a standalone "Thanks" at the end of a draft.
                    int remainingNonEmpty = 0;
                    bool hasMarker = false;
                    for (int j = i + 1; j < cutAt; j++)
                    {
                        var sub = lines[j].Trim();
                        if (sub.Length == 0) continue;
                        remainingNonEmpty++;
                        if (SignatureContactMarker.IsMatch(sub)) hasMarker = true;
                    }
                    if (hasMarker || (remainingNonEmpty >= 1 && remainingNonEmpty <= 6))
                    {
                        cutAt = i;
                        break;
                    }
                }
            }

            // 3. Even without a sign-off word (e.g. a compose window with a manual signature pasted at the bottom),
            //    detect a trailing corporate signature block containing phone/PO Box/company markers in the last 14 lines.
            if (cutAt == lines.Length || cutAt > 0)
            {
                int searchStart = System.Math.Max(0, cutAt - 14);
                int firstMarker = -1;
                for (int i = searchStart; i < cutAt; i++)
                {
                    if (SignatureContactMarker.IsMatch(lines[i]))
                    {
                        firstMarker = i;
                        break;
                    }
                }
                if (firstMarker >= 0)
                {
                    // Walk up 1-3 non-empty short lines above the first phone/email/company marker (name, job title, department).
                    int sigStart = firstMarker;
                    int walkedUp = 0;
                    for (int k = firstMarker - 1; k >= 0 && walkedUp < 3; k--)
                    {
                        var prev = lines[k].Trim();
                        if (prev.Length == 0)
                        {
                            sigStart = k;
                            break;
                        }
                        if (prev.Length <= 55 && !prev.EndsWith(".") && !prev.EndsWith("?") && !prev.EndsWith("!"))
                        {
                            sigStart = k;
                            walkedUp++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (sigStart < cutAt) cutAt = sigStart;
                }
            }

            var sb = new StringBuilder();
            for (int i = 0; i < cutAt; i++)
            {
                if (i > 0) sb.Append("\r\n");
                sb.Append(lines[i]);
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// Splits an entire email thread into individual messages and strips the
        /// signature and classification footer from EVERY message in the chain,
        /// drastically shrinking large threads while keeping every message's actual content.
        /// </summary>
        public static string StripThreadSignatures(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return "";
            var normalized = body.Replace("\r\n", "\n").Replace('\r', '\n');
            var matches = QuotedStart.Matches(normalized);
            if (matches.Count == 0)
            {
                return StripSignatureAndFooter(normalized);
            }

            var sb = new StringBuilder();
            int pos = 0;
            for (int m = 0; m < matches.Count; m++)
            {
                int start = matches[m].Index;
                var matchVal = matches[m].Value.TrimStart();
                bool isSeparatorOnly = (matchVal.StartsWith("-") || matchVal.StartsWith("_"))
                    && (m + 1 < matches.Count)
                    && string.IsNullOrWhiteSpace(normalized.Substring(start + matches[m].Length, matches[m + 1].Index - (start + matches[m].Length)));

                if (start > pos)
                {
                    var segment = StripSignatureAndFooter(normalized.Substring(pos, start - pos));
                    if (segment.Length > 0)
                    {
                        if (sb.Length > 0) sb.Append("\r\n\r\n");
                        sb.Append(segment);
                    }
                }

                if (isSeparatorOnly)
                {
                    if (sb.Length > 0) sb.Append("\r\n\r\n");
                    sb.Append(matches[m].Value.Trim());
                    m++;
                    start = matches[m].Index;
                }

                int nextStart = (m + 1 < matches.Count) ? matches[m + 1].Index : normalized.Length;
                var quotedBlock = normalized.Substring(start, nextStart - start);

                // Keep the header block (From/Sent/To/Cc/Subject lines up to first blank line), then strip signature from its body.
                int headerEnd = quotedBlock.IndexOf("\n\n", System.StringComparison.Ordinal);
                if (headerEnd > 0)
                {
                    var header = quotedBlock.Substring(0, headerEnd).Trim().Replace("\n", "\r\n");
                    var msgBody = StripSignatureAndFooter(quotedBlock.Substring(headerEnd + 2));
                    if (sb.Length > 0 && !isSeparatorOnly) sb.Append("\r\n\r\n");
                    else if (isSeparatorOnly) sb.Append("\r\n");
                    sb.Append(header);
                    if (msgBody.Length > 0) sb.Append("\r\n\r\n").Append(msgBody);
                }
                else
                {
                    var cleaned = StripSignatureAndFooter(quotedBlock);
                    if (cleaned.Length > 0)
                    {
                        if (sb.Length > 0 && !isSeparatorOnly) sb.Append("\r\n\r\n");
                        else if (isSeparatorOnly) sb.Append("\r\n");
                        sb.Append(cleaned);
                    }
                }
                pos = nextStart;
            }
            return sb.ToString().Trim();
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
