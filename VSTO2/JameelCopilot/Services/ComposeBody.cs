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
                    // Check if what follows is a real signature block (name + title/contact markers), not a P.S. paragraph.
                    int remainingNonEmpty = 0;
                    bool hasMarker = false;
                    bool allShortUnpunctuated = true;
                    for (int j = i + 1; j < cutAt; j++)
                    {
                        var sub = lines[j].Trim();
                        if (sub.Length == 0) continue;
                        remainingNonEmpty++;
                        if (SignatureContactMarker.IsMatch(sub)) hasMarker = true;
                        if (sub.Length > 55 || sub.EndsWith(".") || sub.EndsWith("?") || sub.EndsWith("!")) allShortUnpunctuated = false;
                    }
                    if (hasMarker || (allShortUnpunctuated && remainingNonEmpty >= 1 && remainingNonEmpty <= 4))
                    {
                        cutAt = i;
                        break;
                    }
                }
            }

            // 3. Even without a sign-off word (e.g. a compose window with a manual signature pasted at the bottom),
            //    detect a trailing corporate signature block when at least 2 distinct signature/contact lines sit at the bottom
            //    (so mentioning a colleague's email or phone in a normal sentence is never stripped).
            if (cutAt > 0)
            {
                int searchStart = System.Math.Max(0, cutAt - 14);
                int firstMarker = -1;
                int markerCount = 0;
                for (int i = searchStart; i < cutAt; i++)
                {
                    var candidateLine = lines[i].Trim();
                    if (candidateLine.Length == 0 || candidateLine.Length > 95) continue;
                    if (candidateLine.EndsWith("?") || candidateLine.EndsWith("!")) continue;
                    if (SignatureContactMarker.IsMatch(candidateLine))
                    {
                        if (firstMarker < 0) firstMarker = i;
                        markerCount++;
                    }
                }
                if (firstMarker >= 0 && markerCount >= 2)
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

        private static readonly Regex GreetingOrBodyStart = new Regex(
            @"^[ \t]*(Hi\b|Hello\b|Dear\b|Good\s+(morning|afternoon|evening)\b|Greetings\b|Team[,:]|All[,:]|" +
            @"السلام\s+عليكم|مرحبا|مرحباً|أهلا|أهلاً|تحية\s+طيبة|سعادة\b|الأستاذ|الأخ\b|الأخت\b|الإخوة|الزملاء)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex ConversationalWrapperLine = new Regex(
            @"(:\s*$|^---\s*$|هذا\s+نص|إليك\s+نص|الرد\s+المقترح|ما\s+قدرت\s+أجد|لو\s+تبي\s+أعدل|إذا\s+تبي\s+أعدل|هل\s+ترغب\s+في\s+تعديل|let\s+me\s+know\s+if\s+you(\s+want|\s+would\s+like)|here\s+is\s+(the|a|your)\s+(draft|reply|email))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Cleans model-generated draft text before it is inserted into an
        /// Outlook email editor: strips conversational chat wrappers (above/below
        /// "---" or before the greeting), strips a leading "Subject:" line
        /// (putting it into <paramref name="extractedSubject"/>), strips trailing
        /// sign-offs ("Best regards,"), and removes raw Markdown bold/heading
        /// markers (**text**) so Word receives clean plain text.
        /// </summary>
        public static string CleanDraftForEmail(string text, out string extractedSubject)
        {
            extractedSubject = null;
            if (string.IsNullOrWhiteSpace(text)) return "";

            var rawLines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var lines = new System.Collections.Generic.List<string>(rawLines);

            // 1. If the model fenced the email inside --- ... ---, extract the inner block.
            int firstHr = -1, lastHr = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].Trim() == "---")
                {
                    if (firstHr < 0) firstHr = i;
                    else lastHr = i;
                }
            }
            if (firstHr >= 0 && lastHr > firstHr)
            {
                lines = lines.GetRange(firstHr + 1, lastHr - firstHr - 1);
            }
            else if (firstHr >= 0 && firstHr <= 3 && firstHr + 1 < lines.Count)
            {
                lines = lines.GetRange(firstHr + 1, lines.Count - (firstHr + 1));
            }

            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0])) lines.RemoveAt(0);
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1])) lines.RemoveAt(lines.Count - 1);

            // 2. If there is a conversational intro line (e.g. "... لكن هذا نص الإيميل:") followed by a greeting ("Hi ...", "Dear ..."), drop the intro.
            if (lines.Count >= 2 && !GreetingOrBodyStart.IsMatch(lines[0]))
            {
                for (int i = 1; i <= System.Math.Min(3, lines.Count - 1); i++)
                {
                    if (GreetingOrBodyStart.IsMatch(lines[i]) && ConversationalWrapperLine.IsMatch(lines[0]))
                    {
                        lines.RemoveRange(0, i);
                        break;
                    }
                }
            }

            // 3. Extract and strip a leading "Subject: ..." / "الموضوع: ..." line.
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0])) lines.RemoveAt(0);
            if (lines.Count > 0)
            {
                var first = Regex.Replace(lines[0].Trim(), @"^\*\*|\*\*$", "").Trim();
                var subjMatch = Regex.Match(first, @"^(Subject|الموضوع)\s*:\s*(.+)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (subjMatch.Success)
                {
                    extractedSubject = subjMatch.Groups[2].Value.Trim().Trim('*').Trim();
                    lines.RemoveAt(0);
                    while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0])) lines.RemoveAt(0);
                }
            }

            // 4. Drop trailing "---" and conversational offer lines ("لو تبي أعدل أي شي...") or trailing sign-offs ("Best regards,").
            while (lines.Count > 0)
            {
                var last = lines[lines.Count - 1].Trim();
                if (last.Length == 0 || last == "---" || ConversationalWrapperLine.IsMatch(last) || SignOffLine.IsMatch(last))
                {
                    lines.RemoveAt(lines.Count - 1);
                    continue;
                }
                break;
            }

            // 5. Strip any remaining trailing signature + sign-off block.
            var joined = StripSignatureAndFooter(string.Join("\n", lines));
            var finalLines = new System.Collections.Generic.List<string>(joined.Replace("\r\n", "\n").Split('\n'));
            while (finalLines.Count > 0 && (string.IsNullOrWhiteSpace(finalLines[finalLines.Count - 1]) || SignOffLine.IsMatch(finalLines[finalLines.Count - 1].Trim())))
            {
                finalLines.RemoveAt(finalLines.Count - 1);
            }

            // 6. Strip raw Markdown syntax (**bold**, ### headings) so Word editor gets clean plain text.
            var sb = new StringBuilder();
            for (int i = 0; i < finalLines.Count; i++)
            {
                var ln = finalLines[i];
                ln = Regex.Replace(ln, @"^[ \t]*#{1,6}[ \t]+", "");
                ln = Regex.Replace(ln, @"\*\*([^*\n]+)\*\*", "$1");
                ln = Regex.Replace(ln, @"__([^_\n]+)__", "$1");
                if (i > 0) sb.Append("\r\n");
                sb.Append(ln);
            }
            return sb.ToString().Trim();
        }
    }
}
