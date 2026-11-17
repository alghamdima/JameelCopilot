using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JameelCopilot.Services.Tools
{
    /// <summary>
    /// Extracts plain text from email attachments (.docx, .xlsx, .pptx, .pdf,
    /// .txt, .csv, .md, .json, .xml, .log, .html) saved to a temporary file.
    /// Pure file-to-text helper so it can be unit-tested without Outlook COM.
    /// </summary>
    public static class AttachmentTextExtractor
    {
        public const int MaxExtractedChars = 8000;
        public const long MaxAttachmentBytes = 15 * 1024 * 1024; // 15 MB

        private static readonly string[] PlainTextExtensions =
        {
            ".txt", ".csv", ".tsv", ".md", ".json", ".xml", ".log", ".sql", ".yaml", ".yml", ".ini"
        };

        public static bool IsSupportedExtension(string filename)
        {
            var ext = (Path.GetExtension(filename ?? "") ?? "").ToLowerInvariant();
            return ext == ".docx" || ext == ".xlsx" || ext == ".pptx" || ext == ".pdf"
                || ext == ".html" || ext == ".htm"
                || PlainTextExtensions.Contains(ext);
        }

        public static string ExtractFromFile(string filePath, string originalFilename = null, int maxChars = MaxExtractedChars)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return null;
            var ext = (Path.GetExtension(originalFilename ?? filePath) ?? "").ToLowerInvariant();
            try
            {
                string raw;
                switch (ext)
                {
                    case ".docx":
                        raw = ExtractDocx(filePath);
                        break;
                    case ".xlsx":
                        raw = ExtractXlsx(filePath);
                        break;
                    case ".pptx":
                        raw = ExtractPptx(filePath);
                        break;
                    case ".pdf":
                        raw = ExtractPdf(filePath);
                        break;
                    case ".html":
                    case ".htm":
                        raw = StripHtml(File.ReadAllText(filePath, Encoding.UTF8));
                        break;
                    default:
                        if (!PlainTextExtensions.Contains(ext)) return null;
                        raw = File.ReadAllText(filePath, Encoding.UTF8);
                        break;
                }
                return Clamp(raw, maxChars);
            }
            catch
            {
                return null;
            }
        }

        private static string ExtractDocx(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, false))
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body == null) return "";
                var sb = new StringBuilder();
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph p)
                    {
                        var text = p.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(text)) sb.AppendLine(text);
                    }
                    else if (element is Table t)
                    {
                        foreach (var row in t.Elements<TableRow>())
                        {
                            var cells = row.Elements<TableCell>()
                                .Select(c => (c.InnerText ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim());
                            sb.AppendLine(string.Join(" | ", cells));
                        }
                    }
                }
                return sb.ToString().Trim();
            }
        }

        private static string ExtractXlsx(string filePath)
        {
            using (var wb = new XLWorkbook(filePath))
            {
                var sb = new StringBuilder();
                foreach (var ws in wb.Worksheets.Take(8))
                {
                    var range = ws.RangeUsed();
                    if (range == null) continue;
                    sb.AppendLine("[Sheet: " + ws.Name + "]");
                    int rowCount = 0;
                    foreach (var row in range.RowsUsed())
                    {
                        if (++rowCount > 120)
                        {
                            sb.AppendLine("... (rows truncated)");
                            break;
                        }
                        var cells = row.Cells(1, Math.Min(25, range.ColumnCount()))
                            .Select(c => (c.GetFormattedString() ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim());
                        var line = string.Join(" | ", cells).Trim();
                        if (line.Replace("|", "").Trim().Length > 0) sb.AppendLine(line);
                    }
                    sb.AppendLine();
                    if (sb.Length >= MaxExtractedChars) break;
                }
                return sb.ToString().Trim();
            }
        }

        private static string ExtractPptx(string filePath)
        {
            using (var archive = ZipFile.OpenRead(filePath))
            {
                var slides = archive.Entries
                    .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                             && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(e => SlideNumber(e.Name))
                    .Take(50)
                    .ToList();

                var sb = new StringBuilder();
                int index = 0;
                foreach (var entry in slides)
                {
                    index++;
                    using (var stream = entry.Open())
                    {
                        var xdoc = XDocument.Load(stream);
                        var texts = xdoc.Descendants()
                            .Where(el => el.Name.LocalName == "t")
                            .Select(el => (el.Value ?? "").Trim())
                            .Where(t => t.Length > 0)
                            .ToList();
                        if (texts.Count > 0)
                        {
                            sb.AppendLine("[Slide " + index + "] " + string.Join(" ", texts));
                        }
                    }
                    if (sb.Length >= MaxExtractedChars) break;
                }
                return sb.ToString().Trim();
            }
        }

        private static int SlideNumber(string name)
        {
            var m = Regex.Match(name ?? "", @"\d+");
            return m.Success && int.TryParse(m.Value, out var n) ? n : 999;
        }

        private static string ExtractPdf(string filePath)
        {
            var bytes = File.ReadAllBytes(filePath);
            var latin = Encoding.GetEncoding("ISO-8859-1").GetString(bytes);
            var cmap = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var streamMatches = Regex.Matches(latin, @"stream\r?\n([\s\S]*?)\r?\nendstream");
            var decompressedStreams = new System.Collections.Generic.List<string>();
            foreach (Match m in streamMatches)
            {
                var content = DecompressPdfStream(bytes, m.Groups[1].Index, m.Groups[1].Length) ?? m.Groups[1].Value;
                decompressedStreams.Add(content);
                ParseToUnicodeCMap(content, cmap);
            }

            var sb = new StringBuilder();
            foreach (var content in decompressedStreams)
            {
                if (sb.Length >= MaxExtractedChars) break;
                ExtractPdfTextOperators(content, cmap, sb);
            }
            if (sb.Length == 0)
            {
                ExtractPdfTextOperators(latin, cmap, sb);
            }
            return sb.ToString().Trim();
        }

        private static void ParseToUnicodeCMap(string content, System.Collections.Generic.Dictionary<string, string> cmap)
        {
            if (string.IsNullOrEmpty(content) || content.IndexOf("beginbf", StringComparison.Ordinal) < 0) return;
            foreach (Match m in Regex.Matches(content, @"<([0-9A-Fa-f]{2,4})>\s*<([0-9A-Fa-f]{4,16})>"))
            {
                var src = m.Groups[1].Value.ToUpperInvariant();
                var dst = DecodeUtf16Hex(m.Groups[2].Value);
                if (!string.IsNullOrEmpty(dst)) cmap[src] = dst;
            }
            foreach (Match m in Regex.Matches(content, @"<([0-9A-Fa-f]{2,4})>\s*<([0-9A-Fa-f]{2,4})>\s*<([0-9A-Fa-f]{4})>"))
            {
                if (int.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.HexNumber, null, out var start)
                    && int.TryParse(m.Groups[2].Value, System.Globalization.NumberStyles.HexNumber, null, out var end)
                    && int.TryParse(m.Groups[3].Value, System.Globalization.NumberStyles.HexNumber, null, out var target)
                    && end >= start && (end - start) <= 256)
                {
                    int width = m.Groups[1].Value.Length;
                    for (int code = start; code <= end; code++)
                    {
                        cmap[code.ToString("X" + width)] = ((char)(target + (code - start))).ToString();
                    }
                }
            }
        }

        private static string DecodeUtf16Hex(string hex)
        {
            if (string.IsNullOrEmpty(hex) || (hex.Length % 4) != 0) return "";
            var sb = new StringBuilder(hex.Length / 4);
            for (int i = 0; i < hex.Length; i += 4)
            {
                if (int.TryParse(hex.Substring(i, 4), System.Globalization.NumberStyles.HexNumber, null, out var code) && code >= 32)
                {
                    sb.Append((char)code);
                }
            }
            return sb.ToString();
        }

        private static string DecodePdfHexToken(string hex, System.Collections.Generic.Dictionary<string, string> cmap)
        {
            var clean = Regex.Replace(hex ?? "", @"\s+", "").ToUpperInvariant();
            if (clean.Length == 0) return "";
            var sb = new StringBuilder();
            if ((clean.Length % 4) == 0 && cmap.Keys.Any(k => k.Length == 4))
            {
                for (int i = 0; i < clean.Length; i += 4)
                {
                    var key = clean.Substring(i, 4);
                    if (cmap.TryGetValue(key, out var mapped)) sb.Append(mapped);
                    else if (int.TryParse(key, System.Globalization.NumberStyles.HexNumber, null, out var ch) && ch >= 32 && ch < 0xD800) sb.Append((char)ch);
                }
                return sb.ToString();
            }
            for (int i = 0; i + 1 < clean.Length; i += 2)
            {
                var key = clean.Substring(i, 2);
                if (cmap.TryGetValue(key, out var mapped)) sb.Append(mapped);
                else if (int.TryParse(key, System.Globalization.NumberStyles.HexNumber, null, out var ch) && ch >= 32 && ch <= 126) sb.Append((char)ch);
            }
            return sb.ToString();
        }

        private static string DecompressPdfStream(byte[] fullBytes, int offset, int length)
        {
            if (length <= 6 || offset + length > fullBytes.Length) return null;
            try
            {
                // Skip 2-byte zlib header (0x78 0x9C / 0x01 / 0xDA) for raw DeflateStream.
                int start = (fullBytes[offset] == 0x78) ? offset + 2 : offset;
                int count = length - (start - offset);
                using (var ms = new MemoryStream(fullBytes, start, count))
                using (var deflate = new DeflateStream(ms, CompressionMode.Decompress))
                using (var reader = new StreamReader(deflate, Encoding.GetEncoding("ISO-8859-1")))
                {
                    return reader.ReadToEnd();
                }
            }
            catch
            {
                return null;
            }
        }

        private static void ExtractPdfTextOperators(string content, System.Collections.Generic.Dictionary<string, string> cmap, StringBuilder sb)
        {
            foreach (Match tm in Regex.Matches(content ?? "", @"\(([^()\\]*(?:\\.[^()\\]*)*)\)\s*(?:Tj|')|<([0-9A-Fa-f\s]+)>\s*(?:Tj|')|\[((?:[^\[\]]|\([^()]*\)|<[0-9A-Fa-f\s]*>)*)\]\s*TJ"))
            {
                if (tm.Groups[1].Success)
                {
                    var s = UnescapePdfString(tm.Groups[1].Value);
                    if (!string.IsNullOrWhiteSpace(s)) sb.Append(s).Append(' ');
                }
                else if (tm.Groups[2].Success)
                {
                    var s = DecodePdfHexToken(tm.Groups[2].Value, cmap);
                    if (!string.IsNullOrWhiteSpace(s)) sb.Append(s).Append(' ');
                }
                else if (tm.Groups[3].Success)
                {
                    foreach (Match sub in Regex.Matches(tm.Groups[3].Value, @"\(([^()\\]*(?:\\.[^()\\]*)*)\)|<([0-9A-Fa-f\s]+)>"))
                    {
                        var s = sub.Groups[1].Success
                            ? UnescapePdfString(sub.Groups[1].Value)
                            : DecodePdfHexToken(sub.Groups[2].Value, cmap);
                        if (!string.IsNullOrEmpty(s)) sb.Append(s);
                    }
                    sb.Append(' ');
                }
                if (sb.Length >= MaxExtractedChars) break;
            }
        }

        private static string UnescapePdfString(string s)
        {
            return (s ?? "")
                .Replace(@"\n", "\n")
                .Replace(@"\r", " ")
                .Replace(@"\t", " ")
                .Replace(@"\(", "(")
                .Replace(@"\)", ")")
                .Replace(@"\\", "\\");
        }

        private static string StripHtml(string html)
        {
            var noTags = Regex.Replace(html ?? "", @"<[^>]+>", " ");
            return Regex.Replace(System.Net.WebUtility.HtmlDecode(noTags), @"\s+", " ").Trim();
        }

        private static string Clamp(string text, int maxChars)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var clean = text.Trim();
            return clean.Length <= maxChars ? clean : clean.Substring(0, maxChars) + "\n[... attachment text truncated ...]";
        }
    }
}
