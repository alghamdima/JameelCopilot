using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Prefs;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Requests whose answer is data for Jameel Copilot's screen (a ranked list, review
    /// points, proposed actions): the structured prompt core, no tools, and
    /// the helpers to put emails into the prompt and read the answer back.
    /// </summary>
    internal static class StructuredAnswers
    {
        public static ConversationContext NewContext(string instructions, string context, string memory)
        {
            return new ConversationContext
            {
                Structured = true,
                UseTools = false,
                SystemInstructions = instructions ?? "",
                Context = context ?? "",
                Memory = memory ?? "",
                // Sorting and judging want the steadiest answer the administrator allows.
                Temperature = ResponseStyles.TemperatureFor(ResponseStyles.Precise),
            };
        }

        /// <summary>The first JSON object in the answer, or null. Tolerates a code fence or words around it.</summary>
        public static JObject ParseObject(string answer)
        {
            var text = answer ?? "";
            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try
            {
                return StoredJson.Parse(text.Substring(start, end - start + 1));
            }
            catch (JsonException)
            {
                // Stray braces in preamble/trailing text: try every balanced '{'...' }' span.
                for (int i = start; i >= 0 && i < text.Length; i = text.IndexOf('{', i + 1))
                {
                    int depth = 0;
                    bool inStr = false, esc = false;
                    for (int j = i; j < text.Length; j++)
                    {
                        char c = text[j];
                        if (inStr)
                        {
                            if (esc) esc = false;
                            else if (c == '\\') esc = true;
                            else if (c == '"') inStr = false;
                            continue;
                        }
                        if (c == '"') inStr = true;
                        else if (c == '{') depth++;
                        else if (c == '}')
                        {
                            depth--;
                            if (depth == 0)
                            {
                                try
                                {
                                    var parsed = StoredJson.Parse(text.Substring(i, j - i + 1));
                                    if (parsed != null) return parsed;
                                }
                                catch (JsonException) { }
                                break;
                            }
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>Marks a user line Jameel Copilot wrote itself for a chip: it says nothing about the language the user writes in.</summary>
        public const string ChipLineMark = "chip";

        /// <summary>The language the user writes in this chat; with nothing said yet, the language of the emails.</summary>
        public static string ReplyLanguage(IReadOnlyList<JObject> history, string whenNothingSaid = "the language most of these emails are written in")
        {
            var lastUser = (history ?? new JObject[0]).LastOrDefault(m =>
                (string)m?["type"] == "message" && (string)m?["role"] == "user" && m["content"]?.Type == JTokenType.String
                && m[ChipLineMark] == null);
            if (lastUser == null) return whenNothingSaid;
            return CustomApiChatService.DetectPromptLanguage((string)lastUser["content"]) ?? "English";
        }

        /// <summary>
        /// Emails as the prompt carries them: one header line each (sender-written
        /// fields kept on it) and the snippet behind a "> " prefix.
        /// </summary>
        public static string MessageList(IEnumerable<MessageSummary> messages)
        {
            var sb = new StringBuilder();
            foreach (var m in messages ?? new MessageSummary[0])
            {
                sb.Append("[id ").Append(OneLine(m.Id)).Append("] From: ").Append(OneLine(m.From))
                  .Append(" | Received: ").Append(m.ReceivedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture))
                  .Append(m.HasAttachments ? " | has attachments" : "")
                  .Append(" | Subject: ").Append(OneLine(m.Subject)).Append('\n');
                var snippet = OneLine(m.Snippet);
                if (snippet.Length > 0) sb.Append("> ").Append(snippet.Length > 300 ? snippet.Substring(0, 300) : snippet).Append('\n');
            }
            return sb.ToString().TrimEnd('\n');
        }

        /// <summary>Someone else's text as the prompt carries it: every line behind "> ", cut at <paramref name="max"/>.</summary>
        public static string Quote(string text, int max)
        {
            var body = (text ?? "").Trim();
            if (body.Length > max) body = body.Substring(0, max) + "\n[... the rest is cut ...]";
            return "> " + string.Join("\n> ", body.Replace("\r\n", "\n").Split('\r', '\n', '\v', '\f', (char)0x85, (char)0x2028, (char)0x2029));
        }

        /// <summary>Text on one line: every kind of line break becomes a space.</summary>
        public static string OneLine(string text)
        {
            var sb = new StringBuilder();
            foreach (var c in text ?? "")
            {
                sb.Append(c == '\r' || c == '\n' || c == '\v' || c == '\f' || c == '\u0085' || c == '\u2028' || c == '\u2029' ? ' ' : c);
            }
            return sb.ToString().Trim();
        }

        /// <summary>A string field of a returned object, flattened and cut; "" when missing or not text.</summary>
        public static string Field(JToken item, string name, int max = 300)
        {
            var token = item?[name];
            var text = token != null && token.Type == JTokenType.String ? OneLine((string)token) : "";
            return text.Length <= max ? text : text.Substring(0, max);
        }
    }
}
