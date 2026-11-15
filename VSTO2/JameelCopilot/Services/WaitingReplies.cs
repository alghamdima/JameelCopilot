using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    public sealed class WaitingItem
    {
        public MessageSummary Message { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// "Waiting for a reply": emails the user sent that nobody answered. Jameel Copilot
    /// finds the unanswered ones (no reply in the conversation); the model
    /// keeps those that asked for something and says what the user waits for.
    /// </summary>
    internal static class WaitingReplies
    {
        /// <summary>How far back sent mail is looked at, and how long a recipient is given before it counts.</summary>
        public const int Days = 14;
        public const int GraceDays = 1;
        public const int MaxCandidates = 30;

        public static string Instructions(string language) =>
            "The emails under Context were sent by the user and nobody has answered them. Reply language: " + language + ".\n" +
            "Keep the ones where the user is waiting for an answer: they asked a question, requested something, or need a decision or a confirmation.\n" +
            "Leave out thanks, information that needs no answer, replies that close a thread, and automatic messages.\n" +
            "The reason is one short sentence: what the user is waiting for, and from whom.\n" +
            "Format: {\"items\":[{\"id\":\"<id>\",\"reason\":\"<sentence>\"}]}";

        /// <summary>The sent emails as the prompt carries them: recipient-written text there is none, but a subject may be a sender's.</summary>
        public static string Context(IEnumerable<MessageSummary> sent)
        {
            var sb = new StringBuilder();
            foreach (var m in sent ?? new MessageSummary[0])
            {
                sb.Append("[id ").Append(StructuredAnswers.OneLine(m.Id)).Append("] To: ").Append(Recipients(m))
                  .Append(" | Sent: ").Append(m.ReceivedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture))
                  .Append(" | Subject: ").Append(StructuredAnswers.OneLine(m.Subject)).Append('\n');
                var snippet = StructuredAnswers.OneLine(m.Snippet);
                if (snippet.Length > 0) sb.Append("> ").Append(snippet.Length > 400 ? snippet.Substring(0, 400) : snippet).Append('\n');
            }
            return sb.ToString().TrimEnd('\n');
        }

        /// <summary>The emails the model kept, oldest first (the longest wait leads); an unknown id is dropped.</summary>
        public static List<WaitingItem> Parse(JObject answer, IReadOnlyList<MessageSummary> candidates)
        {
            var items = new List<WaitingItem>();
            foreach (var entry in (answer?["items"] as JArray ?? new JArray()).OfType<JObject>())
            {
                var id = StructuredAnswers.Field(entry, "id", 100);
                var message = (candidates ?? new MessageSummary[0]).FirstOrDefault(m => string.Equals(m.Id, id, StringComparison.Ordinal));
                if (message == null || items.Any(i => i.Message.Id == id)) continue;
                items.Add(new WaitingItem { Message = message, Reason = StructuredAnswers.Field(entry, "reason", 240) });
            }
            return items.OrderBy(i => i.Message.ReceivedAt).ToList();
        }

        /// <summary>One group of the ranked-list card, whose row button asks for a follow-up.</summary>
        public static JArray ToJson(IReadOnlyList<WaitingItem> items, DateTimeOffset now)
        {
            return new JArray(new JObject(
                new JProperty("level", "medium"),
                new JProperty("label", "Waiting for a reply"),
                new JProperty("button", new JObject(new JProperty("label", "Follow up"), new JProperty("type", "followup"))),
                new JProperty("items", new JArray(items.Select(i => new JObject(
                    new JProperty("id", i.Message.Id),
                    new JProperty("from", "To: " + Recipients(i.Message)),
                    new JProperty("subject", StructuredAnswers.OneLine(i.Message.Subject)),
                    new JProperty("when", i.Message.ReceivedAt.ToString("o", CultureInfo.InvariantCulture)),
                    new JProperty("reason", (i.Reason.Length > 0 ? i.Reason + " " : "") + "(" + Waited(i.Message.ReceivedAt, now) + ")")))))));
        }

        public static string AsText(IReadOnlyList<WaitingItem> items)
        {
            var sb = new StringBuilder("Waiting for a reply:\n");
            foreach (var i in items)
            {
                sb.Append("- To ").Append(Recipients(i.Message)).Append(" - ").Append(StructuredAnswers.OneLine(i.Message.Subject))
                  .Append(", sent ").Append(i.Message.ReceivedAt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                  .Append(i.Reason.Length > 0 ? ": " + i.Reason : "").Append('\n');
            }
            return sb.ToString().TrimEnd('\n');
        }

        internal static string Waited(DateTimeOffset sent, DateTimeOffset now)
        {
            var days = Math.Max(0, (now.Date - sent.ToOffset(now.Offset).Date).Days);
            return days == 0 ? "sent today" : days == 1 ? "sent yesterday" : "sent " + days.ToString(CultureInfo.InvariantCulture) + " days ago";
        }

        private static string Recipients(MessageSummary m)
        {
            var to = (m?.To ?? new string[0]).Select(StructuredAnswers.OneLine).Where(s => s.Length > 0).ToList();
            return to.Count == 0 ? "(nobody)" : to.Count <= 2 ? string.Join("; ", to) : to[0] + "; " + to[1] + " +" + (to.Count - 2).ToString(CultureInfo.InvariantCulture);
        }
    }
}
