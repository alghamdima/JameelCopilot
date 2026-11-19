using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    public sealed class TidyItem
    {
        public MessageSummary Message { get; set; }
        /// <summary>"mark_read" or "flag".</summary>
        public string Action { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// "Tidy my inbox" (Copilot's triage, kept to safe, reversible changes):
    /// for unread, unflagged mail, mark as read what needs nothing from the
    /// user and flag what asks something of them. Jameel Copilot proposes; the user ticks
    /// and applies; Undo puts each email back.
    /// </summary>
    internal static class InboxTidy
    {
        public static readonly string[] Actions = { "flag", "mark_read" };

        // Unflagged only, so Undo (no flag) is exactly how it was.
        public static SearchMessagesArgs Args(DateTimeOffset now) => new SearchMessagesArgs
        {
            FolderId = "inbox", Scope = "current_folder", ReadStatus = "unread", FlagStatus = "unflagged", DateFrom = now.AddDays(-14), MaxResults = 50, SortOrder = "newest",
        };

        public static string Instructions(string language) =>
            "Help the user tidy the unread emails under Context. Reply language: " + language + ".\n" +
            "mark_read: the email needs nothing from the user - newsletters, automatic notifications, announcements, mail sent for information.\n" +
            "flag: the email asks the user to do or answer something.\n" +
            "Leave out any email you are not sure about. The reason is one short sentence.\n" +
            "Format: {\"items\":[{\"id\":\"<id>\",\"action\":\"mark_read|flag\",\"reason\":\"<sentence>\"}]}";

        /// <summary>The proposals; unknown ids and actions are dropped, each email once.</summary>
        public static List<TidyItem> Parse(JObject answer, IReadOnlyList<MessageSummary> candidates)
        {
            var byId = candidates.ToDictionary(m => m.Id, StringComparer.Ordinal);
            var items = new List<TidyItem>();
            foreach (var entry in (answer?["items"] as JArray ?? new JArray()).OfType<JObject>())
            {
                var id = StructuredAnswers.Field(entry, "id", 100);
                var action = StructuredAnswers.Field(entry, "action", 20).ToLowerInvariant();
                if (!byId.TryGetValue(id, out var message) || !Actions.Contains(action) || items.Any(i => i.Message.Id == id)) continue;
                items.Add(new TidyItem { Message = message, Action = action, Reason = StructuredAnswers.Field(entry, "reason", 240) });
            }
            return items;
        }

        public static JArray ToJson(IReadOnlyList<TidyItem> items)
        {
            var labels = new Dictionary<string, string> { ["flag"] = "Flag for follow-up", ["mark_read"] = "Mark as read" };
            return new JArray(Actions.Select(action => new JObject(
                new JProperty("action", action),
                new JProperty("label", labels[action]),
                new JProperty("items", new JArray(items.Where(i => i.Action == action).Select(i => new JObject(
                    new JProperty("id", i.Message.Id),
                    new JProperty("from", StructuredAnswers.OneLine(i.Message.From)),
                    new JProperty("subject", StructuredAnswers.OneLine(i.Message.Subject)),
                    new JProperty("when", i.Message.ReceivedAt.ToString("o", CultureInfo.InvariantCulture)),
                    new JProperty("reason", i.Reason))))))));
        }

        public static string AsText(IReadOnlyList<TidyItem> items)
        {
            var sb = new StringBuilder("Proposed tidy-up (nothing changed until the user applies it):");
            foreach (var i in items)
            {
                sb.Append("\n- ").Append(i.Action == "flag" ? "flag: " : "mark read: ")
                  .Append(StructuredAnswers.OneLine(i.Message.From)).Append(" - ").Append(StructuredAnswers.OneLine(i.Message.Subject));
            }
            return sb.ToString();
        }
    }
}
