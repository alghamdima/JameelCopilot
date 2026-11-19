using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    public sealed class PriorityItem
    {
        public MessageSummary Message { get; set; }
        /// <summary>"high", "medium" or "low".</summary>
        public string Level { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// "Prioritize my inbox": which emails need the user, which can wait,
    /// which are for information. The model ranks the emails Jameel Copilot gives it; Jameel Copilot
    /// keeps only ids it gave and puts every email it gave somewhere.
    /// </summary>
    internal static class InboxPriorities
    {
        public static readonly string[] Levels = { "high", "medium", "low" };

        /// <summary>The emails to rank: unread mail of the last 7 days, topped up with all mail of the last 2 days when that is little.</summary>
        public static SearchMessagesArgs UnreadArgs(DateTimeOffset now) => new SearchMessagesArgs
        {
            FolderId = "inbox", Scope = "current_folder", ReadStatus = "unread", DateFrom = now.AddDays(-7), MaxResults = 40, SortOrder = "newest",
        };

        public static SearchMessagesArgs RecentArgs(DateTimeOffset now) => new SearchMessagesArgs
        {
            FolderId = "inbox", Scope = "current_folder", DateFrom = now.AddDays(-2), MaxResults = 25, SortOrder = "newest",
        };

        public static List<MessageSummary> Candidates(IEnumerable<MessageSummary> unread, IEnumerable<MessageSummary> recent)
        {
            var list = new List<MessageSummary>();
            foreach (var m in (unread ?? new MessageSummary[0]).Concat(recent ?? new MessageSummary[0]))
            {
                if (m == null || string.IsNullOrEmpty(m.Id) || list.Any(x => x.Id == m.Id)) continue;
                if (list.Count >= 40) break;
                list.Add(m);
            }
            return list;
        }

        public static string Instructions(string language) =>
            "Sort the emails under Context by how much they need the user. Reply language: " + language + ".\n" +
            "high: needs the user's action or reply soon - a person asks them something, a deadline, an approval, a problem.\n" +
            "medium: relevant to the user's work, can wait.\n" +
            "low: newsletters, automatic notifications, announcements, mail for information only.\n" +
            "Put every email in once. The reason is one short sentence: why, and what the user should do.\n" +
            "Format: {\"items\":[{\"id\":\"<id>\",\"priority\":\"high|medium|low\",\"reason\":\"<sentence>\"}]}";

        /// <summary>The ranked emails; an email the answer left out is "low" with no reason; an unknown id is dropped.</summary>
        public static List<PriorityItem> Parse(JObject answer, IReadOnlyList<MessageSummary> candidates)
        {
            var byId = candidates.ToDictionary(m => m.Id, StringComparer.Ordinal);
            var items = new List<PriorityItem>();
            foreach (var entry in (answer?["items"] as JArray ?? new JArray()).OfType<JObject>())
            {
                var id = StructuredAnswers.Field(entry, "id", 100);
                if (!byId.TryGetValue(id, out var message) || items.Any(i => i.Message.Id == id)) continue;
                var level = StructuredAnswers.Field(entry, "priority", 10).ToLowerInvariant();
                items.Add(new PriorityItem
                {
                    Message = message,
                    Level = Levels.Contains(level) ? level : "low",
                    Reason = StructuredAnswers.Field(entry, "reason", 240),
                });
            }
            foreach (var missing in candidates.Where(m => items.All(i => i.Message.Id != m.Id)))
            {
                items.Add(new PriorityItem { Message = missing, Level = "low", Reason = "" });
            }
            return items;
        }

        /// <summary>The groups as the page takes them.</summary>
        public static JArray ToJson(IReadOnlyList<PriorityItem> items)
        {
            var labels = new Dictionary<string, string> { ["high"] = "Needs you", ["medium"] = "Can wait", ["low"] = "For information" };
            return new JArray(Levels.Select(level => new JObject(
                new JProperty("level", level),
                new JProperty("label", labels[level]),
                new JProperty("items", new JArray(items.Where(i => i.Level == level).Select(i => new JObject(
                    new JProperty("id", i.Message.Id),
                    new JProperty("from", StructuredAnswers.OneLine(i.Message.From)),
                    new JProperty("subject", StructuredAnswers.OneLine(i.Message.Subject)),
                    new JProperty("when", i.Message.ReceivedAt.ToString("o", CultureInfo.InvariantCulture)),
                    new JProperty("reason", i.Reason))))))));
        }

        /// <summary>What the chat keeps, so a follow-up question knows the ranking.</summary>
        public static string AsText(IReadOnlyList<PriorityItem> items)
        {
            var sb = new StringBuilder();
            foreach (var level in Levels)
            {
                var group = items.Where(i => i.Level == level).ToList();
                if (group.Count == 0) continue;
                sb.Append(level == "high" ? "Needs you" : level == "medium" ? "Can wait" : "For information").Append(":\n");
                foreach (var i in group)
                {
                    sb.Append("- ").Append(StructuredAnswers.OneLine(i.Message.From)).Append(" - ")
                      .Append(StructuredAnswers.OneLine(i.Message.Subject))
                      .Append(i.Reason.Length > 0 ? ": " + i.Reason : "").Append('\n');
                }
            }
            return sb.ToString().TrimEnd('\n');
        }
    }
}
