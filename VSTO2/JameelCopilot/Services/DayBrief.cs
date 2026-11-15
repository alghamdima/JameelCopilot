using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    /// <summary>Something an email asks of the user. The user can make it an Outlook task.</summary>
    public sealed class TodoItem
    {
        public string Text { get; set; }
        public DateTime? Due { get; set; }

        /// <summary>The email it comes from.</summary>
        public MessageSummary Message { get; set; }
    }

    /// <summary>Something worth knowing that asks nothing of the user.</summary>
    public sealed class DayTopic
    {
        public string Title { get; set; }
        public string Summary { get; set; }
        public List<MessageSummary> Messages { get; set; }
    }

    /// <summary>
    /// "Start my day": what the recent mail asks of the user, what else is
    /// worth knowing, and today's meetings. The model reads the emails Jameel Copilot
    /// gives it; the meetings come from the calendar as they are. Jameel Copilot keeps
    /// only ids it gave. "Action items" is the same to-do list for one email.
    /// </summary>
    internal static class DayBrief
    {
        public const int MaxTodos = 8;
        public const int MaxTopics = 5;

        private const string TodoRules =
            "Each to-do is one short sentence that starts with a verb and says what the user must do, for whom. " +
            "due is the date the email names for it, as yyyy-MM-dd, or \"\" when it names none: never guess a date.\n";

        public static string Instructions(string language) =>
            "Prepare the user's day from the emails under Context. Reply language: " + language + ".\n" +
            "todos: what the emails ask the user to do or answer, most urgent first, at most " + MaxTodos + ", each with the id of its email. " + TodoRules +
            "topics: at most " + MaxTopics + " things worth knowing that ask nothing of the user. Each has a short title, a summary of one or two sentences, and the ids of its emails (at most 3).\n" +
            "Leave out newsletters, advertising and automatic notifications unless one asks something of the user. An email goes under todos or under topics, not both.\n" +
            "Format: {\"todos\":[{\"id\":\"<id>\",\"task\":\"<sentence>\",\"due\":\"<yyyy-MM-dd or empty>\"}]," +
            "\"topics\":[{\"title\":\"<title>\",\"summary\":\"<sentences>\",\"ids\":[\"<id>\"]}]}";

        /// <summary>For "Action items": the to-dos of the one email under Context.</summary>
        public static string ActionItemInstructions(string language) =>
            "List what the email under Context asks of the user, who is named under \"The user\". Reply language: " + language + ".\n" +
            "Only what the user must do or answer: leave out what others are asked to do and what is already done. At most " + MaxTodos + ". " + TodoRules +
            "With nothing asked of the user, return an empty list.\n" +
            "Format: {\"todos\":[{\"task\":\"<sentence>\",\"due\":\"<yyyy-MM-dd or empty>\"}]}";

        /// <summary>One email as the prompt carries it: its header on one line, its text behind "> ".</summary>
        public static string EmailContext(MessageDetail message, string ownerName, string ownerEmail)
        {
            return "The user: " + StructuredAnswers.OneLine((ownerName + " <" + ownerEmail + ">").Trim()) + "\n" +
                   "From: " + StructuredAnswers.OneLine(message?.From) +
                   " | To: " + StructuredAnswers.OneLine(string.Join("; ", message?.To ?? new string[0])) +
                   " | Received: " + (message?.ReceivedAt ?? DateTimeOffset.MinValue).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) +
                   " | Subject: " + StructuredAnswers.OneLine(message?.Subject) + "\n" +
                   StructuredAnswers.Quote(message?.BodyPlaintext, 6000);
        }

        /// <summary>
        /// The to-dos. With <paramref name="only"/> every one belongs to that
        /// email; otherwise one whose id Jameel Copilot did not give is dropped.
        /// </summary>
        public static List<TodoItem> ParseTodos(JObject answer, IReadOnlyList<MessageSummary> candidates, MessageSummary only = null)
        {
            var todos = new List<TodoItem>();
            foreach (var entry in (answer?["todos"] as JArray ?? new JArray()).OfType<JObject>())
            {
                if (todos.Count >= MaxTodos) break;
                var text = StructuredAnswers.Field(entry, "task", 240);
                var message = only ?? Known(candidates, StructuredAnswers.Field(entry, "id", 100));
                if (text.Length == 0 || message == null) continue;
                todos.Add(new TodoItem
                {
                    Text = text,
                    Message = message,
                    Due = DateTime.TryParseExact(StructuredAnswers.Field(entry, "due", 10), "yyyy-MM-dd", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var due) ? due : (DateTime?)null,
                });
            }
            return todos;
        }

        public static List<DayTopic> ParseTopics(JObject answer, IReadOnlyList<MessageSummary> candidates)
        {
            var topics = new List<DayTopic>();
            foreach (var entry in (answer?["topics"] as JArray ?? new JArray()).OfType<JObject>())
            {
                if (topics.Count >= MaxTopics) break;
                var title = StructuredAnswers.Field(entry, "title", 80);
                var messages = (entry["ids"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                    .Select(t => Known(candidates, (string)t)).Where(m => m != null).Distinct().Take(3).ToList();
                if (title.Length == 0 || messages.Count == 0) continue;
                topics.Add(new DayTopic { Title = title, Summary = StructuredAnswers.Field(entry, "summary", 400), Messages = messages });
            }
            return topics;
        }

        private static MessageSummary Known(IReadOnlyList<MessageSummary> candidates, string id)
        {
            return (candidates ?? new MessageSummary[0]).FirstOrDefault(m => string.Equals(m.Id, id, StringComparison.Ordinal));
        }

        /// <summary>The to-do rows as the page takes them. The row's place in the list is what its button names.</summary>
        public static JArray TodosJson(IEnumerable<TodoItem> todos)
        {
            return new JArray((todos ?? new TodoItem[0]).Select(t => new JObject(
                new JProperty("text", t.Text),
                new JProperty("due", t.Due.HasValue ? t.Due.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : ""),
                new JProperty("id", t.Message.Id),
                new JProperty("from", StructuredAnswers.OneLine(t.Message.From)),
                new JProperty("subject", StructuredAnswers.OneLine(t.Message.Subject)))));
        }

        /// <summary>The whole card as the page takes it.</summary>
        public static JObject ToJson(DateTime day, IReadOnlyList<TodoItem> todos, IReadOnlyList<DayTopic> topics, IReadOnlyList<CalendarEvent> events)
        {
            return new JObject(
                new JProperty("date", day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                new JProperty("todos", TodosJson(todos)),
                new JProperty("topics", new JArray(topics.Select(t => new JObject(
                    new JProperty("title", t.Title),
                    new JProperty("summary", t.Summary),
                    new JProperty("emails", new JArray(t.Messages.Select(m => new JObject(
                        new JProperty("id", m.Id),
                        new JProperty("from", StructuredAnswers.OneLine(m.From)),
                        new JProperty("subject", StructuredAnswers.OneLine(m.Subject)),
                        new JProperty("when", m.ReceivedAt.ToString("o", CultureInfo.InvariantCulture)))))))))),
                // null (not an empty list) when the calendar could not be read: the page says so.
                new JProperty("meetings", events == null ? (JToken)JValue.CreateNull() : new JArray(events.Select(e => new JObject(
                    new JProperty("subject", StructuredAnswers.OneLine(e.Subject)),
                    new JProperty("start", CalendarArgs.Local(e.Start)),
                    new JProperty("end", CalendarArgs.Local(e.End)),
                    new JProperty("all_day", e.AllDay),
                    new JProperty("location", StructuredAnswers.OneLine(e.Location)))))));
        }

        /// <summary>What the chat keeps, so a follow-up question knows the card.</summary>
        public static string AsText(IReadOnlyList<TodoItem> todos, IReadOnlyList<DayTopic> topics, IReadOnlyList<CalendarEvent> events)
        {
            var sb = new StringBuilder();
            if (todos != null && todos.Count > 0)
            {
                sb.Append("To do:\n");
                foreach (var t in todos)
                {
                    sb.Append("- ").Append(t.Text)
                      .Append(t.Due.HasValue ? " (due " + t.Due.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ")" : "")
                      .Append(" [").Append(StructuredAnswers.OneLine(t.Message.From)).Append(" - ").Append(StructuredAnswers.OneLine(t.Message.Subject)).Append("]\n");
                }
            }
            if (topics != null && topics.Count > 0)
            {
                sb.Append("Catch up:\n");
                foreach (var t in topics) sb.Append("- ").Append(t.Title).Append(t.Summary.Length > 0 ? ": " + t.Summary : "").Append('\n');
            }
            if (events != null)
            {
                sb.Append("Today's meetings:").Append(events.Count == 0 ? " none\n" : "\n");
                foreach (var e in events)
                {
                    sb.Append("- ").Append(e.AllDay ? "all day" : e.Start.ToString("HH:mm", CultureInfo.InvariantCulture) + "-" + e.End.ToString("HH:mm", CultureInfo.InvariantCulture))
                      .Append(' ').Append(StructuredAnswers.OneLine(e.Subject)).Append('\n');
                }
            }
            return sb.Length == 0 ? "Nothing is asked of you." : sb.ToString().TrimEnd('\n');
        }

        /// <summary>The note of the Outlook task made from a to-do: where it came from.</summary>
        public static string TaskNote(TodoItem todo)
        {
            var m = todo?.Message;
            return m == null ? "" : "From an email: " + StructuredAnswers.OneLine(m.From) + " - " + StructuredAnswers.OneLine(m.Subject) +
                   " (" + m.ReceivedAt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ")\r\nAdded by Jameel Copilot.";
        }
    }
}
