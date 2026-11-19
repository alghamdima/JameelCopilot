using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.TaskPane.InboxCopilot
{
    /// <summary>
    /// A chip above the Inbox pane's input: a ready-made request.
    /// </summary>
    public sealed class QuickActionChip
    {
        /// <summary>The chip that writes a reply to the selected email as a draft card.</summary>
        public const string ReplyChipId = "reply";

        /// <summary>The chip that ranks the recent unread mail.</summary>
        public const string PrioritizeChipId = "prioritize";

        /// <summary>The chip that proposes marking read and flagging, for the user to apply.</summary>
        public const string TidyChipId = "tidy";

        /// <summary>The chip that shows the day: to-dos, topics and today's meetings.</summary>
        public const string DayChipId = "day";

        /// <summary>The chip that lists what the selected email asks of the user, each with "Add to tasks".</summary>
        public const string TodosChipId = "todos";

        /// <summary>The chip that lists sent emails nobody answered.</summary>
        public const string WaitingChipId = "waiting";

        public string Label { get; set; }

        /// <summary>What is sent as the user's message, or put in the input.</summary>
        public string Prompt { get; set; }

        /// <summary>
        /// "send" sends the prompt; "fill" puts it in the input because it has
        /// a [placeholder] only the user can complete; "action" is run by the
        /// pane itself (<see cref="Id"/>).
        /// </summary>
        public string Mode { get; set; } = "send";

        public string Id { get; set; }
        public bool Primary { get; set; }

        /// <summary>
        /// The chips for what is selected, in groups: the selected email (or
        /// emails), the mailbox, and the reports that used to have a pane of
        /// their own.
        /// </summary>
        public static IReadOnlyList<QuickActionGroup> Groups(int selectionCount)
        {
            var groups = new List<QuickActionGroup>();
            if (selectionCount == 1)
            {
                groups.Add(new QuickActionGroup("This email",
                    new QuickActionChip { Label = "Draft a reply", Mode = "action", Id = ReplyChipId, Primary = true },
                    Send("Summarize", "Summarize the selected message and the rest of its conversation thread."),
                    new QuickActionChip { Label = "Action items", Mode = "action", Id = TodosChipId },
                    Send("Translate", "Translate the main message body of the selected email (ignore classification footers and signatures): into English if the main body is written in Arabic, into Arabic if the main body is written in any other language. Give the translation only, and keep names, numbers and dates as they are."),
                    Send("Schedule a meeting", "Help me schedule a meeting about the selected email (default 1 hour, excluding 12:30-13:30 break; To/sender as Required and any CC as Optional). Before checking calendars, briefly confirm with me: 1) duration (1 hour default) and room vs Microsoft Teams, and 2) if there are CC recipients on the email, whether to check their calendar availability too or keep them Optional without checking their calendars.")));
            }
            else if (selectionCount >= 2)
            {
                groups.Add(new QuickActionGroup("Selected emails",
                    Send("Summarize all", "Summarize all the selected messages."),
                    Send("Triage", "Triage the selected messages: which need action, which can be archived, which can be marked read?")));
            }

            groups.Add(new QuickActionGroup("My mailbox",
                new QuickActionChip { Label = "Start my day", Mode = "action", Id = DayChipId, Primary = selectionCount == 0 },
                new QuickActionChip { Label = "Prioritize my inbox", Mode = "action", Id = PrioritizeChipId },
                new QuickActionChip { Label = "Waiting for a reply", Mode = "action", Id = WaitingChipId },
                new QuickActionChip { Label = "Tidy my inbox", Mode = "action", Id = TidyChipId },
                Send("Unread", "Summarize all my unread messages. Group by sender or topic. Be concise."),
                Send("Today", "Show me everything I received today, grouped by sender. Highlight anything that looks urgent."),
                Send("Today's agenda", "List my meetings today in order from my calendar: time, subject, who, where. Point out clashes and gaps longer than an hour."),
                Send("Prepare for my next meeting", "Find my next upcoming meeting in my calendar today or tomorrow. Run at most 2 targeted searches in my recent mail of the last 30 days using distinctive subject words or the organizer's name, read the top 3 to 5 most relevant messages, and immediately brief me: what it is about, what was agreed, what is still open, and what I should prepare.")));

            groups.Add(new QuickActionGroup("Reports",
                Report("Week digest", "Summarize what came into my Inbox over the past 7 days. Group by sender or topic. Highlight urgent items and emails I'm directly addressed in."),
                Report("My to-dos", "Find action items I need to do based on emails from the past 7 days. Read the relevant messages, extract TODOs/deadlines/asks. Group by who's waiting on what."),
                Report("With a person", "Summarize my recent email conversations with [name or email]. Show the chronological flow and key decisions/topics."),
                Report("Project status", "Summarize the status of [topic/project name]. Find relevant emails, read the most recent ones, and give me: latest update, open questions, action items, key participants."),
                Report("Email stats", "Give me email statistics for the past 30 days: top 10 senders, busiest days, breakdown by folder."),
                Report("While I was out", "I was out from [start date] to [end date]. Show me what's important from that timeframe: urgent items, direct asks, replies needed. De-prioritize newsletters and automated mail.")));
            return groups;
        }

        /// <summary>The groups as the page takes them.</summary>
        public static JArray ToJson(IEnumerable<QuickActionGroup> groups)
        {
            return new JArray(groups.Select(g => new JObject(
                new JProperty("title", g.Title),
                new JProperty("chips", new JArray(g.Chips.Select(c =>
                {
                    var chip = new JObject(new JProperty("label", c.Label), new JProperty("mode", c.Mode));
                    if (c.Prompt != null) chip.Add("prompt", c.Prompt);
                    if (c.Id != null) chip.Add("id", c.Id);
                    if (c.Primary) chip.Add("primary", true);
                    return chip;
                }))))));
        }

        private static QuickActionChip Send(string label, string prompt)
        {
            return new QuickActionChip { Label = label, Prompt = prompt };
        }

        // A report whose text still has a [placeholder] waits in the input for the user.
        private static QuickActionChip Report(string label, string prompt)
        {
            return new QuickActionChip { Label = label, Prompt = prompt, Mode = prompt.Contains("[") ? "fill" : "send" };
        }
    }

    public sealed class QuickActionGroup
    {
        public QuickActionGroup(string title, params QuickActionChip[] chips)
        {
            Title = title;
            Chips = chips;
        }

        public string Title { get; }
        public IReadOnlyList<QuickActionChip> Chips { get; }
    }
}
