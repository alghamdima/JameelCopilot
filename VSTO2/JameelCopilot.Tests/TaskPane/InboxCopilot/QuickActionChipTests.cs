using System.Linq;
using JameelCopilot.TaskPane.InboxCopilot;
using Xunit;

namespace JameelCopilot.Tests.TaskPane.InboxCopilot
{
    /// <summary>
    /// One Inbox pane: the chips for the selected email, for the mailbox, and
    /// the reports that used to live in a second, look-alike pane.
    /// </summary>
    public class QuickActionChipTests
    {
        private static string[] Titles(int selected) => QuickActionChip.Groups(selected).Select(g => g.Title).ToArray();

        private static string[] Labels(int selected, string group) =>
            QuickActionChip.Groups(selected).Single(g => g.Title == group).Chips.Select(c => c.Label).ToArray();

        [Fact]
        public void NoSelection_OffersTheMailboxAndTheReports()
        {
            Assert.Equal(new[] { "My mailbox", "Reports" }, Titles(0));
            Assert.Equal(new[] { "Start my day", "Prioritize my inbox", "Waiting for a reply", "Tidy my inbox", "Unread", "Today", "Today's agenda", "Prepare for my next meeting" }, Labels(0, "My mailbox"));
            Assert.Equal(QuickActionChip.TidyChipId, QuickActionChip.Groups(0).Single(g => g.Title == "My mailbox").Chips[3].Id);
            Assert.Equal(QuickActionChip.PrioritizeChipId, QuickActionChip.Groups(0).Single(g => g.Title == "My mailbox").Chips[1].Id);
            Assert.Equal(QuickActionChip.WaitingChipId, QuickActionChip.Groups(0).Single(g => g.Title == "My mailbox").Chips[2].Id);
            var day = QuickActionChip.Groups(0).Single(g => g.Title == "My mailbox").Chips[0];   // the day leads when nothing is selected
            Assert.Equal("action", day.Mode);
            Assert.Equal(QuickActionChip.DayChipId, day.Id);
            Assert.True(day.Primary);
            Assert.False(QuickActionChip.Groups(1).Single(g => g.Title == "My mailbox").Chips[0].Primary);   // Draft a reply leads then
        }

        [Fact]
        public void OneEmailSelected_LeadsWithDraftAReply()
        {
            Assert.Equal(new[] { "This email", "My mailbox", "Reports" }, Titles(1));
            Assert.Equal(new[] { "Draft a reply", "Summarize", "Action items", "Translate", "Schedule a meeting" }, Labels(1, "This email"));
            var actionItems = QuickActionChip.Groups(1)[0].Chips[2];   // a card with "Add to tasks", not a text answer
            Assert.Equal("action", actionItems.Mode);
            Assert.Equal(QuickActionChip.TodosChipId, actionItems.Id);
            Assert.Contains("into Arabic", QuickActionChip.Groups(1)[0].Chips[3].Prompt);

            var reply = QuickActionChip.Groups(1)[0].Chips[0];
            Assert.Equal("action", reply.Mode);
            Assert.Equal(QuickActionChip.ReplyChipId, reply.Id);
            Assert.True(reply.Primary);
            Assert.Null(reply.Prompt);
        }

        [Fact]
        public void SeveralSelected_OffersToWorkOnThemTogether_NotAReply()
        {
            Assert.Equal(new[] { "Selected emails", "My mailbox", "Reports" }, Titles(3));
            Assert.Equal(new[] { "Summarize all", "Triage" }, Labels(3, "Selected emails"));
            Assert.DoesNotContain(QuickActionChip.Groups(3).SelectMany(g => g.Chips), c => c.Id == QuickActionChip.ReplyChipId);
        }

        // The Reports pane is gone; its six reports are a group here.
        [Fact]
        public void Reports_AreAllHere()
        {
            Assert.Equal(
                new[] { "Week digest", "My to-dos", "With a person", "Project status", "Email stats", "While I was out" },
                Labels(0, "Reports"));
        }

        // The owner once got "[name or email]" sent as it stood, and then "stopped".
        [Fact]
        public void PromptWithAPlaceholder_WaitsInTheInput_TheOthersAreSent()
        {
            foreach (var chip in QuickActionChip.Groups(1).SelectMany(g => g.Chips).Where(c => c.Mode != "action"))
            {
                Assert.False(string.IsNullOrWhiteSpace(chip.Prompt), chip.Label);
                Assert.Equal(chip.Prompt.Contains("[") ? "fill" : "send", chip.Mode);
            }
            var reports = QuickActionChip.Groups(0).Single(g => g.Title == "Reports").Chips;
            Assert.Equal(new[] { "With a person", "Project status", "While I was out" }, reports.Where(c => c.Mode == "fill").Select(c => c.Label));
        }

        // A chip is what the user reads: no tool name in it.
        [Fact]
        public void Prompts_NameNoTool()
        {
            Assert.All(QuickActionChip.Groups(1).SelectMany(g => g.Chips), c => Assert.DoesNotContain("outlook_", c.Prompt ?? ""));
        }

        [Fact]
        public void ToJson_IsWhatThePageTakes()
        {
            var json = QuickActionChip.ToJson(QuickActionChip.Groups(1));

            Assert.Equal("This email", (string)json[0]["title"]);
            Assert.Equal(
                "{\"label\":\"Draft a reply\",\"mode\":\"action\",\"id\":\"reply\",\"primary\":true}",
                json[0]["chips"][0].ToString(Newtonsoft.Json.Formatting.None));
            Assert.Equal("send", (string)json[0]["chips"][1]["mode"]);
            Assert.StartsWith("Summarize the selected message", (string)json[0]["chips"][1]["prompt"]);
            Assert.Null(json[0]["chips"][1]["id"]);
            Assert.Equal("fill", (string)json[2]["chips"][2]["mode"]);
        }
    }
}
