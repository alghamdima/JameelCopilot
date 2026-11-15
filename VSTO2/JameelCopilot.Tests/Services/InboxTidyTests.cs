using System;
using System.IO;
using System.Linq;
using JameelCopilot.Services;
using JameelCopilot.Services.Tools;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>
    /// "Tidy my inbox" (Copilot's triage), kept to two reversible changes the
    /// user applies: mark as read, and flag.
    /// </summary>
    public class InboxTidyTests
    {
        private static MessageSummary Mail(string id) => new MessageSummary
        {
            Id = id, From = "Sender " + id, Subject = "Subject " + id, ReceivedAt = new DateTimeOffset(2026, 11, 13, 9, 0, 0, TimeSpan.Zero),
        };

        private static readonly MessageSummary[] Mails = { Mail("a"), Mail("b"), Mail("c") };

        [Fact]
        public void Parse_KeepsKnownIdsAndActions_OnceEach()
        {
            var answer = StructuredAnswers.ParseObject("{\"items\":[" +
                "{\"id\":\"a\",\"action\":\"MARK_READ\",\"reason\":\"Newsletter.\"}," +
                "{\"id\":\"b\",\"action\":\"flag\",\"reason\":\"Asks for approval.\"}," +
                "{\"id\":\"c\",\"action\":\"delete\",\"reason\":\"no\"}," +
                "{\"id\":\"zz\",\"action\":\"flag\"}," +
                "{\"id\":\"a\",\"action\":\"flag\"}]}");

            var items = InboxTidy.Parse(answer, Mails);

            Assert.Equal(new[] { "a:mark_read", "b:flag" }, items.Select(i => i.Message.Id + ":" + i.Action));
            Assert.Empty(InboxTidy.Parse(null, Mails));
        }

        // Only unread, unflagged mail is proposed, so Undo (unread, no flag) puts it back exactly.
        [Fact]
        public void Candidates_AreUnreadAndUnflagged_FromTheLastTwoWeeks()
        {
            var now = new DateTimeOffset(2026, 11, 14, 10, 0, 0, TimeSpan.FromHours(3));
            var args = InboxTidy.Args(now);

            Assert.Equal("unread", args.ReadStatus);
            Assert.Equal("unflagged", args.FlagStatus);
            Assert.Equal(now.AddDays(-14), args.DateFrom);
        }

        [Fact]
        public void ToJson_GroupsByAction_AndInstructions_OfferOnlyTheTwoChanges()
        {
            var items = InboxTidy.Parse(StructuredAnswers.ParseObject("{\"items\":[{\"id\":\"b\",\"action\":\"flag\",\"reason\":\"Asks.\"}]}"), Mails);
            var groups = InboxTidy.ToJson(items);

            Assert.Equal(new[] { "flag", "mark_read" }, groups.Select(g => (string)g["action"]));
            Assert.Equal("b", (string)groups[0]["items"][0]["id"]);
            Assert.Empty(groups[1]["items"]);

            var text = InboxTidy.Instructions("English");
            Assert.Contains("{\"items\":[{\"id\":\"<id>\",\"action\":\"mark_read|flag\",\"reason\":\"<sentence>\"}]}", text);
            Assert.Contains("Leave out any email you are not sure about.", text);
            Assert.StartsWith("Proposed tidy-up (nothing changed until the user applies it):", InboxTidy.AsText(items));
        }

        // Apply and Undo act on the host's own list, never on ids from the page; nothing is moved or deleted.
        [Fact]
        public void ApplyAndUndo_UseTheHostsList_AndOnlyReadAndFlag()
        {
            var controller = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "TaskPane", "InboxCopilot", "InboxCopilotController.cs"));
            var apply = controller.Substring(controller.IndexOf("private async Task ApplyTidyAsync(", StringComparison.Ordinal));
            apply = apply.Substring(0, apply.IndexOf("// \"Reply\" on a ranked row", StringComparison.Ordinal));

            Assert.Contains("_tidy.TryGetValue(cardId, out var items)", apply);
            Assert.Contains("items.Where(i => wanted.Contains(i.Message.Id))", apply);
            Assert.Contains("_surface.FlagMessage(item.Message.Id, apply ? \"todo\" : \"none\")", apply);
            Assert.Contains("_surface.MarkAsRead(item.Message.Id, apply)", apply);
            foreach (var verb in new[] { "Move", "Delete", "SetCategory", ".Send(" }) Assert.DoesNotContain(verb, apply);
        }
    }
}
