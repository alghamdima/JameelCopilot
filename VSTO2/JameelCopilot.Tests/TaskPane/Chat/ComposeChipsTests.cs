using System.Linq;
using Newtonsoft.Json.Linq;
using JameelCopilot.TaskPane.Chat;
using Xunit;
using ActionType = JameelCopilot.Services.CustomApiChatService.ActionType;

namespace JameelCopilot.Tests.TaskPane.Chat
{
    /// <summary>
    /// The chips of a compose window and the buttons of a draft card. They
    /// replace the old Actions tab: every result is a card the user decides on.
    /// </summary>
    public class ComposeChipsTests
    {
        private static string[] Labels(JToken group) => group["chips"].Select(c => (string)c["label"]).ToArray();

        [Fact]
        public void Reply_OffersToWriteTheReply_ThenToEditWhatWasWritten()
        {
            var groups = ComposeChips.Groups(isReply: true);

            Assert.Equal(new[] { "Write my reply", "Edit what I wrote" }, groups.Select(g => (string)g["title"]));
            Assert.Equal(new[] { "Formal", "Friendly", "Short", "Detailed", "In Arabic", "In English" }, Labels(groups[0]));
            Assert.Equal(new[] { "Review", "Proofread", "Improve", "Shorten", "Lengthen", "Formal", "Friendly", "To Arabic", "To English" }, Labels(groups[1]));
            Assert.Equal(ComposeChips.ReviewChipId, (string)groups[1]["chips"][0]["id"]);
            Assert.Null(groups[1]["chips"][0]["primary"]);   // in a reply, writing the reply leads
            Assert.All(groups.SelectMany(g => g["chips"]), chip => Assert.Equal("action", (string)chip["mode"]));
            Assert.True((bool)groups[0]["chips"][0]["primary"]);
        }

        [Fact]
        public void NewEmail_HasNoReplyToWrite()
        {
            var groups = ComposeChips.Groups(isReply: false);

            Assert.Equal("Edit what I wrote", (string)Assert.Single(groups)["title"]);
            Assert.True((bool)groups[0]["chips"][0]["primary"]);   // Review leads in a new email
        }

        [Fact]
        public void EveryChip_StandsForItsAction()
        {
            var chips = ComposeChips.Groups(isReply: true).SelectMany(g => g["chips"]).Where(c => (string)c["id"] != ComposeChips.ReviewChipId).ToList();
            Assert.False(ComposeChips.TryParse(ComposeChips.ReviewChipId, out _));   // Review is not a rewrite

            foreach (var chip in chips)
            {
                Assert.True(ComposeChips.TryParse((string)chip["id"], out _), (string)chip["id"]);
            }
            Assert.True(ComposeChips.TryParse("reply:Shorten", out var shorten));
            Assert.Equal(ActionType.Shorten, shorten);
            Assert.True(ComposeChips.TryParse("edit:Revise", out var improve));
            Assert.Equal(ActionType.Revise, improve);
        }

        // The id comes from the page: anything else runs nothing.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Formal")]
        [InlineData("reply:")]
        [InlineData("reply:Draft")]
        [InlineData("edit:Custom")]
        [InlineData("edit:formal")]
        [InlineData("edit:3")]
        [InlineData("edit:Nonsense")]
        [InlineData("other:Formal")]
        public void UnknownChipId_IsRefused(string id)
        {
            Assert.False(ComposeChips.TryParse(id, out _));
        }

        [Theory]
        [InlineData(ActionType.Formal, true, "Reply draft - Formal")]
        [InlineData(ActionType.Shorten, true, "Reply draft - Short")]
        [InlineData(ActionType.Lengthen, true, "Reply draft - Detailed")]
        [InlineData(ActionType.Proofread, true, "Reply draft")]
        [InlineData(ActionType.Proofread, false, "Proofread")]
        [InlineData(ActionType.Revise, false, "Improve")]
        [InlineData(ActionType.Shorten, false, "Shorten")]
        [InlineData(ActionType.Draft, false, "Draft")]
        [InlineData(ActionType.Draft, true, "Draft")]
        [InlineData(ActionType.Custom, false, "Draft with the review's changes")]
        [InlineData(ActionType.Arabic, true, "Reply draft - In Arabic")]
        [InlineData(ActionType.English, false, "Translate to English")]
        public void DraftTitle_SaysWhatTheCardHolds(ActionType action, bool writesReply, string expected)
        {
            Assert.Equal(expected, ComposeChips.DraftTitle(action, writesReply));
        }

        // With text of the user's in the email the draft stands in for it; with none it goes in.
        [Fact]
        public void DraftButtons_LeadWithWhatTheUserMostLikelyWants()
        {
            var forEmpty = ComposeChips.DraftButtons(hasOwnText: false);
            var forText = ComposeChips.DraftButtons(hasOwnText: true);

            Assert.Equal(new[] { "insert", "regenerate", "discard" }, forEmpty.Select(b => (string)b["action"]));
            Assert.Equal(new[] { "replace", "insert", "regenerate", "discard" }, forText.Select(b => (string)b["action"]));
            Assert.Equal("Replace my text", (string)forText[0]["label"]);
            Assert.True((bool)forEmpty[0]["primary"] && (bool)forText[0]["primary"]);
            Assert.Single(forText.Where(b => b["primary"] != null));
            Assert.Equal("discard", (string)Assert.Single(ComposeChips.DiscardOnly())["action"]);
            Assert.Equal(new[] { "replace_selection", "regenerate", "discard" }, ComposeChips.SelectionButtons().Select(b => (string)b["action"]));
            Assert.Equal("Replace selection", (string)ComposeChips.SelectionButtons()[0]["label"]);
        }

        [Theory]
        [InlineData(ActionType.Formal, true, "", "Write my reply (formal).")]
        [InlineData(ActionType.Proofread, true, "", "Write my reply (standard).")]
        [InlineData(ActionType.Proofread, false, "", "Proofread what I wrote.")]
        [InlineData(ActionType.Revise, false, "", "Improve what I wrote.")]
        [InlineData(ActionType.Draft, false, " say yes ", "Write the email: say yes")]
        [InlineData(ActionType.Custom, false, "", "Apply the review's changes.")]
        [InlineData(ActionType.English, true, "", "Write my reply (in english).")]
        [InlineData(ActionType.Arabic, false, "", "Translate what I wrote to Arabic.")]
        public void Asked_IsTheUsersSideOfTheDraftInTheChat(ActionType action, bool writesReply, string instructions, string expected)
        {
            Assert.Equal(expected, ComposeChips.Asked(action, writesReply, instructions));
        }
        // A follow-up goes to the people the user's own email went to: reply all on the sent email.
        [Fact]
        public void FollowUpButtons_OpenTheFollowUp_ToEveryoneTheEmailWentTo()
        {
            var buttons = ComposeChips.FollowUpButtons();

            Assert.Equal(new[] { "reply_all", "regenerate", "discard" }, buttons.Select(b => (string)b["action"]));
            Assert.Equal("Open follow-up", (string)buttons[0]["label"]);
            Assert.True((bool)buttons[0]["primary"]);
        }
    }
}
