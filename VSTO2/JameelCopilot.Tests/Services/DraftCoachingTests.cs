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
    /// "Review" (Copilot's coaching) and working on a selection only.
    /// </summary>
    public class DraftCoachingTests
    {
        [Fact]
        public void Parse_KeepsSixPoints_WithOurOwnIds_OneLineEach()
        {
            var json = "{\"summary\":\"Clear,\\nbut one question is not answered.\",\"points\":[" +
                       string.Join(",", Enumerable.Range(1, 8).Select(i => "{\"id\":\"x" + i + "\",\"issue\":\"issue " + i + "\",\"suggestion\":\"fix " + i + "\"}")) +
                       ",{\"id\":\"9\",\"issue\":\"no fix\"}]}";
            var answer = StructuredAnswers.ParseObject(json);

            var points = DraftCoaching.Parse(answer);

            Assert.Equal(DraftCoaching.MaxPoints, points.Count);
            Assert.Equal(new[] { "1", "2", "3", "4", "5", "6" }, points.Select(p => p.Id));
            Assert.Equal("fix 1", points[0].Suggestion);
            Assert.Equal("Clear, but one question is not answered.", DraftCoaching.Summary(answer));
            Assert.Empty(DraftCoaching.Parse(null));
        }

        [Fact]
        public void ApplyInstructions_AreTheChosenChangesAndNothingElse()
        {
            var points = DraftCoaching.Parse(StructuredAnswers.ParseObject(
                "{\"points\":[{\"issue\":\"a\",\"suggestion\":\"Answer the question about the date.\"},{\"issue\":\"b\",\"suggestion\":\"Mention the attachment.\"}]}"));

            var text = DraftCoaching.ApplyInstructions(points.Take(1));

            Assert.Equal("Make these changes to the text and nothing else. Keep everything else as it is, in the same language:\n- Answer the question about the date.", text);
            Assert.Equal("Fine.\n- a -> Answer the question about the date.\n- b -> Mention the attachment.", DraftCoaching.AsText("Fine.", points));
        }

        [Fact]
        public void Instructions_CheckTheAnswerToneClarityAndAttachments()
        {
            var text = DraftCoaching.Instructions("Arabic");

            Assert.Contains("Reply language: Arabic.", text);
            Assert.Contains("does it answer everything the message being replied to asks", text);
            Assert.Contains("does it mention an attachment that is not attached", text);
            Assert.Contains("No points when it is ready to send.", text);
            Assert.Contains("{\"summary\":\"<one sentence on the draft as a whole>\",\"points\":[", text);
        }

        // A review can only see a missing attachment if the context lists them.
        [Fact]
        public void Context_ListsTheWindowsAttachments()
        {
            var state = new ComposeStateResult { SenderName = "Omar", Subject = "RE: Budget", BodyPlaintext = "See attached." };

            Assert.DoesNotContain("Attachments:", ComposePrompt.Context(state, includeDraft: true));
            state.Attachments = new AttachmentSummary[0];
            Assert.Contains("\nAttachments: (none)\n", ComposePrompt.Context(state, includeDraft: true));
            state.Attachments = new[] { new AttachmentSummary { Filename = "budget.xlsx" }, new AttachmentSummary { Filename = "notes\n.pdf" } };
            Assert.Contains("\nAttachments: budget.xlsx, notes .pdf\n", ComposePrompt.Context(state, includeDraft: true));
        }

        [Fact]
        public void WithOwnText_IsTheSameWindowWithTheSelectionAsTheText()
        {
            var state = new ComposeStateResult
            {
                Subject = "RE: Budget", SenderName = "Omar", BodyPlaintext = "Line one. Line two.", QuotedPlaintext = "From: Sara",
                ToRecipients = new[] { "Sara" },
            };

            var only = ComposePrompt.WithOwnText(state, "Line two.");

            Assert.Equal("Line two.", only.BodyPlaintext);
            Assert.Equal("From: Sara", only.QuotedPlaintext);
            Assert.Equal(new[] { "Sara" }, only.ToRecipients);
            Assert.Equal("Line one. Line two.", state.BodyPlaintext);   // the window itself is not changed
        }

        // Replace selection puts the text back only where it was and only if it is still the same.
        [Fact]
        public void ReplaceSelection_ChecksTheTextIsUnchanged_AndStaysInTheUsersOwnPart()
        {
            var editor = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "Services", "Tools", "ComposeDocument.cs"));
            var read = editor.Substring(editor.IndexOf("public static ComposeSelection ReadSelection(", StringComparison.Ordinal));
            read = read.Substring(0, read.IndexOf("public static bool ReplaceSelection(", StringComparison.Ordinal));
            var replace = editor.Substring(editor.IndexOf("public static bool ReplaceSelection(", StringComparison.Ordinal));
            replace = replace.Substring(0, replace.IndexOf("private static Word.Document Editor(", StringComparison.Ordinal));

            Assert.Contains("if (selection.Start < 0 || selection.End > ownEnd) return null;", read);
            Assert.Contains("if (ComposeBody.FromWord(range.Text) != selection.Text) return false;", replace);
            Assert.DoesNotContain("Body =", editor);

            var controller = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "TaskPane", "Chat", "ChatController.cs"));
            Assert.Contains("ComposeDocument.ReplaceSelection(inspector, draft.Selection, draft.Text)", controller);
            Assert.Contains("_reviews.TryGetValue(reviewId, out var points)", controller);   // Apply uses the host's points
        }
    }
}
