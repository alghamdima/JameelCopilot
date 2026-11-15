using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using JameelCopilot.Services;
using JameelCopilot.Services.Tools;
using JameelCopilot.TaskPane.Chat;
using JameelCopilot.Tests.Helpers;
using Xunit;
using ActionType = JameelCopilot.Services.CustomApiChatService.ActionType;

namespace JameelCopilot.Tests.TaskPane
{
    /// <summary>
    /// One Jameel Copilot pane. The owner saw two look-alike Inbox panes and a tab of
    /// buttons; now there is one page, and writing a reply ends in a card the
    /// user decides on.
    /// </summary>
    public class OnePaneTests
    {
        private static string Source(params string[] path) =>
            File.ReadAllText(RepoFiles.Find(new[] { "VSTO2", "JameelCopilot" }.Concat(path).ToArray()));

        // -- Draft a reply from the Inbox ----------------------------------------

        private static MessageDetail Selected() => new MessageDetail
        {
            Id = "m1",
            Subject = "Budget review",
            From = "Sara Example",
            BodyPlaintext = "Dear Omar,\n\nPlease send the budget file.\n\nRegards,\nSara Example",
        };

        [Fact]
        public void ReplyTo_IsAReplyWindowWithNothingTyped_WrittenByTheOwner()
        {
            var state = ComposePrompt.ReplyTo(Selected(), "Omar Example", "omar@example.com");

            Assert.Equal("Omar Example", state.SenderName);
            Assert.Equal("omar@example.com", state.SenderEmail);
            Assert.Equal(new[] { "Sara Example" }, state.ToRecipients);
            Assert.Equal("RE: Budget review", state.Subject);
            Assert.Equal("", state.BodyPlaintext);
            Assert.StartsWith("From: Sara Example\nSubject: Budget review\n\nDear Omar,", state.QuotedPlaintext);

            // So the action writes the owner's reply, and the model is told who writes to whom.
            Assert.True(ComposePrompt.WritesReply(ActionType.Revise, state));
            var context = ComposePrompt.Context(state, includeDraft: false);
            Assert.StartsWith("From (the user, who is writing this email): Omar Example <omar@example.com>\nTo: Sara Example\n", context);
            Assert.Contains("\n> Please send the budget file.", context);
        }

        [Theory]
        [InlineData("RE: Budget review", "RE: Budget review")]
        [InlineData("re: Budget review", "re: Budget review")]
        [InlineData("", "RE:")]
        public void ReplyTo_DoesNotStackReplyPrefixes(string subject, string expected)
        {
            var message = Selected();
            message.Subject = subject;

            Assert.Equal(expected, ComposePrompt.ReplyTo(message, "Omar Example", "").Subject);
        }

        [Fact]
        public void ReplyTo_NothingSelected_IsAnEmptyWindow()
        {
            var state = ComposePrompt.ReplyTo(null, null, null);

            Assert.Equal("", state.QuotedPlaintext);
            Assert.Empty(state.ToRecipients);
            Assert.False(ComposePrompt.WritesReply(ActionType.Revise, state));
        }

        [Fact]
        public void InboxReplyDraft_RunsWithoutTools_AndOnlyAButtonOpensAReplyWindow()
        {
            var controller = Source("TaskPane", "InboxCopilot", "InboxCopilotController.cs");

            Assert.Contains("var ctx = CustomApiChatService.NewActionContext(action);", controller);
            Assert.Contains("ComposePrompt.ReplyTo(message, ownerName, ownerEmail)", controller);
            // The reply window is opened in one place: the handler of the card's buttons.
            Assert.Single(Regex.Matches(controller, @"OpenReplyWithText\(").Cast<Match>());
            var handler = controller.Substring(controller.IndexOf("private void HandleDraftAction(", StringComparison.Ordinal));
            Assert.Contains("_surface.OpenReplyWithText(draft.Message.Id, what == \"reply_all\", draft.Text)", handler);
        }

        // Jameel Copilot opens Outlook's reply window with the text in it. It never sends, saves or deletes.
        [Fact]
        public void OpenReply_PutsTheTextInTheEditor_AndSendsNothing()
        {
            var surface = Source("Services", "Tools", "LiveOutlookSurface.cs");
            var method = surface.Substring(surface.IndexOf("public bool OpenReplyWithText(", StringComparison.Ordinal));
            method = method.Substring(0, method.IndexOf("public IReadOnlyList<FolderResult> ListFolders()", StringComparison.Ordinal));

            Assert.Contains("ComposeDocument.InsertAtTop(inspector, text)", method);
            Assert.Contains("reply.Display(false);", method);
            Assert.DoesNotContain(".Send(", method);
            Assert.DoesNotContain(".Save(", method);
            Assert.DoesNotContain("Body =", method);
            Assert.DoesNotContain(".Send(", surface);
        }

        // -- draft cards: the scripts both panes send ----------------------------

        [Fact]
        public void FinalizeScript_EndsACardWithItsButtons()
        {
            Assert.Equal(
                "jameelcopilot.finalizeDraft(\"draft_3\", [{\"label\":\"Discard\",\"action\":\"discard\"}], {\"status\":\"stopped\",\"error\":false});",
                ComposeChips.FinalizeScript("draft_3", ComposeChips.DiscardOnly(), "stopped", null, false));

            var script = ComposeChips.FinalizeScript("d\"1", ComposeChips.ReplyWindowButtons(), "not in an email yet", "Fill the [brackets].", true);
            Assert.StartsWith("jameelcopilot.finalizeDraft(\"d\\\"1\", [{\"label\":\"Open reply\",\"action\":\"reply\",\"primary\":true},{\"label\":\"Reply all\",\"action\":\"reply_all\"},", script);
            Assert.EndsWith("{\"status\":\"not in an email yet\",\"error\":true,\"note\":\"Fill the [brackets].\"});", script);
        }

        [Fact]
        public void MarkScript_SaysWhereTheDraftWent_OrThatItCouldNot()
        {
            Assert.Equal("jameelcopilot.markDraft(\"draft_1\", \"in the email\", true);",
                ComposeChips.MarkScript("draft_1", true, "in the email", "could not reach the email"));
            Assert.Equal("jameelcopilot.markDraft(\"draft_1\", \"could not reach the email\", false);",
                ComposeChips.MarkScript("draft_1", false, "in the email", "could not reach the email"));
        }

        // -- one pane -------------------------------------------------------------

        [Fact]
        public void InboxPane_AnswersQuestionsAndWritesReports_WithBothTexts()
        {
            Assert.Contains(
                "SystemInstructions = SystemPromptBuilder.Pane(\"inbox\") + \"\\n\\n\" + SystemPromptBuilder.Pane(\"reports\"),",
                Source("TaskPane", "InboxCopilot", "InboxCopilotController.cs"));
            Assert.StartsWith("When the user asks for a report", SystemPromptBuilder.Pane("reports"));
            Assert.DoesNotContain("Reports pane", SystemPromptBuilder.Pane("reports"));
        }

        [Fact]
        public void Ribbon_HasOneJcButtonInTheInbox_NoSecondPane()
        {
            var ribbon = Source("Ribbon.xml");

            Assert.DoesNotContain("btnReports", ribbon);
            Assert.DoesNotContain("OnReportsClick", ribbon + Source("Ribbon.cs"));
            Assert.DoesNotContain("Jameel Copilot Reports", Source("ThisAddIn.cs"));
            Assert.False(Directory.Exists(Path.Combine(Path.GetDirectoryName(RepoFiles.Find("VSTO2", "JameelCopilot", "ThisAddIn.cs")), "TaskPane", "InboxReports")));
        }

        [Fact]
        public void ComposePane_IsThePageAlone()
        {
            var pane = Source("TaskPane", "AITaskPane.cs") + Source("TaskPane", "AITaskPane.Designer.cs");

            Assert.DoesNotContain("TabControl", pane);
            Assert.DoesNotContain("btnProofread", pane);
            Assert.Contains("new ChatController(", pane);
        }

        // An upgrade replaces the page's files; the browser must not run the old
        // script against the new host. Every version asks for its own address.
        [Fact]
        public void ThePage_IsAskedForByVersion()
        {
            var version = typeof(WebView2Bootstrap).Assembly.GetName().Version.ToString();
            var index = Source("WebUI", "index.html");
            Assert.Contains("href=\"styles.css\"", index);
            Assert.Contains("src=\"markdown.js\"", index);
            Assert.Contains("src=\"chat.js\"", index);

            var versioned = WebView2Bootstrap.VersionedIndex(index, "1.2.3.0");

            Assert.Contains("href=\"styles.css?v=1.2.3.0\"", versioned);
            Assert.Contains("src=\"markdown.js?v=1.2.3.0\"", versioned);
            Assert.Contains("src=\"chat.js?v=1.2.3.0\"", versioned);
            Assert.DoesNotContain("src=\"chat.js\"", versioned);
            Assert.Equal("https://" + WebView2Bootstrap.VirtualHost + "/index.html?v=" + version, WebView2Bootstrap.PageUrl);
            Assert.True(WebView2Bootstrap.IsTrustedSource(WebView2Bootstrap.PageUrl));
            Assert.Contains("Navigate(WebView2Bootstrap.PageUrl)", Source("TaskPane", "Chat", "ChatController.cs"));
            Assert.Contains("Navigate(WebView2Bootstrap.PageUrl)", Source("TaskPane", "InboxCopilot", "InboxCopilotController.cs"));
        }

        // "When Jameel Copilot tells me about an email, let me click it and open it."
        [Fact]
        public void AListedEmail_OpensInOutlook_OnlyIfJcListedIt()
        {
            var surface = Source("Services", "Tools", "LiveOutlookSurface.cs");
            var method = surface.Substring(surface.IndexOf("public bool OpenMessage(", StringComparison.Ordinal));
            method = method.Substring(0, method.IndexOf("public void GetOwner(", StringComparison.Ordinal));

            Assert.Contains("_ids.Resolve(messageId)", method);   // unknown ids throw: only Jameel Copilot's own resolve
            Assert.Contains("catch (KeyNotFoundException) { return false; }", method);
            Assert.Contains(".Display(false)", method);
            foreach (var verb in new[] { ".Send(", ".Delete(", ".Move(", ".Save(" }) Assert.DoesNotContain(verb, method);

            foreach (var controller in new[] { Source("TaskPane", "Chat", "ChatController.cs"), Source("TaskPane", "InboxCopilot", "InboxCopilotController.cs") })
            {
                Assert.Contains("case \"open_message\":", controller);
                Assert.Contains("_surface.OpenMessage(id)", controller);
            }
            Assert.Contains("listed under your answer", SystemPromptBuilder.Pane("inbox"));
        }

        // -- what the chat is told about writing email text -----------------------

        // The owner's screenshot: a chat-drafted reply with a "Subject:" line, two languages and a signed name.
        [Fact]
        public void ChatPanes_GiveEmailTextOnly_WithNoSignature()
        {
            var inbox = SystemPromptBuilder.Pane("inbox");
            Assert.Contains("give the email text only: no subject line, nothing before or after it, in one language", inbox);
            Assert.Contains("no signature: no closing line such as \"Best regards\", no name, title or contact details.", inbox);

            Assert.Contains("answer with the email text only, nothing before or after it, so it can be inserted as it is.",
                SystemPromptBuilder.Pane("compose"));
        }
    }
}
