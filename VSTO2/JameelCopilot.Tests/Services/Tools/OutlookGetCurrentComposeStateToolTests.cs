using System;
using System.Threading;
using System.Threading.Tasks;
using JameelCopilot.Services.Tools;
using Xunit;

namespace JameelCopilot.Tests.Services.Tools
{
    public class OutlookGetCurrentComposeStateToolTests
    {
        [Fact]
        public async Task Execute_ProjectsComposeStateToJson()
        {
            var surface = new Surface
            {
                State = new ComposeStateResult
                {
                    Subject = "Q4 review",
                    SenderName = "Alice",
                    SenderEmail = "alice@example.com",
                    BodyPlaintext = "hi",
                    BodyTruncated = false,
                    ToRecipients = new[] { "bob@example.com" },
                    CcRecipients = new string[0],
                    BccRecipients = new string[0],
                    Attachments = new AttachmentSummary[0],
                }
            };
            var tool = new OutlookGetCurrentComposeStateTool();
            var json = await tool.ExecuteAsync("{\"include_full_body\":false}", surface, CancellationToken.None);
            Assert.Contains("\"subject\":\"Q4 review\"", json);
            Assert.Contains("\"sender_email\":\"alice@example.com\"", json);
            Assert.Contains("\"body_truncated\":false", json);
            Assert.Contains("\"to\":[\"bob@example.com\"]", json);
        }

        // What the user wrote and the message being replied to are two fields:
        // the second is not the user's draft.
        [Fact]
        public async Task Execute_ReturnsTheQuotedMessageApartFromTheUsersText()
        {
            var surface = new Surface
            {
                State = new ComposeStateResult
                {
                    BodyPlaintext = "I will send it today.",
                    QuotedPlaintext = "From: Sara Example",
                }
            };
            var json = await new OutlookGetCurrentComposeStateTool().ExecuteAsync("{}", surface, CancellationToken.None);

            Assert.Contains("\"body_plaintext\":\"I will send it today.\"", json);
            Assert.Contains("\"quoted_plaintext\":\"From: Sara Example\"", json);
            var description = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false).ToString();
            Assert.Contains("body_plaintext is what the user has written", description);
        }

        [Fact]
        public async Task Execute_IncludesInReplyToWhenPresent()
        {
            var surface = new Surface
            {
                State = new ComposeStateResult
                {
                    Subject = "RE: Q4",
                    BodyPlaintext = "",
                    ToRecipients = new string[0],
                    CcRecipients = new string[0],
                    BccRecipients = new string[0],
                    Attachments = new AttachmentSummary[0],
                    InReplyTo = new InReplyTo
                    {
                        ThreadTopic = "Q4 review thread",
                        LastNMessages = new[]
                        {
                            new ThreadMessage
                            {
                                From = "bob@example.com",
                                ReceivedAt = DateTimeOffset.Parse("2026-05-10T12:00:00Z"),
                                Snippet = "lgtm",
                            }
                        }
                    }
                }
            };
            var tool = new OutlookGetCurrentComposeStateTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"in_reply_to\"", json);
            Assert.Contains("\"thread_topic\":\"Q4 review thread\"", json);
            Assert.Contains("\"snippet\":\"lgtm\"", json);
        }

        [Fact]
        public async Task Execute_PassesIncludeFullBodyFlagToSurface()
        {
            bool? observed = null;
            var surface = new Surface
            {
                StateFactory = full =>
                {
                    observed = full;
                    return new ComposeStateResult
                    {
                        Subject = "x",
                        BodyPlaintext = "",
                        ToRecipients = new string[0],
                        CcRecipients = new string[0],
                        BccRecipients = new string[0],
                        Attachments = new AttachmentSummary[0],
                    };
                }
            };
            var tool = new OutlookGetCurrentComposeStateTool();
            await tool.ExecuteAsync("{\"include_full_body\":true}", surface, CancellationToken.None);
            Assert.True(observed);
        }

        private sealed class Surface : MinimalSurface
        {
            public ComposeStateResult State { get; set; }
            public Func<bool, ComposeStateResult> StateFactory { get; set; }
            public override ComposeStateResult GetCurrentComposeState(bool includeFullBody)
                => StateFactory != null ? StateFactory(includeFullBody) : State;
        }
    }
}
