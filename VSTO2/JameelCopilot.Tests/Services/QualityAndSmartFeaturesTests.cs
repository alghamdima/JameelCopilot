using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services;
using JameelCopilot.Services.Export;
using JameelCopilot.Services.Prefs;
using JameelCopilot.Services.Tools;
using JameelCopilot.TaskPane.InboxCopilot;
using JameelCopilot.Tests.Helpers;
using Xunit;
using ActionType = JameelCopilot.Services.CustomApiChatService.ActionType;

namespace JameelCopilot.Tests.Services
{
    public class QualityAndSmartFeaturesTests
    {
        [Fact]
        public void ChatArchive_PreservesChipLineMark_AcrossSaveAndRestore()
        {
            var store = new FakeUserStore();
            var archive = new ChatArchive(store);
            var chipMsg = new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "user"),
                new JProperty("content", "Start my day."),
                new JProperty(StructuredAnswers.ChipLineMark, true));
            var assistantMsg = new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "assistant"),
                new JProperty("content", "To do: Review contract"));
            var arabicUserMsg = new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "user"),
                new JProperty("content", "\u0644\u062E\u0635 \u0627\u0644\u0627\u064A\u0645\u064A\u0644"));

            var saved = new SavedChat
            {
                Scope = ChatArchive.MailboxScope,
                Title = "Day",
                Messages = ChatArchive.FromHistory(new[] { arabicUserMsg, assistantMsg, chipMsg, assistantMsg }),
            };
            archive.Save(saved);

            var reloaded = archive.Load(saved.Id);
            var restoredHistory = ChatArchive.ToHistory(reloaded);

            // Even though the most recent user line was an English chip ("Start my day."),
            // its ChipLineMark survived reload so ReplyLanguage still detects Arabic!
            Assert.Equal("Arabic", StructuredAnswers.ReplyLanguage(restoredHistory));
        }

        [Fact]
        public void StructuredAnswers_ParseObject_RecoversJsonWithOuterBraces()
        {
            const string noisy = "Note {internal}: here is the result:\n{\"todos\":[{\"text\":\"Send report\"}],\"topics\":[]}\nThanks {done}.";
            var parsed = StructuredAnswers.ParseObject(noisy);

            Assert.NotNull(parsed);
            Assert.Equal("Send report", (string)parsed["todos"]?[0]?["text"]);
        }

        [Fact]
        public void DayBrief_EmailContext_IncludesCcAndExtractedAttachmentText()
        {
            var msg = new MessageDetail
            {
                Id = "m1",
                From = "Sara Example <sara@example.com>",
                To = new[] { "Team <team@example.com>" },
                Cc = new[] { "Omar Example <omar@example.com>" },
                Subject = "Steering Committee MoM",
                ReceivedAt = new DateTimeOffset(2026, 11, 15, 9, 0, 0, TimeSpan.Zero),
                BodyPlaintext = "Please find attached the minutes.",
                Attachments = new[]
                {
                    new AttachmentSummary
                    {
                        Filename = "MoM_Nov15.pdf",
                        ContentPlaintext = "Action 1: Omar Example to approve vendor SLA by 2026-11-20.",
                    },
                },
            };

            var ctx = DayBrief.EmailContext(msg, "Omar Example", "omar@example.com");

            Assert.Contains("Cc: Omar Example <omar@example.com>", ctx);
            Assert.Contains("MoM_Nov15.pdf", ctx);
            Assert.Contains("Action 1: Omar Example to approve vendor SLA", ctx);
        }

        [Theory]
        [InlineData("SAR 12,450.75", 12450.75)]
        [InlineData("1,500 SAR", 1500.0)]
        [InlineData("SR 850.50", 850.50)]
        [InlineData("2,300 \u0631.\u0633", 2300.0)]
        [InlineData("(SAR 400.00)", -400.0)]
        public void ExcelCellCoercer_ParsesSaudiRiyalAmounts(string raw, double expected)
        {
            var coerced = ExcelCellCoercer.Coerce(new JValue(raw), ExcelColumnType.Currency);

            Assert.IsType<double>(coerced);
            Assert.Equal(expected, (double)coerced, 2);
        }

        [Fact]
        public void ComposeBody_DoesNotStripDraftParagraphMentioningColleagueEmail()
        {
            const string draft = "Hi Sara,\n\nPlease share the signed PO with billing@example.com before Thursday.";
            var cleaned = ComposeBody.StripSignatureAndFooter(draft);

            Assert.Contains("billing@example.com before Thursday.", cleaned);
        }

        [Fact]
        public void ReplyIntentInstructions_TailorApproveClarifyAndDecline()
        {
            var state = ComposePrompt.ReplyTo(new MessageDetail
            {
                Subject = "Budget request",
                From = "Sara Example <sara@example.com>",
                Cc = new[] { "Finance <finance@example.com>" },
                BodyPlaintext = "Can we proceed with phase 2?",
            }, "Omar Example", "omar@example.com");

            Assert.Equal(new[] { "Finance <finance@example.com>" }, state.CcRecipients);
            Assert.Contains("approval", ComposePrompt.Instructions(ActionType.Revise, state, "approve"), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("clarifying questions", ComposePrompt.Instructions(ActionType.Revise, state, "clarify"), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("decline", ComposePrompt.Instructions(ActionType.Revise, state, "decline"), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void BuildPeopleHints_DeduplicatesOrgAndThreadParticipants_AndOmitsSelf()
        {
            var org = new OrgContextInfo
            {
                User = new OrgPerson { Name = "Omar Example", Email = "omar@example.com", JobTitle = "Head of Data" },
                Manager = new OrgPerson { Name = "Khalid Example", Email = "khalid@example.com", JobTitle = "VP Technology" },
                DirectReports = new[]
                {
                    new OrgPerson { Name = "Lina Example", Email = "lina@example.com", JobTitle = "Data Engineer" },
                },
                Peers = new OrgPerson[0],
            };

            var hints = InboxCopilotPromptBuilder.BuildPeopleHints(
                org,
                new[] { "Sara Example <sara@example.com>", "Omar Example <omar@example.com>", "Lina Example <lina@example.com>" });

            Assert.Equal(3, hints.Count);
            Assert.DoesNotContain(hints, h => string.Equals((string)h["email"], "omar@example.com", StringComparison.OrdinalIgnoreCase));
            Assert.Equal("sara@example.com", (string)hints[0]["email"]);
        }

        [Fact]
        public void IsDraftLikeAnswer_ShowsReplyButtonsOnlyOnActualEmailDrafts()
        {
            Assert.True(InboxCopilotController.IsDraftLikeAnswer("Write a short reply confirming receipt", "Thank you, received and under review.", false));
            Assert.True(InboxCopilotController.IsDraftLikeAnswer("\u0627\u0643\u062A\u0628 \u0631\u062F \u0645\u062E\u062A\u0635\u0631", "\u0634\u0643\u0631\u0627\u064B \u0644\u0643\u0645\u060C \u062A\u0645 \u0627\u0644\u0627\u0633\u062A\u0644\u0627\u0645.", false));
            Assert.True(InboxCopilotController.IsDraftLikeAnswer("anything", "Translated body", isTranslateChip: true));

            // Markdown tables and report headings are never email replies.
            Assert.False(InboxCopilotController.IsDraftLikeAnswer("Summarize my emails", "### Summary\n\n| From | Subject |\n|---|---|\n| Sara | PO |", false));
            Assert.False(InboxCopilotController.IsDraftLikeAnswer("When is the deadline?", "The deadline mentioned in the email is Thursday 20 Nov.", false));
        }

        [Fact]
        public void CleanDraftForEmail_StripsConversationalWrapper_SubjectLine_BoldMarkers_AndSignOff()
        {
            const string screenshot1 =
                "\u0645\u0627 \u0642\u062F\u0631\u062A \u0623\u062C\u062F \u0628\u0631\u064A\u062F\u064A \u0634\u0647\u0646\u0648\u0627\u0632 \u0648\u0627\u0645\u0627\u0646 \u0627\u0644\u0644\u0647 \u0641\u064A \u062F\u0644\u064A\u0644 \u0627\u0644\u0639\u0646\u0627\u0648\u064A\u0646\u060C \u0644\u0643\u0646 \u0647\u0630\u0627 \u0646\u0635 \u0627\u0644\u0625\u064A\u0645\u064A\u0644:\n\n" +
                "---\n\n" +
                "Hi Shanawaz & Aaman,\n\n" +
                "Following up on the Sprinklr AI Integration POC connectivity setup:\n\n" +
                "- **Shanawaz:** handling the F5 VIP.\n" +
                "- **Aaman:** handling the DNS in Qodadi.\n\n" +
                "Please keep us updated on progress.\n\n" +
                "---\n\n" +
                "\u0644\u0648 \u062A\u0628\u064A \u0623\u0639\u062F\u0644 \u0623\u064A \u0634\u064A \u0623\u0648 \u0623\u0636\u064A\u0641 \u062A\u0641\u0627\u0635\u064A\u0644\u060C \u0642\u0648\u0644\u064A.";

            var cleaned1 = ComposeBody.CleanDraftForEmail(screenshot1, out var subj1);

            Assert.Null(subj1);
            Assert.StartsWith("Hi Shanawaz & Aaman,", cleaned1);
            Assert.Contains("- Shanawaz: handling the F5 VIP.", cleaned1);
            Assert.DoesNotContain("**", cleaned1);
            Assert.DoesNotContain("---", cleaned1);
            Assert.DoesNotContain("\u0645\u0627 \u0642\u062F\u0631\u062A", cleaned1);
            Assert.DoesNotContain("\u0644\u0648 \u062A\u0628\u064A", cleaned1);

            const string screenshot3 =
                "Subject: Central Bank Requests - Data Team Handling\n\n" +
                "Dear Fatima,\n\n" +
                "This is to confirm that any requests received from the Central Bank will be directed to the Data team for handling.\n\n" +
                "Please feel free to route any Central Bank-related requests to us, and we will ensure timely follow-up.\n\n" +
                "Best regards,";

            var cleaned3 = ComposeBody.CleanDraftForEmail(screenshot3, out var subj3);

            Assert.Equal("Central Bank Requests - Data Team Handling", subj3);
            Assert.StartsWith("Dear Fatima,", cleaned3);
            Assert.EndsWith("ensure timely follow-up.", cleaned3);
            Assert.DoesNotContain("Subject:", cleaned3);
            Assert.DoesNotContain("Best regards", cleaned3);
        }

        [Fact]
        public void MeetingTimes_IgnoreTentative_TreatsUnacceptedSlotsAsAvailable()
        {
            var day = new DateTime(2026, 11, 15);   // Sunday
            // Slot 20..22 (10:00..11:00) is '1' (Tentative / unaccepted), rest is '2' (Busy).
            var chars = new string('2', 48).ToCharArray();
            chars[20] = '1';
            chars[21] = '1';
            var grid = new string(chars);

            var strict = MeetingTimes.Find(new[] { grid }, day, day.AddHours(9), day.AddHours(17), 60,
                new TimeSpan(9, 0, 0), new TimeSpan(17, 0, 0), MeetingTimes.WorkDays("sun-thu"), ignoreTentative: false);
            Assert.Empty(strict);

            var withTentativeIgnored = MeetingTimes.Find(new[] { grid }, day, day.AddHours(9), day.AddHours(17), 60,
                new TimeSpan(9, 0, 0), new TimeSpan(17, 0, 0), MeetingTimes.WorkDays("sun-thu"), ignoreTentative: true);
            Assert.Single(withTentativeIgnored);
            Assert.Equal(day.AddHours(10), withTentativeIgnored[0].Item1);
            Assert.Equal(day.AddHours(11), withTentativeIgnored[0].Item2);
        }

        [Fact]
        public void MeetingTimes_ResolvePeriodRange_HandlesTodayThisWeekNextWeekAndCustom()
        {
            var wedMorning = new DateTime(2026, 11, 18, 10, 0, 0);   // Wednesday

            MeetingTimes.ResolvePeriodRange("today", "", wedMorning, out var fToday, out var tToday);
            Assert.Equal(wedMorning.Date, fToday.Date);
            Assert.Equal(wedMorning.Date, tToday.Date);

            MeetingTimes.ResolvePeriodRange("this_week", "", wedMorning, out var fThis, out var tThis);
            Assert.Equal(new DateTime(2026, 11, 18), fThis.Date);
            Assert.Equal(new DateTime(2026, 11, 19), tThis.Date);   // Thursday of this week

            MeetingTimes.ResolvePeriodRange("next_week", "", wedMorning, out var fNext, out var tNext);
            Assert.Equal(new DateTime(2026, 11, 22), fNext.Date);   // Next Sunday
            Assert.Equal(new DateTime(2026, 11, 26), tNext.Date);   // Next Thursday

            MeetingTimes.ResolvePeriodRange("custom", "2026-11-25", wedMorning, out var fCust, out var tCust);
            Assert.Equal(new DateTime(2026, 11, 25, 8, 0, 0), fCust);
            Assert.Equal(new DateTime(2026, 11, 25, 18, 0, 0), tCust);
        }
    }
}
