using JameelCopilot.Services;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>
    /// A compose window holds what the user typed, their signature and the
    /// message being replied to. Only the first is the user's email: a reply
    /// with nothing typed yet has no text of the user's at all.
    /// </summary>
    public class ComposeBodyTests
    {
        private const string Quoted =
            "From: Sara Example <sara@example.com>\r\n" +
            "Sent: Thursday, November 12, 2026 9:15 AM\r\n" +
            "To: Omar Example <omar@example.com>\r\n" +
            "Subject: Budget review\r\n" +
            "\r\n" +
            "Dear Omar,\r\n\r\nPlease send the budget file.\r\n\r\nRegards,\r\nSara Example\r\n";

        [Fact]
        public void FromPlainText_Reply_SplitsWhatTheUserWroteFromTheQuotedMessage()
        {
            var body = ComposeBody.FromPlainText("I will send it today.\r\n\r\n" + Quoted);

            Assert.Equal("I will send it today.", body.Own);
            Assert.StartsWith("From: Sara Example", body.Quoted);
            Assert.Contains("Please send the budget file.", body.Quoted);
        }

        [Fact]
        public void FromPlainText_ReplyWithNothingTyped_HasNoTextOfTheUsers()
        {
            var body = ComposeBody.FromPlainText("\r\n\r\n" + Quoted);

            Assert.Equal("", body.Own);
            Assert.Contains("Dear Omar,", body.Quoted);
        }

        // Outlook writes the header in its own display language: From / Sent / To / Subject in Arabic.
        [Fact]
        public void FromPlainText_ArabicHeader_SplitsTheSameWay()
        {
            var quoted =
                "من: Sara Example <sara@example.com>\r\n" +
                "تم الإرسال: 12 نوفمبر 2026\r\n" +
                "إلى: Omar Example\r\n" +
                "الموضوع: Budget review\r\n\r\nPlease send the budget file.";

            var body = ComposeBody.FromPlainText("I will send it today.\r\n\r\n" + quoted);

            Assert.Equal("I will send it today.", body.Own);
            Assert.StartsWith("من: Sara Example", body.Quoted);
        }

        [Theory]
        [InlineData("-----Original Message-----\r\nFrom: Sara Example\r\nSent: Thursday\r\n\r\nPlease send the budget file.")]
        [InlineData("________________________________\r\nFrom: Sara Example\r\nSent: Thursday\r\n\r\nPlease send the budget file.")]
        public void FromPlainText_SeparatorLine_BelongsToTheQuotedMessage(string quoted)
        {
            var body = ComposeBody.FromPlainText("I will send it today.\r\n\r\n" + quoted);

            Assert.Equal("I will send it today.", body.Own);
            Assert.Equal(quoted, body.Quoted);
        }

        [Fact]
        public void FromPlainText_NewEmail_IsAllTheUsers()
        {
            var body = ComposeBody.FromPlainText("Hello Sara,\r\n\r\nFrom: my side everything is ready.\r\nThanks");

            Assert.Equal("Hello Sara,\r\n\r\nFrom: my side everything is ready.\r\nThanks", body.Own);
            Assert.Equal("", body.Quoted);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("  \r\n ")]
        public void FromPlainText_Nothing_IsEmpty(string text)
        {
            var body = ComposeBody.FromPlainText(text);

            Assert.Equal("", body.Own);
            Assert.Equal("", body.Quoted);
        }

        // Word ends a paragraph with \r, a table cell with \r\a, a line break
        // with \v and stands a picture in as \x01.
        [Fact]
        public void FromWord_TurnsWordsMarksIntoPlainLines()
        {
            Assert.Equal("Hello Sara,\n\nLine one\nLine two\nA\nB", ComposeBody.FromWord("\rHello Sara,\r\rLine one\vLine two\r\u0001A\r\aB\r\a\r"));
            Assert.Equal("", ComposeBody.FromWord(null));
        }

        // Word takes \r as the end of a paragraph; \r\n would give two.
        [Fact]
        public void ToWord_EveryLineBecomesOneParagraph()
        {
            Assert.Equal("Hello Sara,\r\rI will send it today.", ComposeBody.ToWord("Hello Sara,\r\n\r\nI will send it today.\n\n"));
            Assert.Equal("a\rb", ComposeBody.ToWord("a\nb"));
            Assert.Equal("", ComposeBody.ToWord(null));
        }

        [Fact]
        public void StripSignatureAndFooter_RemovesManualCorporateSignatureAndClassificationFooter()
        {
            var emailWithSig =
                "Dear Team,\r\n\r\n" +
                "Please find attached the MoM for your review.\r\n\r\n" +
                "Regards,\r\n" +
                "Saud Example\r\n" +
                "PMO Senior Specialist\r\n" +
                "+966 12 5550000\r\n" +
                "saud@example.com\r\n" +
                "PO Box: 54609\r\n" +
                "This email is classified as Internal Use - \u0644\u0644\u0625\u0633\u062A\u062E\u062F\u0627\u0645 \u0627\u0644\u062F\u0627\u062E\u0644\u064A";

            var clean = ComposeBody.StripSignatureAndFooter(emailWithSig);

            Assert.Equal("Dear Team,\r\n\r\nPlease find attached the MoM for your review.", clean);
        }

        [Fact]
        public void StripThreadSignatures_StripsSignaturesFromEveryMessageInLargeThread()
        {
            var thread =
                "Approved from my side.\r\n\r\n" +
                "BR\r\n" +
                "Omar Example\r\n" +
                "+966 55 555 0000\r\n\r\n" +
                "From: Sara Example <sara@example.com>\r\n" +
                "Sent: Sunday, November 15, 2026 3:35 PM\r\n" +
                "To: Omar Example <omar@example.com>\r\n" +
                "Subject: RE: Budget\r\n\r\n" +
                "Kindly review the attached file.\r\n\r\n" +
                "Sara Example\r\n" +
                "IT Analyst\r\n" +
                "+966 12 5551111\r\n" +
                "PO Box: 00000\r\n" +
                "This email is classified as Internal Use";

            var cleaned = ComposeBody.StripThreadSignatures(thread);

            Assert.Contains("Approved from my side.", cleaned);
            Assert.Contains("From: Sara Example <sara@example.com>", cleaned);
            Assert.Contains("Kindly review the attached file.", cleaned);
            Assert.DoesNotContain("PO Box", cleaned);
            Assert.DoesNotContain("5551111", cleaned);
            Assert.DoesNotContain("classified as", cleaned);
        }
    }
}
