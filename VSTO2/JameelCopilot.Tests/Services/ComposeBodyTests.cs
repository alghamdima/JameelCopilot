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
    }
}
