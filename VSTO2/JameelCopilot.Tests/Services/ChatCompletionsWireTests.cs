using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    public class ChatCompletionsWireTests
    {
        private static JObject Msg(string role, string content) =>
            new JObject(new JProperty("type", "message"), new JProperty("role", role), new JProperty("content", content));

        private static JObject Call(string id, string name, string args) =>
            new JObject(new JProperty("type", "function_call"), new JProperty("call_id", id),
                new JProperty("name", name), new JProperty("arguments", args));

        private static JObject Output(string id, string output) =>
            new JObject(new JProperty("type", "function_call_output"), new JProperty("call_id", id),
                new JProperty("output", output));

        private static string Json(JToken token) => token.ToString(Formatting.None);

        [Fact]
        public void ToMessages_SystemThenUserThenAssistant()
        {
            var messages = ChatCompletionsWire.ToMessages("S", new[] { Msg("user", "hi"), Msg("assistant", "hello") });

            Assert.Equal(
                "[{\"role\":\"system\",\"content\":\"S\"},{\"role\":\"user\",\"content\":\"hi\"},{\"role\":\"assistant\",\"content\":\"hello\"}]",
                Json(messages));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ToMessages_EmptyInstructions_NoSystemMessage(string instructions)
        {
            var messages = ChatCompletionsWire.ToMessages(instructions, new[] { Msg("user", "hi") });

            Assert.Equal("[{\"role\":\"user\",\"content\":\"hi\"}]", Json(messages));
        }

        [Fact]
        public void ToMessages_AssistantTextThenTwoCalls_BecomesOneAssistantMessage()
        {
            var history = new[]
            {
                Msg("user", "find it"),
                Msg("assistant", "ok"),
                Call("c1", "n1", "{}"), Call("c2", "n2", "{\"a\":1}"),
                Output("c1", "o1"), Output("c2", "o2"),
            };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(
                "[{\"role\":\"user\",\"content\":\"find it\"},"
                + "{\"role\":\"assistant\",\"content\":\"ok\",\"tool_calls\":["
                + "{\"id\":\"c1\",\"type\":\"function\",\"function\":{\"name\":\"n1\",\"arguments\":\"{}\"}},"
                + "{\"id\":\"c2\",\"type\":\"function\",\"function\":{\"name\":\"n2\",\"arguments\":\"{\\\"a\\\":1}\"}}]},"
                + "{\"role\":\"tool\",\"tool_call_id\":\"c1\",\"content\":\"o1\"},"
                + "{\"role\":\"tool\",\"tool_call_id\":\"c2\",\"content\":\"o2\"}]",
                Json(messages));
        }

        [Fact]
        public void ToMessages_CallsWithoutPrecedingText_ContentIsNull()
        {
            var history = new[] { Msg("user", "find it"), Call("c1", "n1", "{}"), Output("c1", "o1") };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(
                "[{\"role\":\"user\",\"content\":\"find it\"},"
                + "{\"role\":\"assistant\",\"content\":null,\"tool_calls\":["
                + "{\"id\":\"c1\",\"type\":\"function\",\"function\":{\"name\":\"n1\",\"arguments\":\"{}\"}}]},"
                + "{\"role\":\"tool\",\"tool_call_id\":\"c1\",\"content\":\"o1\"}]",
                Json(messages));
        }

        // An API rejects a tool call that has no result, so a call left without
        // its output (a turn cancelled mid-dispatch) is not sent back.
        [Fact]
        public void ToMessages_CallWithoutOutput_IsDropped()
        {
            var history = new[] { Msg("user", "find it"), Msg("assistant", "ok"), Call("c1", "n1", "{}"), Msg("user", "again") };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(
                "[{\"role\":\"user\",\"content\":\"find it\"},{\"role\":\"assistant\",\"content\":\"ok\"},{\"role\":\"user\",\"content\":\"again\"}]",
                Json(messages));
        }

        [Fact]
        public void ToMessages_SecondToolRound_StartsANewAssistantMessage()
        {
            var history = new[]
            {
                Msg("user", "find it"),
                Call("c1", "n1", "{}"), Output("c1", "o1"),
                Msg("assistant", "one more"),
                Call("c2", "n2", "{}"), Output("c2", "o2"),
            };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(5, messages.Count);
            Assert.Equal("c1", (string)messages[1]["tool_calls"][0]["id"]);
            Assert.Equal("tool", (string)messages[2]["role"]);
            Assert.Equal("one more", (string)messages[3]["content"]);
            Assert.Equal("c2", (string)messages[3]["tool_calls"][0]["id"]);
            Assert.Equal("c2", (string)messages[4]["tool_call_id"]);
        }

        // Two rounds with no text between them are two steps, not one parallel step.
        [Fact]
        public void ToMessages_TwoRoundsWithoutText_AreTwoAssistantMessages()
        {
            var history = new[]
            {
                Msg("user", "find it"),
                Call("c1", "n1", "{}"), Output("c1", "o1"),
                Call("c2", "n2", "{}"), Output("c2", "o2"),
            };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(
                "[{\"role\":\"user\",\"content\":\"find it\"},"
                + "{\"role\":\"assistant\",\"content\":null,\"tool_calls\":["
                + "{\"id\":\"c1\",\"type\":\"function\",\"function\":{\"name\":\"n1\",\"arguments\":\"{}\"}}]},"
                + "{\"role\":\"tool\",\"tool_call_id\":\"c1\",\"content\":\"o1\"},"
                + "{\"role\":\"assistant\",\"content\":null,\"tool_calls\":["
                + "{\"id\":\"c2\",\"type\":\"function\",\"function\":{\"name\":\"n2\",\"arguments\":\"{}\"}}]},"
                + "{\"role\":\"tool\",\"tool_call_id\":\"c2\",\"content\":\"o2\"}]",
                Json(messages));
        }

        // A model may reuse an id in a later round, or in the same one.
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ToMessages_RepeatedId_EachResultStaysWithItsOwnCall(bool sameRound)
        {
            var history = sameRound
                ? new[] { Msg("user", "q"), Call("c", "n1", "{}"), Call("c", "n2", "{}"), Output("c", "first"), Output("c", "second") }
                : new[] { Msg("user", "q"), Call("c", "n1", "{}"), Output("c", "first"), Call("c", "n2", "{}"), Output("c", "second") };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(new[] { "first", "second" },
                messages.Where(m => (string)m["role"] == "tool").Select(m => (string)m["content"]));
            Assert.Equal(new[] { "n1", "n2" },
                messages.Where(m => m["tool_calls"] != null).SelectMany(m => m["tool_calls"]).Select(c => (string)c["function"]["name"]));
        }

        // Only the calls that have a result are sent back, each with its own.
        [Fact]
        public void ToMessages_RoundWithAMissingResult_KeepsTheAnsweredCallOnly()
        {
            var history = new[] { Msg("user", "q"), Call("c1", "n1", "{}"), Call("c2", "n2", "{}"), Output("c2", "o2") };

            var messages = ChatCompletionsWire.ToMessages(null, history);

            Assert.Equal(3, messages.Count);
            Assert.Equal("c2", (string)Assert.Single(messages[1]["tool_calls"])["id"]);
            Assert.Equal(("c2", "o2"), ((string)messages[2]["tool_call_id"], (string)messages[2]["content"]));
        }

        // A result with no call before it has nothing to answer.
        [Fact]
        public void ToMessages_ResultWithoutACall_IsDropped()
        {
            var messages = ChatCompletionsWire.ToMessages(null, new[] { Msg("user", "q"), Output("c1", "o1") });

            Assert.Equal("[{\"role\":\"user\",\"content\":\"q\"}]", Json(messages));
        }

        [Fact]
        public void ToTools_WrapsEachEntryInFunction()
        {
            var catalog = JArray.Parse(
                "[{\"type\":\"function\",\"name\":\"a\",\"description\":\"d\",\"parameters\":{\"type\":\"object\",\"properties\":{}}}]");

            var tools = ChatCompletionsWire.ToTools(catalog);

            Assert.Equal(
                "[{\"type\":\"function\",\"function\":{\"name\":\"a\",\"description\":\"d\",\"parameters\":{\"type\":\"object\",\"properties\":{}}}}]",
                Json(tools));
        }

        [Fact]
        public void BuildRequest_WithTools_SetsToolChoiceAuto()
        {
            var tools = JArray.Parse("[{\"type\":\"function\",\"function\":{\"name\":\"a\"}}]");

            var request = ChatCompletionsWire.BuildRequest("m1", new JArray(), tools, null, "alice", true);

            Assert.Equal("auto", (string)request["tool_choice"]);
            Assert.Equal("a", (string)request["tools"][0]["function"]["name"]);
        }

        [Fact]
        public void BuildRequest_NoTools_OmitsToolsAndToolChoice()
        {
            Assert.Null(ChatCompletionsWire.BuildRequest("m1", new JArray(), null, null, "alice", true)["tools"]);
            var emptyTools = ChatCompletionsWire.BuildRequest("m1", new JArray(), new JArray(), null, "alice", true);
            Assert.Null(emptyTools["tools"]);
            Assert.Null(emptyTools["tool_choice"]);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void BuildRequest_NoEffort_OmitsReasoningEffort(string effort)
        {
            Assert.Null(ChatCompletionsWire.BuildRequest("m1", new JArray(), null, effort, "alice", true)["reasoning_effort"]);
        }

        [Fact]
        public void BuildRequest_Effort_IsSent()
        {
            Assert.Equal("low", (string)ChatCompletionsWire.BuildRequest("m1", new JArray(), null, "low", "alice", true)["reasoning_effort"]);
        }

        [Fact]
        public void BuildRequest_TemperatureAndMaxTokens_AreSentWhenGiven()
        {
            var request = ChatCompletionsWire.BuildRequest("m1", new JArray(), null, null, "alice", true, 0.2, 800);

            Assert.Equal(0.2, (double)request["temperature"]);
            Assert.Equal(800, (int)request["max_tokens"]);
        }

        [Fact]
        public void BuildRequest_NoTemperatureOrMaxTokens_OmitsBoth()
        {
            var request = ChatCompletionsWire.BuildRequest("m1", new JArray(), null, null, "alice", true);

            Assert.Null(request["temperature"]);
            Assert.Null(request["max_tokens"]);
        }

        // Review Focus 1: the wire format always uses a dot.
        [Fact]
        public void BuildRequest_OnAnArabicPc_WritesTheTemperatureWithADot()
        {
            using (TestCultures.ArabicDecimals())
            {
                var request = ChatCompletionsWire.BuildRequest("m1", new JArray(), null, null, "alice", true, 0.2, null);

                Assert.Contains("\"temperature\":0.2", Json(request));
            }
        }

        [Fact]
        public void BuildRequest_SetsModelMessagesUserAndStream()
        {
            var messages = JArray.Parse("[{\"role\":\"user\",\"content\":\"hi\"}]");

            var request = ChatCompletionsWire.BuildRequest("m1", messages, null, null, "alice", false);

            Assert.Equal("m1", (string)request["model"]);
            Assert.Equal("hi", (string)request["messages"][0]["content"]);
            Assert.Equal("alice", (string)request["user"]);
            Assert.False((bool)request["stream"]);
        }
    }
}
