using System;
using System.Collections.Generic;
using JameelCopilot.Services;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    public class ChatStreamReaderTests
    {
        private readonly List<string> _texts = new List<string>();
        private readonly List<string> _started = new List<string>();

        private ChatStreamReader NewReader() =>
            new ChatStreamReader(t => _texts.Add(t), (id, name) => _started.Add(id + ":" + name));

        private static string Content(string text) =>
            "data: {\"choices\":[{\"index\":0,\"delta\":{\"content\":" + Newtonsoft.Json.JsonConvert.ToString(text) + "}}]}";

        private static string ToolFragment(string fragmentJson) =>
            "data: {\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[" + fragmentJson + "]}}]}";

        [Fact]
        public void Text_AccumulatesContentDeltas()
        {
            var reader = NewReader();

            reader.Feed(Content("Hel"));
            reader.Feed("data:{\"choices\":[{\"delta\":{\"content\":\"lo\"}}]}"); // no space after "data:"

            Assert.Equal("Hello", reader.Text);
            Assert.Equal(new[] { "Hel", "lo" }, _texts);
            Assert.Empty(reader.ToolCalls);
        }

        [Fact]
        public void ToolCall_ArgumentsAssembleAcrossFragments()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"type\":\"function\",\"function\":{\"name\":\"outlook_search_messages\",\"arguments\":\"\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"function\":{\"arguments\":\"{\\\"from\\\":\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"function\":{\"arguments\":\"\\\"Alice\\\"}\"}}"));

            var call = Assert.Single(reader.ToolCalls);
            Assert.Equal("function_call", (string)call["type"]);
            Assert.Equal("c1", (string)call["call_id"]);
            Assert.Equal("outlook_search_messages", (string)call["name"]);
            Assert.Equal("{\"from\":\"Alice\"}", (string)call["arguments"]);
            Assert.Equal(new[] { "c1:outlook_search_messages" }, _started);
        }

        [Fact]
        public void TwoToolCalls_ByIndex_StaySeparate()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{\"}}"));
            reader.Feed(ToolFragment("{\"index\":1,\"id\":\"c2\",\"function\":{\"name\":\"b\",\"arguments\":\"{\"}}"));
            reader.Feed(ToolFragment("{\"index\":1,\"function\":{\"arguments\":\"\\\"y\\\":2}\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"function\":{\"arguments\":\"\\\"x\\\":1}\"}}"));

            Assert.Equal(2, reader.ToolCalls.Count);
            Assert.Equal("c1", (string)reader.ToolCalls[0]["call_id"]);
            Assert.Equal("{\"x\":1}", (string)reader.ToolCalls[0]["arguments"]);
            Assert.Equal("c2", (string)reader.ToolCalls[1]["call_id"]);
            Assert.Equal("{\"y\":2}", (string)reader.ToolCalls[1]["arguments"]);
            Assert.Equal(new[] { "c1:a", "c2:b" }, _started);
        }

        // Some gateways send each parallel call whole, all numbered 0: a new id is a new call.
        [Fact]
        public void TwoWholeCalls_SharingAnIndex_StaySeparate()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{}\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c2\",\"function\":{\"name\":\"b\",\"arguments\":\"{\\\"from\\\":\\\"Bob\\\"}\"}}"));

            Assert.Equal(2, reader.ToolCalls.Count);
            Assert.Equal(("c1", "a", "{}"),
                ((string)reader.ToolCalls[0]["call_id"], (string)reader.ToolCalls[0]["name"], (string)reader.ToolCalls[0]["arguments"]));
            Assert.Equal(("c2", "b", "{\"from\":\"Bob\"}"),
                ((string)reader.ToolCalls[1]["call_id"], (string)reader.ToolCalls[1]["name"], (string)reader.ToolCalls[1]["arguments"]));
            Assert.Equal(new[] { "c1:a", "c2:b" }, _started);
        }

        // Other gateways repeat the call's id on every fragment: still one call.
        [Fact]
        public void SameIdOnEveryFragment_IsOneCall()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"arguments\":\"}\"}}"));

            Assert.Equal("{}", (string)Assert.Single(reader.ToolCalls)["arguments"]);
        }

        // A call's result is matched to it by id, so a call that arrives without one gets one.
        [Fact]
        public void CallWithoutId_GetsItsOwn()
        {
            var first = NewReader();
            first.Feed(ToolFragment("{\"index\":0,\"function\":{\"name\":\"a\",\"arguments\":\"{}\"}}"));
            first.Feed(ToolFragment("{\"index\":1,\"function\":{\"name\":\"b\",\"arguments\":\"{}\"}}"));
            var second = NewReader();
            second.Feed(ToolFragment("{\"index\":0,\"id\":\"\",\"function\":{\"name\":\"a\",\"arguments\":\"{}\"}}"));

            var ids = new[]
            {
                (string)first.ToolCalls[0]["call_id"], (string)first.ToolCalls[1]["call_id"], (string)second.ToolCalls[0]["call_id"],
            };
            Assert.All(ids, id => Assert.False(string.IsNullOrEmpty(id)));
            Assert.Equal(3, new HashSet<string>(ids).Count);
            // The pane is told the same id the result will carry.
            Assert.Equal(new[] { ids[0] + ":a", ids[1] + ":b", ids[2] + ":a" }, _started);
        }

        // An id that arrives after the name does not start a second call.
        [Fact]
        public void IdAfterTheName_IsStillOneCall()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"function\":{\"name\":\"a\",\"arguments\":\"{\"}}"));
            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"late\",\"function\":{\"arguments\":\"}\"}}"));

            Assert.Equal("{}", (string)Assert.Single(reader.ToolCalls)["arguments"]);
            Assert.Single(_started);
        }

        // Some gateways leave "index" out when the model makes a single call.
        [Fact]
        public void ToolCall_WithoutIndex_IsIndexZero()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{\"}}"));
            reader.Feed(ToolFragment("{\"function\":{\"arguments\":\"}\"}}"));

            var call = Assert.Single(reader.ToolCalls);
            Assert.Equal("{}", (string)call["arguments"]);
        }

        // A model that sends a whole call in one chunk, with no arguments field.
        [Fact]
        public void ToolCall_WithoutArguments_GetsEmptyObject()
        {
            var reader = NewReader();

            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"outlook_list_folders\"}}"));

            Assert.Equal("{}", (string)Assert.Single(reader.ToolCalls)["arguments"]);
        }

        [Theory]
        [InlineData("")]
        [InlineData(": keep-alive")]
        [InlineData("event: message")]
        [InlineData("data: {\"choices\":[]}")]
        [InlineData("data: {\"choices\":[],\"usage\":{\"total_tokens\":12}}")]
        [InlineData("data: {\"choices\":[{\"index\":0,\"delta\":{\"role\":\"assistant\"}}]}")]
        [InlineData("data: {\"choices\":[{\"index\":0,\"delta\":{\"content\":null}}]}")]
        [InlineData("data: {\"choices\":[{\"index\":0,\"delta\":{},\"finish_reason\":\"stop\"}]}")]
        [InlineData("data: {not json")]
        [InlineData("data: [1,2]")]
        public void IgnoresLinesThatCarryNothing(string line)
        {
            var reader = NewReader();

            reader.Feed(line);

            Assert.Equal("", reader.Text);
            Assert.Empty(reader.ToolCalls);
            Assert.Null(reader.Error);
            Assert.False(reader.Done);
            Assert.Empty(_texts);
        }

        [Fact]
        public void Feed_Null_IsIgnored()
        {
            var reader = NewReader();

            reader.Feed(null);

            Assert.Equal("", reader.Text);
        }

        [Fact]
        public void Done_OnDoneMarker()
        {
            var reader = NewReader();

            reader.Feed(Content("hi"));
            reader.Feed("data: [DONE]");

            Assert.True(reader.Done);
            Assert.Equal("hi", reader.Text);
        }

        [Theory]
        [InlineData("data: {\"error\":{\"message\":\"boom\",\"type\":\"server_error\"}}")]
        [InlineData("data: {\"error\":\"boom\"}")]
        public void Error_OnErrorObject(string line)
        {
            var reader = NewReader();

            reader.Feed(Content("partial"));
            reader.Feed(line);

            Assert.Equal("boom", reader.Error);
            Assert.Equal("partial", reader.Text);
        }

        private static string Finish(string reason) =>
            "data: {\"choices\":[{\"index\":0,\"delta\":{},\"finish_reason\":\"" + reason + "\"}]}";

        // The administrator's answer-length limit stops a text mid-sentence,
        // and nothing on the screen tells it from a finished answer.
        [Fact]
        public void FinishReasonLength_EndsTheTextWithANotice_Once()
        {
            var reader = NewReader();

            reader.Feed(Content("The report shows"));
            reader.Feed(Finish("length"));
            reader.Feed(Finish("length"));

            Assert.Equal("The report shows\n\n" + ApiErrors.CutAtLimit, reader.Text);
            Assert.Equal(new[] { "The report shows", "\n\n" + ApiErrors.CutAtLimit }, _texts);
        }

        [Fact]
        public void FinishReasonLength_InTheLastTextChunk_ComesAfterItsText()
        {
            var reader = NewReader();

            reader.Feed("data: {\"choices\":[{\"delta\":{\"content\":\"The report shows\"},\"finish_reason\":\"length\"}]}");

            Assert.Equal("The report shows\n\n" + ApiErrors.CutAtLimit, reader.Text);
        }

        [Theory]
        [InlineData("stop")]
        [InlineData("tool_calls")]
        [InlineData("content_filter")]
        public void OtherFinishReasons_AddNothing(string reason)
        {
            var reader = NewReader();

            reader.Feed(Content("Done."));
            reader.Feed(Finish(reason));
            reader.Feed("data: {\"choices\":[{\"delta\":{},\"finish_reason\":null}]}");
            reader.Feed("data: {\"choices\":[{\"delta\":{},\"finish_reason\":{\"a\":1}}]}");

            Assert.Equal("Done.", reader.Text);
        }

        // With no text there is nothing to mark: an empty answer and a cut
        // tool call are reported as what they are.
        [Fact]
        public void FinishReasonLength_WithoutText_AddsNothing()
        {
            var empty = NewReader();
            empty.Feed(Finish("length"));
            Assert.Equal("", empty.Text);

            var call = NewReader();
            call.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{\"}}"));
            call.Feed(Finish("length"));
            Assert.Equal("", call.Text);
        }

        [Fact]
        public void Callbacks_AreOptional()
        {
            var reader = new ChatStreamReader(null, null);

            reader.Feed(Content("hi"));
            reader.Feed(ToolFragment("{\"index\":0,\"id\":\"c1\",\"function\":{\"name\":\"a\",\"arguments\":\"{}\"}}"));

            Assert.Equal("hi", reader.Text);
            Assert.Single(reader.ToolCalls);
        }
    }
}
