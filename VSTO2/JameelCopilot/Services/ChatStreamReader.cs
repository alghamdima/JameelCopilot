using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Assembles one streamed Chat Completions answer from its lines: the
    /// text, and the tool calls whose arguments arrive in fragments.
    /// </summary>
    internal sealed class ChatStreamReader
    {
        private readonly Action<string> _onText;
        private readonly Action<string, string> _onToolCallStarted;
        private readonly StringBuilder _text = new StringBuilder();
        private readonly List<PendingCall> _calls = new List<PendingCall>();                       // in arrival order
        private readonly Dictionary<int, PendingCall> _open = new Dictionary<int, PendingCall>();   // the call each index is filling

        /// <param name="onText">Called with each text fragment as it arrives.</param>
        /// <param name="onToolCallStarted">Called once per tool call, with its id and name.</param>
        public ChatStreamReader(Action<string> onText, Action<string, string> onToolCallStarted)
        {
            _onText = onText;
            _onToolCallStarted = onToolCallStarted;
        }

        public string Text => _text.ToString();

        /// <summary>The calls so far, in the history item shape (<c>function_call</c>).</summary>
        public IReadOnlyList<JObject> ToolCalls => _calls.Select(c => c.ToItem()).ToList();

        /// <summary>The message of an error object sent inside the stream, else null.</summary>
        public string Error { get; private set; }

        public bool Done { get; private set; }

        /// <summary>The service stopped the text at the answer-length limit; the text ends with a notice.</summary>
        public bool CutAtLimit { get; private set; }

        /// <summary>Takes one line of the response body; lines that aren't <c>data:</c> are ignored.</summary>
        public void Feed(string sseLine)
        {
            if (sseLine == null || !sseLine.StartsWith("data:", StringComparison.Ordinal)) return;
            var payload = sseLine.Substring(5).Trim();
            if (payload == "[DONE]")
            {
                Done = true;
                return;
            }

            JObject chunk;
            try
            {
                chunk = JToken.Parse(payload) as JObject;
            }
            catch (JsonException)
            {
                return;
            }
            if (chunk == null) return;

            var error = chunk["error"];
            if (error != null && error.Type != JTokenType.Null)
            {
                Error = error.Type == JTokenType.String
                    ? (string)error
                    : (string)(error as JObject)?["message"] ?? error.ToString(Formatting.None);
                return;
            }

            // A usage-only chunk has no choices; a first chunk may carry only the role.
            var choice = (chunk["choices"] as JArray)?.FirstOrDefault() as JObject;
            var delta = choice?["delta"] as JObject;
            if (delta != null) Take(delta);

            // Stopped at the answer-length limit: say so, or the text passes for a
            // finished answer. With no text there is nothing to mark.
            var finish = choice?["finish_reason"];
            if (!CutAtLimit && finish != null && finish.Type == JTokenType.String && (string)finish == "length"
                && _text.Length > 0 && _calls.Count == 0)
            {
                CutAtLimit = true;
                var notice = "\n\n" + ApiErrors.CutAtLimit;
                _text.Append(notice);
                _onText?.Invoke(notice);
            }
        }

        private void Take(JObject delta)
        {
            var content = delta["content"];
            if (content != null && content.Type == JTokenType.String)
            {
                var text = (string)content;
                if (text.Length > 0)
                {
                    _text.Append(text);
                    _onText?.Invoke(text);
                }
            }

            var fragments = delta["tool_calls"] as JArray;
            if (fragments == null) return;
            foreach (var fragment in fragments.OfType<JObject>())
            {
                Apply(fragment);
            }
        }

        // The first fragment of a call carries its id and name; the rest carry
        // pieces of the arguments, matched by index (absent = 0). A different
        // id on the same index is the next call: some gateways send each
        // parallel call whole and number them all 0.
        private void Apply(JObject fragment)
        {
            var indexToken = fragment["index"];
            var index = indexToken != null && indexToken.Type == JTokenType.Integer ? (int)indexToken : 0;
            var idToken = fragment["id"];
            var id = idToken != null && idToken.Type == JTokenType.String ? (string)idToken : "";

            if (!_open.TryGetValue(index, out var call)
                || (id.Length > 0 && call.Id != null && !call.IdIsOurs && call.Id != id))
            {
                call = new PendingCall();
                _calls.Add(call);
                _open[index] = call;
            }
            if (call.Id == null && id.Length > 0)
            {
                call.Id = id;
            }

            var function = fragment["function"] as JObject;
            if (function != null)
            {
                var name = function["name"];
                if (string.IsNullOrEmpty(call.Name) && name != null && name.Type == JTokenType.String)
                {
                    call.Name = (string)name;
                }
                var arguments = function["arguments"];
                if (arguments != null && arguments.Type == JTokenType.String)
                {
                    call.Arguments.Append((string)arguments);
                }
            }

            if (!call.Started && !string.IsNullOrEmpty(call.Name))
            {
                call.Started = true;
                if (call.Id == null)
                {
                    // The result is matched to its call by id, here and in the pane.
                    call.Id = "call_" + Guid.NewGuid().ToString("N");
                    call.IdIsOurs = true;
                }
                _onToolCallStarted?.Invoke(call.Id, call.Name);
            }
        }

        private sealed class PendingCall
        {
            public string Id;
            public bool IdIsOurs;
            public string Name;
            public bool Started;
            public readonly StringBuilder Arguments = new StringBuilder();

            public JObject ToItem()
            {
                return new JObject(
                    new JProperty("type", "function_call"),
                    new JProperty("call_id", Id ?? ""),
                    new JProperty("name", Name ?? ""),
                    new JProperty("arguments", Arguments.Length == 0 ? "{}" : Arguments.ToString()));
            }
        }
    }
}
