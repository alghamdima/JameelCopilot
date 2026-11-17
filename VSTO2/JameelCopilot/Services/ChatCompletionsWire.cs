using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services
{
    /// <summary>
    /// The Chat Completions request shape the Custom API speaks (the common
    /// LLM API format), built from the conversation history this add-in keeps.
    /// </summary>
    internal static class ChatCompletionsWire
    {
        /// <summary>
        /// History items (<c>message</c>, <c>function_call</c>,
        /// <c>function_call_output</c>) as Chat Completions messages. A
        /// round is a run of calls followed by their results: its calls join
        /// the assistant message just before them, or one of their own, and
        /// the results follow as <c>tool</c> messages. Each round is its own
        /// assistant message, so two rounds never read as one parallel step.
        /// </summary>
        public static JArray ToMessages(string systemInstructions, IEnumerable<JObject> history)
        {
            var items = (history ?? Enumerable.Empty<JObject>()).Where(i => i != null).ToList();
            var messages = new JArray();
            if (!string.IsNullOrEmpty(systemInstructions))
            {
                messages.Add(Message("system", systemInstructions));
            }

            int lastUserIndex = items.FindLastIndex(i =>
                (string)i["type"] == "message" && ((string)i["role"] ?? "user") == "user");

            JObject assistant = null;                 // the text message the next round's calls may join
            var at = 0;
            while (at < items.Count)
            {
                var type = (string)items[at]["type"];
                if (type == "message")
                {
                    var role = (string)items[at]["role"] ?? "user";
                    var message = Message(role, TextOf(items[at]["content"]));
                    messages.Add(message);
                    assistant = role == "assistant" ? message : null;
                    at++;
                }
                else if (type == "function_call")
                {
                    var fromOlderTurn = lastUserIndex >= 0 && at < lastUserIndex;
                    var calls = TakeRun(items, ref at, "function_call");
                    var outputs = TakeRun(items, ref at, "function_call_output");
                    AddRound(messages, assistant, calls, outputs, fromOlderTurn);
                    assistant = null;
                }
                else
                {
                    at++;                             // a result with no call before it
                }
            }
            EnforcePayloadBudget(messages);
            return messages;
        }

        private const int MaxPayloadChars = 90000;

        private static void EnforcePayloadBudget(JArray messages)
        {
            int total = messages.Sum(m => ((string)m["content"] ?? "").Length);
            if (total <= MaxPayloadChars) return;

            // Trim oldest tool messages (except the very last 2 messages) first so recent tool results stay intact.
            for (int i = 1; i < messages.Count - 2 && total > MaxPayloadChars; i++)
            {
                var msg = messages[i] as JObject;
                if (msg == null || (string)msg["role"] != "tool") continue;
                var content = (string)msg["content"] ?? "";
                if (content.Length > 800)
                {
                    var shortened = content.Substring(0, 800) + "...[trimmed to fit context window]";
                    total -= (content.Length - shortened.Length);
                    msg["content"] = shortened;
                }
            }
        }

        private static List<JObject> TakeRun(List<JObject> items, ref int at, string type)
        {
            var run = new List<JObject>();
            while (at < items.Count && (string)items[at]["type"] == type)
            {
                run.Add(items[at++]);
            }
            return run;
        }

        // A result goes to the first call of its round that has the same id
        // and no result yet, so a repeated id still pairs in order. An API
        // rejects a call without a result (a turn stopped mid-run): such a
        // call is not sent back.
        private const int MaxOlderToolOutputChars = 600;

        private static void AddRound(JArray messages, JObject assistant, List<JObject> calls, List<JObject> outputs, bool fromOlderTurn = false)
        {
            var resultOf = new string[calls.Count];
            foreach (var output in outputs)
            {
                var id = (string)output["call_id"] ?? "";
                for (var k = 0; k < calls.Count; k++)
                {
                    if (resultOf[k] != null || ((string)calls[k]["call_id"] ?? "") != id) continue;
                    resultOf[k] = (string)output["output"] ?? "";
                    break;
                }
            }
            if (resultOf.All(r => r == null)) return;

            if (assistant == null)
            {
                assistant = new JObject(new JProperty("role", "assistant"), new JProperty("content", JValue.CreateNull()));
                messages.Add(assistant);
            }
            var toolCalls = new JArray();
            assistant["tool_calls"] = toolCalls;
            var results = new List<JObject>();
            for (var k = 0; k < calls.Count; k++)
            {
                if (resultOf[k] == null) continue;
                var id = (string)calls[k]["call_id"] ?? "";
                var args = (string)calls[k]["arguments"] ?? "{}";
                if (fromOlderTurn && args.Length > MaxOlderToolOutputChars)
                {
                    args = "{\"_trimmed\":true}";
                }
                var content = resultOf[k];
                if (fromOlderTurn && content.Length > MaxOlderToolOutputChars)
                {
                    content = content.Substring(0, MaxOlderToolOutputChars) + "...[older tool output trimmed]";
                }
                toolCalls.Add(new JObject(
                    new JProperty("id", id),
                    new JProperty("type", "function"),
                    new JProperty("function", new JObject(
                        new JProperty("name", (string)calls[k]["name"] ?? ""),
                        new JProperty("arguments", args)))));
                results.Add(new JObject(
                    new JProperty("role", "tool"),
                    new JProperty("tool_call_id", id),
                    new JProperty("content", content)));
            }
            messages.Add(results);
        }

        /// <summary>
        /// The tool catalog (<c>{type, name, description, parameters}</c>
        /// entries) in the Chat Completions form, each wrapped in <c>function</c>.
        /// </summary>
        public static JArray ToTools(JArray catalog)
        {
            var tools = new JArray();
            foreach (var entry in (catalog ?? new JArray()).OfType<JObject>())
            {
                var function = new JObject(entry.Properties().Where(p => p.Name != "type").Select(p => p.DeepClone()));
                tools.Add(new JObject(new JProperty("type", "function"), new JProperty("function", function)));
            }
            return tools;
        }

        /// <summary>
        /// One request body. No tools means no <c>tools</c> and no
        /// <c>tool_choice</c>; no effort, temperature or token limit means
        /// that field is left out.
        /// </summary>
        public static JObject BuildRequest(
            string model, JArray messages, JArray tools, string reasoningEffort, string user, bool stream,
            double? temperature = null, int? maxTokens = null)
        {
            var request = new JObject(
                new JProperty("model", model ?? ""),
                new JProperty("messages", messages ?? new JArray()));
            if (tools != null && tools.Count > 0)
            {
                request["tools"] = tools;
                request["tool_choice"] = "auto";
            }
            if (!string.IsNullOrEmpty(reasoningEffort))
            {
                request["reasoning_effort"] = reasoningEffort;
            }
            if (temperature.HasValue)
            {
                request["temperature"] = temperature.Value;
            }
            if (maxTokens.HasValue)
            {
                request["max_tokens"] = maxTokens.Value;
            }
            request["stream"] = stream;
            if (!string.IsNullOrEmpty(user))
            {
                request["user"] = user;
            }
            return request;
        }

        private static JObject Message(string role, string content)
        {
            return new JObject(new JProperty("role", role), new JProperty("content", content ?? ""));
        }

        // Content is a string, or an array of parts that each carry "text".
        private static string TextOf(JToken content)
        {
            if (content == null || content.Type == JTokenType.Null) return "";
            if (content.Type == JTokenType.String) return (string)content;
            var parts = content as JArray;
            if (parts == null) return content.ToString();
            var text = new StringBuilder();
            foreach (var part in parts.OfType<JObject>())
            {
                text.Append((string)part["text"] ?? "");
            }
            return text.ToString();
        }
    }
}
