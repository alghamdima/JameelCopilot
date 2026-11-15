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
                    var calls = TakeRun(items, ref at, "function_call");
                    var outputs = TakeRun(items, ref at, "function_call_output");
                    AddRound(messages, assistant, calls, outputs);
                    assistant = null;
                }
                else
                {
                    at++;                             // a result with no call before it
                }
            }
            return messages;
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
        private static void AddRound(JArray messages, JObject assistant, List<JObject> calls, List<JObject> outputs)
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
                toolCalls.Add(new JObject(
                    new JProperty("id", id),
                    new JProperty("type", "function"),
                    new JProperty("function", new JObject(
                        new JProperty("name", (string)calls[k]["name"] ?? ""),
                        new JProperty("arguments", (string)calls[k]["arguments"] ?? "{}")))));
                results.Add(new JObject(
                    new JProperty("role", "tool"),
                    new JProperty("tool_call_id", id),
                    new JProperty("content", resultOf[k])));
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
