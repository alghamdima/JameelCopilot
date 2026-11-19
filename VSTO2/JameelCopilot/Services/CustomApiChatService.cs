using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Diagnostics;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Text generation through the company's Custom API: a Chat Completions
    /// endpoint (the common LLM API format) at <c>{Config.ApiBaseUrl}/chat/completions</c>,
    /// with a bearer key. This is the add-in's only network destination.
    /// </summary>
    public sealed class CustomApiChatService : IDisposable
    {
        public enum ActionType
        {
            Proofread,
            Revise,
            Draft,
            Shorten,
            Lengthen,
            Formal,
            Friendly,
            Custom,
            Arabic,
            English
        }

        private const int MaxToolRounds = 16;

        private readonly HttpClient _http;
        private readonly bool _ownsHttp;
        private bool _disposed;

        public CustomApiChatService()
            : this(BuildDefaultHttpClient(), ownsHttp: true)
        {
        }

        // Test seam.
        public CustomApiChatService(HttpClient httpClient, bool ownsHttp = false)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _ownsHttp = ownsHttp;
        }

        private static HttpClient BuildDefaultHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            // A redirect must never carry mail content or the key to another host.
            var handler = new HttpClientHandler { AllowAutoRedirect = false };
            // Covers the wait for the first byte; the stream itself ends on Stop.
            return new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(2) };
        }

        /// <summary>
        /// One chat turn: sends the conversation, runs the tools the model
        /// asks for (up to 16 rounds) and returns its final answer. Failures
        /// come back in the result with a message fit to show the user.
        /// </summary>
        public async Task<TurnResult> RunTurnAsync(
            ConversationContext context,
            string userMessage,
            IToolHost toolHost,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (toolHost == null) throw new ArgumentNullException(nameof(toolHost));
            if (sink == null) sink = new ChatEventSink();

            context.History.Add(new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "user"),
                new JProperty("content", userMessage ?? "")));

            var appended = new List<JObject>();
            var result = new TurnResult { AppendedItems = appended };
            if (!Config.IsConfigured)
            {
                return Failed(result, sink, ApiErrors.NotConfigured);
            }

            for (var round = 1; round <= MaxToolRounds; round++)
            {
                result.RoundsUsed = round;
                var messages = ChatCompletionsWire.ToMessages(
                    SystemPromptBuilder.Build(context.UseTools && !context.Structured, context.SystemInstructions, context.Context, DateTimeOffset.Now, context.Memory, context.Structured),
                    context.History);
                if (context.UseTools && !context.Structured && string.IsNullOrWhiteSpace(context.Memory) && messages.Count > 2)
                {
                    var latestLang = DetectPromptLanguage(userMessage);
                    if (latestLang != null && (userMessage ?? "").IndexOf("Translate ", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        var sysToken = messages[0] as JObject;
                        if (sysToken != null && (string)sysToken["role"] == "system")
                        {
                            sysToken["content"] = ((string)sysToken["content"] ?? "")
                                + "\n\nIMPORTANT FOR THIS TURN: The user's latest message is written in " + latestLang
                                + ". Reply strictly in " + latestLang
                                + " (even if the email body, subject, tool outputs, or earlier chat turns are in another language).";
                        }
                    }
                }
                if (context.UseTools && !context.Structured && round >= 6 && messages.Count > 2)
                {
                    var sysWrap = messages[0] as JObject;
                    if (sysWrap != null && (string)sysWrap["role"] == "system")
                    {
                        sysWrap["content"] = ((string)sysWrap["content"] ?? "")
                            + "\n\nCRITICAL TOOL BUDGET NOTICE: You have already executed multiple search/read rounds. Do NOT run further broad mailbox searches; synthesize and write your complete final answer for the user now based on the evidence already gathered (call outlook_read_messages at most once if you have unread IDs, otherwise reply immediately).";
                    }
                }
                var body = ChatCompletionsWire.BuildRequest(
                    Config.Model,
                    messages,
                    context.UseTools
                        ? ChatCompletionsWire.ToTools(ToolCatalogSchema.BuildResponsesToolsArray(context.IncludeWriteTools))
                        : null,
                    Config.ReasoningEffort,
                    Environment.UserName,
                    stream: true,
                    temperature: context.Temperature,
                    maxTokens: Config.MaxOutputTokens);

                Action<string, string> onToolCallStarted = null;
                if (context.UseTools) onToolCallStarted = (callId, name) => sink.OnToolCallStart(callId, name, "");
                var answer = await SendAsync(body, sink.OnTokenDelta, onToolCallStarted, cancellationToken)
                    .ConfigureAwait(false);

                if (answer.Error != null)
                {
                    return Failed(result, sink, answer.Error);
                }

                // On Stop, what already arrived stays in the conversation.
                if (answer.Text.Length > 0)
                {
                    var assistant = new JObject(
                        new JProperty("type", "message"),
                        new JProperty("role", "assistant"),
                        new JProperty("content", answer.Text));
                    context.History.Add(assistant);
                    appended.Add(assistant);
                    sink.OnAssistantMessageComplete(answer.Text);
                    result.FinalAssistantText = answer.Text;
                }

                if (answer.Cancelled)
                {
                    result.StopReason = StopReason.Cancelled;
                    return result;
                }

                // A tool that was not offered is not run, whatever the answer asks for.
                var calls = context.UseTools ? answer.ToolCalls : new JObject[0];
                if (calls.Count == 0)
                {
                    if (answer.Text.Length == 0)
                    {
                        return Failed(result, sink, ApiErrors.EmptyAnswer);
                    }
                    sink.OnRoundBoundary();
                    result.StopReason = StopReason.Completed;
                    return result;
                }

                JObject[] outputs;
                try
                {
                    outputs = await Task.WhenAll(
                        calls.Select(call => DispatchAsync(toolHost, sink, call, cancellationToken)))
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    result.StopReason = StopReason.Cancelled;
                    return result;
                }

                // The round as the wire format reads it: its calls, then their results.
                context.History.AddRange(calls);
                context.History.AddRange(outputs);
                appended.AddRange(calls);
                appended.AddRange(outputs);

                if (cancellationToken.IsCancellationRequested)
                {
                    result.StopReason = StopReason.Cancelled;
                    return result;
                }
                sink.OnRoundBoundary();
            }

            result.StopReason = StopReason.MaxRoundsReached;
            return result;
        }

        /// <summary>
        /// Sends one short request with the configured address, key and model.
        /// Returns null when the service answers, else the message to show.
        /// </summary>
        public async Task<string> TestConnectionAsync(CancellationToken cancellationToken)
        {
            if (!Config.IsConfigured) return ApiErrors.NotConfigured;
            var messages = new JArray(new JObject(new JProperty("role", "user"), new JProperty("content", "ping")));
            // No max_tokens: some models reject it, and others answer nothing with a cap of one.
            var body = ChatCompletionsWire.BuildRequest(Config.Model, messages, null, null, Environment.UserName, stream: false);
            var answer = await SendAsync(body, null, null, cancellationToken).ConfigureAwait(false);
            return answer.Cancelled ? "Cancelled." : answer.Error;
        }

        private static TurnResult Failed(TurnResult result, ChatEventSink sink, string message)
        {
            sink.OnError(message);
            result.StopReason = StopReason.Error;
            result.ErrorMessage = message;
            return result;
        }

        internal static string DetectPromptLanguage(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            int arabic = 0, latin = 0;
            foreach (var c in text)
            {
                if (c >= '\u0600' && c <= '\u06FF') arabic++;
                else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) latin++;
            }
            if (arabic + latin == 0) return null;
            return arabic > 0 && arabic >= latin / 4 ? "Arabic" : "English";
        }

        // One request and its answer. Never throws for a failed or cancelled
        // request: the answer says what happened.
        private async Task<Answer> SendAsync(
            JObject body, Action<string> onText, Action<string, string> onToolCallStarted, CancellationToken cancellationToken)
        {
            var reader = new ChatStreamReader(onText, onToolCallStarted);
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, Config.ApiBaseUrl.AbsoluteUri.TrimEnd('/') + "/chat/completions"))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Config.ApiKey);
                    request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");

                    using (var response = await _http.SendAsync(
                        request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            var code = (int)response.StatusCode;
                            if (code == 500 || code == 502 || code == 503 || code == 504)
                            {
                                TraceLog.Write("Transient HTTP " + code + "; retrying once", "CustomApi");
                                try { await Task.Delay(350, cancellationToken).ConfigureAwait(false); } catch { }
                                using (var retryReq = new HttpRequestMessage(HttpMethod.Post, Config.ApiBaseUrl.AbsoluteUri.TrimEnd('/') + "/chat/completions"))
                                {
                                    retryReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Config.ApiKey);
                                    retryReq.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
                                    using (var retryResp = await _http.SendAsync(retryReq, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                                    {
                                        if (retryResp.IsSuccessStatusCode)
                                        {
                                            using (cancellationToken.Register(() => retryResp.Dispose()))
                                            {
                                                var retryMedia = retryResp.Content.Headers.ContentType?.MediaType;
                                                if (string.Equals(retryMedia, "application/json", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    FeedWholeAnswer(reader, await retryResp.Content.ReadAsStringAsync().ConfigureAwait(false));
                                                }
                                                else
                                                {
                                                    using (var stream = await retryResp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                                                    using (var lines = new StreamReader(stream, Encoding.UTF8))
                                                    {
                                                        string line;
                                                        while (!reader.Done && reader.Error == null
                                                            && (line = await lines.ReadLineAsync().ConfigureAwait(false)) != null)
                                                        {
                                                            cancellationToken.ThrowIfCancellationRequested();
                                                            reader.Feed(line);
                                                        }
                                                    }
                                                }
                                            }
                                            if (reader.Error != null) return Answer.Fail(ApiErrors.ErrorInAnswer);
                                            return new Answer(reader.Text, reader.ToolCalls, null, cancelled: false);
                                        }
                                    }
                                }
                            }
                            // The body may repeat mail content or internals: it is neither shown nor logged.
                            TraceLog.Write("Request failed: HTTP " + code, "CustomApi");
                            return Answer.Fail(ApiErrors.ForStatus(response.StatusCode, RetryAfterOf(response)));
                        }

                        // On .NET Framework a pending read on the response stream ignores its token;
                        // disposing the response is what ends a stream that stalls after the headers.
                        using (cancellationToken.Register(() => response.Dispose()))
                        {
                            var mediaType = response.Content.Headers.ContentType?.MediaType;
                            if (string.Equals(mediaType, "application/json", StringComparison.OrdinalIgnoreCase))
                            {
                                // A gateway that answers in one body even when asked to stream.
                                FeedWholeAnswer(reader, await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                            }
                            else
                            {
                                using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                                using (var lines = new StreamReader(stream, Encoding.UTF8))
                                {
                                    string line;
                                    while (!reader.Done && reader.Error == null
                                        && (line = await lines.ReadLineAsync().ConfigureAwait(false)) != null)
                                    {
                                        cancellationToken.ThrowIfCancellationRequested();
                                        reader.Feed(line);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) when (cancellationToken.IsCancellationRequested && IsTransportOrCancel(ex))
            {
                return new Answer(reader.Text, new JObject[0], null, cancelled: true);
            }
            catch (Exception ex) when (IsTransportOrCancel(ex))
            {
                // Name resolution, TLS, connection reset or timeout.
                TraceLog.Write("Request failed: " + ex.GetType().Name, "CustomApi");
                return Answer.Fail(ApiErrors.ForException(ex));
            }

            if (reader.Error != null)
            {
                TraceLog.Write("The service sent an error object in its answer", "CustomApi");
                return Answer.Fail(ApiErrors.ErrorInAnswer);
            }
            return new Answer(reader.Text, reader.ToolCalls, null, cancelled: false);
        }

        private static bool IsTransportOrCancel(Exception ex)
        {
            return ex is OperationCanceledException || ex is HttpRequestException || ex is IOException
                || ex is WebException || ex is ObjectDisposedException;
        }

        // A complete (non-streamed) answer, given to the stream reader as one chunk.
        private static void FeedWholeAnswer(ChatStreamReader reader, string json)
        {
            JObject whole;
            try
            {
                whole = JToken.Parse(json ?? "") as JObject;
            }
            catch (JsonException)
            {
                return;
            }
            if (whole == null) return;

            var choice = (whole["choices"] as JArray)?.FirstOrDefault() as JObject;
            var message = choice?["message"] as JObject;
            if (message == null)
            {
                reader.Feed("data: " + whole.ToString(Formatting.None)); // an error object, or nothing usable
                return;
            }

            // Whole answers number their calls by position, not by "index".
            var calls = message["tool_calls"] as JArray;
            if (calls != null)
            {
                for (var i = 0; i < calls.Count; i++)
                {
                    if (calls[i] is JObject call) call["index"] = i;
                }
            }
            var chunk = new JObject(new JProperty("choices", new JArray(new JObject(
                new JProperty("delta", message), new JProperty("finish_reason", choice["finish_reason"])))));
            reader.Feed("data: " + chunk.ToString(Formatting.None));
        }

        private static TimeSpan? RetryAfterOf(HttpResponseMessage response)
        {
            var retry = response.Headers.RetryAfter;
            if (retry == null) return null;
            if (retry.Delta.HasValue) return retry.Delta;
            if (retry.Date.HasValue) return retry.Date.Value - DateTimeOffset.UtcNow;
            return null;
        }

        // Runs one tool call and returns its function_call_output item. A tool
        // that fails reports the failure to the model; the turn goes on.
        private static async Task<JObject> DispatchAsync(
            IToolHost toolHost, ChatEventSink sink, JObject call, CancellationToken cancellationToken)
        {
            var name = (string)call["name"] ?? "";
            var args = (string)call["arguments"] ?? "{}";
            var callId = (string)call["call_id"] ?? "";
            TraceLog.Write("Dispatch " + name + " call_id=" + callId + " args=" + DescribeArgs(args), "CustomApi");

            string outputJson;
            var ok = true;
            try
            {
                outputJson = await toolHost.DispatchAsync(name, args, cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrEmpty(outputJson))
                {
                    outputJson = "{}";
                }
                else if (outputJson.IndexOf("\"error\"", StringComparison.Ordinal) >= 0)
                {
                    // Tools answer {"error":{"code":...,"message":...}} on failure.
                    ok = false;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                outputJson = new JObject(new JProperty("error", new JObject(
                    new JProperty("code", ex.GetType().Name),
                    new JProperty("message", ex.Message ?? "")))).ToString(Formatting.None);
                ok = false;
            }

            sink.OnToolCallResult(callId, ok, Summarize(outputJson), outputJson);
            return new JObject(
                new JProperty("type", "function_call_output"),
                new JProperty("call_id", callId),
                new JProperty("output", outputJson));
        }

        /// <summary>Tool arguments as they may appear in the log: their size, never their content.</summary>
        internal static string DescribeArgs(string args)
        {
            return (args ?? "").Length + " chars";
        }

        private static string Summarize(string outputJson)
        {
            const int max = 120;
            var s = (outputJson ?? "").Replace('\n', ' ').Replace('\r', ' ');
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        private sealed class Answer
        {
            public Answer(string text, IReadOnlyList<JObject> toolCalls, string error, bool cancelled)
            {
                Text = text ?? "";
                ToolCalls = toolCalls ?? new JObject[0];
                Error = error;
                Cancelled = cancelled;
            }

            public string Text { get; }
            public IReadOnlyList<JObject> ToolCalls { get; }
            public string Error { get; }
            public bool Cancelled { get; }

            public static Answer Fail(string message)
            {
                return new Answer("", null, message, cancelled: false);
            }
        }

        // -------------------------------------------------------------------
        // Prompts for the Actions tab
        // -------------------------------------------------------------------

        public static string GetSystemPrompt(ActionType action)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return "You are a professional editor. Review the email for grammar, spelling, punctuation, and clarity issues. Return the corrected email text only. Do not add any explanations.";
                case ActionType.Revise:
                    return "You are a professional writing assistant. Improve the email clarity, flow, and impact. Return only the revised email text without any explanations.";
                case ActionType.Draft:
                    return "You are a professional email writer. Write a clear, professional email based on the instructions. If replying to an email thread, write only your reply - do not include the previous messages. Return only the email text you are composing.";
                case ActionType.Shorten:
                    return "You are a professional editor. Condense this email to be more concise while keeping essential information. Return only the shortened email text.";
                case ActionType.Lengthen:
                    return "You are a professional writer. Expand this email with more detail while maintaining professionalism. Return only the expanded email text.";
                case ActionType.Formal:
                    return "You are a professional editor. Rewrite this email in a more formal tone suitable for business. Return only the rewritten email text.";
                case ActionType.Friendly:
                    return "You are a professional editor. Rewrite this email in a warmer, friendlier tone while remaining professional. Return only the rewritten email text.";
                case ActionType.Arabic:
                    return "You are a professional translator. Translate this email into Arabic: natural, professional Arabic, not word for word. Keep names, email addresses, numbers and dates as they are. Do not include or translate any email signature, sign-off, or classification footer. Return only the translated email body.";
                case ActionType.English:
                    return "You are a professional translator. Translate this email into English: natural, professional English, not word for word. Keep names, email addresses, numbers and dates as they are. Do not include or translate any email signature, sign-off, or classification footer. Return only the translated email body.";
                case ActionType.Custom:
                default:
                    return "You are a professional email writing assistant. Help the user with their email based on their instructions. Return only the result.";
            }
        }

        /// <summary>
        /// The context for one Actions-tab rewrite. No tools: the window's
        /// quoted thread travels with the request and is not the user's own,
        /// and a rewrite needs nothing from the mailbox. The caller adds what
        /// the window holds (<see cref="ComposePrompt"/>).
        /// </summary>
        public static ConversationContext NewActionContext(ActionType action)
        {
            return new ConversationContext { SystemInstructions = GetSystemPrompt(action), UseTools = false };
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_ownsHttp)
            {
                _http.Dispose();
            }
        }
    }
}
