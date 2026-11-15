using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using JameelCopilot.Diagnostics;
using JameelCopilot.Services;
using JameelCopilot.Services.Export;
using JameelCopilot.Services.Prefs;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// Owns the Chat tab's WebView2 lifecycle and the JS&#x2194;C# bridge that
    /// wires user input into <see cref="CustomApiChatService.RunTurnAsync"/>.
    /// It also writes and edits the email itself (chips and the Draft
    /// button), showing each result as a draft card.
    /// Per-Inspector instance, constructed by <see cref="AITaskPane"/> after
    /// <see cref="AITaskPane.Bind"/> hands it the tool host + surface.
    /// </summary>
    public sealed class ChatController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CustomApiChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly ExportBridge _exportBridge;
        private readonly PrefsBridge _prefs;
        private readonly ChatsBridge _chats;
        private readonly Dictionary<string, MeetingProposal> _meetings = new Dictionary<string, MeetingProposal>(StringComparer.Ordinal);
        private readonly RulesBridge _rules;
        private readonly Outlook.Inspector _inspector;
        private readonly Dictionary<string, Draft> _drafts = new Dictionary<string, Draft>(StringComparer.Ordinal);

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public ChatController(
            Control hostContainer,
            CustomApiChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store,
            UserPrefsService prefs = null,
            ChatArchive chats = null,
            UserMemoryService memory = null,
            string scope = null,
            string title = null,
            Outlook.Inspector inspector = null,
            UserMemoryService prompts = null)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
            _prefs = new PrefsBridge(prefs, RunScript);
            _chats = new ChatsBridge(chats, memory, _store, RunScript, scope, title, () => _turnInFlight, null,
                prompts, () => { if (_isReady && !_isDisposed) PushContextStripFromSurface(); });
            _rules = new RulesBridge(surface, RunScript);
            _inspector = inspector;
        }

        /// <summary>
        /// Initialize the WebView2, extract WebUI resources, wire bridge.
        /// Awaited by <see cref="AITaskPane.Bind"/>. On failure, leaves a
        /// friendly fallback label on the tab.
        /// </summary>
        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "ChatController");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                TraceLog.Write("WebView2 runtime NOT installed; showing fallback", "ChatController");
                ShowFallback("Jameel Copilot needs the WebView2 runtime. Contact IT.");
                return;
            }
            TraceLog.Write("WebView2 runtime detected; constructing WebView2 control", "ChatController");

            _webView = new WebView2 { Dock = DockStyle.Fill };
            TraceLog.Write("WebView2 control constructed", "ChatController");
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);
            TraceLog.Write("WebView2 control added to host container; about to await Bootstrap.InitializeAsync", "ChatController");

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                TraceLog.Write("Bootstrap.InitializeAsync awaited OK", "ChatController");
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                TraceLog.Write("WebMessageReceived subscribed; navigating", "ChatController");
                _webView.CoreWebView2.Navigate(WebView2Bootstrap.PageUrl);
                TraceLog.Write("Navigate called", "ChatController");
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController.InitializeAsync: " + ex);
                ShowFallback("WebView2 failed to initialize: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null)
            {
                _fallbackLabel.Text = message;
                return;
            }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // Only the add-in's own page may drive the host.
                if (!WebView2Bootstrap.IsTrustedSource(e.Source)) return;
                var json = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                TraceLog.Write("WebMessageReceived: " + type, "ChatController");
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController bridge parse error: " + ex);
            }
        }

        private async Task HandleHostMessageAsync(string type, JObject payload)
        {
            try
            {
                if (_exportBridge != null && await _exportBridge.HandleAsync(
                    type, payload, _activeCts?.Token ?? CancellationToken.None).ConfigureAwait(false))
                {
                    return;
                }
                if (_prefs.Handle(type, payload))
                {
                    return;
                }

                if (_chats.Handle(type, payload) || _rules.Handle(type, payload))
                {
                    return;
                }

                switch (type)
                {
                    case "ready":
                        OnWebViewReady();
                        _ = _prefs.PushAsync();
                        _ = _chats.OpenAsync();
                        break;
                    case "send":
                        _ = StartTurnAsync((string)payload?["text"] ?? "");
                        break;
                    case "stop":
                        try { _activeCts?.Cancel(); } catch { }
                        break;
                    case "meeting_open":
                        OpenMeeting((string)payload?["call"], payload?["slot"]?.Type == JTokenType.Integer ? (int)payload["slot"] : -1);
                        break;
                    case "open_message":
                        OpenMessage((string)payload?["id"]);
                        break;
                    case "chip_action":
                        if ((string)payload?["id"] == ComposeChips.ReviewChipId) _ = RunReviewAsync();
                        else if (ComposeChips.TryParse((string)payload?["id"], out var chipAction)) _ = RunDraftAsync(chipAction, "");
                        break;
                    case "review_apply":
                        ApplyReview((string)payload?["id"], payload?["picks"] as JArray);
                        break;
                    case "draft":
                        _ = RunDraftAsync(CustomApiChatService.ActionType.Draft, (string)payload?["text"] ?? "");
                        break;
                    case "draft_action":
                        HandleDraftAction((string)payload?["id"], (string)payload?["action"]);
                        break;
                    case "insert_text":
                        InsertAnswer((string)payload?["text"]);
                        break;
                    case "copy":
                        var clip = _store.ExportForClipboard();
                        try { Clipboard.SetText(clip ?? ""); } catch { /* clipboard occasionally throws on Outlook */ }
                        break;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("HandleHostMessageAsync EXCEPTION: " + ex, "ChatController");
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "ChatController");
            _isReady = true;
            _ = RunScript("jameelcopilot.applyTheme('light');");
            _ = RunScript("jameelcopilot.setComposeMode(true);");
            PushContextStripFromSurface();
            TraceLog.Write("OnWebViewReady completed", "ChatController");
        }

        private void PushContextStripFromSurface()
        {
            if (_surface == null) { TraceLog.Write("PushContextStrip: _surface is null", "ChatController"); return; }
            try
            {
                TraceLog.Write(">> PushContextStrip calling GetCurrentComposeState", "ChatController");
                var state = _surface.GetCurrentComposeState(includeFullBody: false);
                TraceLog.Write("<< PushContextStrip GetCurrentComposeState returned", "ChatController");
                var isReply = !string.IsNullOrWhiteSpace(state?.QuotedPlaintext);
                var ctx = new JObject(
                    new JProperty("subject", state?.Subject ?? ""),
                    new JProperty("recipients", new JArray((state?.ToRecipients ?? new string[0]).Take(3))),
                    new JProperty("thread", isReply ? "Jameel Copilot writes as you and leaves your signature alone." : ""));
                _ = RunScript("jameelcopilot.setContextStrip(" + ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");
                var chips = ComposeChips.Groups(isReply);
                chips.Add(_chats.PromptsGroup());
                _ = RunScript("jameelcopilot.setQuickActions(" + chips.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStrip EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("PushContextStrip error: " + ex);
            }
        }

        private async Task StartTurnAsync(string userText)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "ChatController");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "ChatController");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("jameelcopilot.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                // RunTurnAsync mutates context.History in place (adds the user
                // message, then assistant/function-call/output items). We start
                // with a snapshot of the existing store, let the turn evolve it,
                // and then sync the diff back into the store at the end.
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = SystemPromptBuilder.Pane("compose"),
                    Context = BuildComposeContext(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    Temperature = _prefs.Temperature,
                    Memory = _chats.Memory,
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                // Sync newly-appended items (user msg + assistant + tool round-
                // trips) back into the store so the next turn starts from the
                // updated history.
                for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                {
                    _store.Append(ctx.History[i]);
                }
                _chats.Save();

                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(assistantId) +
                                ", {stopped:true});");
            }
            catch (Exception ex)
            {
                await RunScript("jameelcopilot.showError(" + JsString(ex.Message ?? "") + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("jameelcopilot.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "ChatController");
            }
        }

        // -- writing the email ---------------------------------------------

        // A draft the page shows as a card. Its text is kept here: what goes
        // into the email is what Jameel Copilot wrote, not what the page sends back.
        private sealed class Draft
        {
            public CustomApiChatService.ActionType Action;
            public string Instructions;
            public string Text;
            public ComposeSelection Selection;   // set when the draft is for the selected text only
        }

        // -- Review (coaching) -------------------------------------------------

        // Structured answers are JSON for Jameel Copilot's screen: nothing of them streams to the page.
        private sealed class QuietSink : ChatEventSink { }

        // The points of each review card, kept here: Apply uses these, never text from the page.
        private readonly Dictionary<string, List<CoachPoint>> _reviews = new Dictionary<string, List<CoachPoint>>(StringComparer.Ordinal);

        /// <summary>Reviews the user's draft and shows the points as a card the user picks from.</summary>
        private async Task RunReviewAsync()
        {
            if (_turnInFlight || !_isReady) return;
            var compose = _surface?.GetCurrentComposeState(includeFullBody: true);
            if (string.IsNullOrWhiteSpace(compose?.BodyPlaintext))
            {
                await RunScript("jameelcopilot.appendAuditRow(" + JsString("Write your email first, then press Review.") + ");");
                return;
            }

            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var id = "review_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(id) + ", " + JsString("Reviewing your draft...") + ");");
            try
            {
                var language = StructuredAnswers.ReplyLanguage(_store.Snapshot(), "the language the draft is written in");
                var ctx = StructuredAnswers.NewContext(DraftCoaching.Instructions(language),
                    ComposePrompt.Context(compose, includeDraft: true), _chats.Memory);
                var result = await _chat.RunTurnAsync(ctx, "Review my draft.", _toolHost, new QuietSink(), _activeCts.Token);
                var answer = StructuredAnswers.ParseObject(result.FinalAssistantText);
                if (result.StopReason != StopReason.Completed || answer == null)
                {
                    var why = result.StopReason == StopReason.Cancelled ? "Stopped."
                        : result.StopReason == StopReason.Error ? (result.ErrorMessage ?? "The AI service returned an error.")
                        : "Jameel Copilot could not review the draft this time. Try again.";
                    await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString(why) + ");");
                    await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {error:true});");
                    return;
                }

                var summary = DraftCoaching.Summary(answer);
                var points = DraftCoaching.Parse(answer);
                _reviews[id] = points;
                await RunScript("jameelcopilot.showReview(" + JsString(id) + ", " + JsString(summary) + ", " +
                                DraftCoaching.ToJson(points).ToString(Newtonsoft.Json.Formatting.None) + ");");
                _store.Append(HistoryMessage("user", "Review my draft."));
                _store.Append(HistoryMessage("assistant", DraftCoaching.AsText(summary, points)));
                _chats.Save();
            }
            catch (OperationCanceledException)
            {
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Stopped.") + ");");
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunReviewAsync failed: " + ex.GetType().Name, "ChatController");
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Jameel Copilot could not review the draft this time. Try again.") + ");");
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {error:true});");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("jameelcopilot.setComposerEnabled(true, false);");
            }
        }

        // "Apply selected" on a review card: a draft card with those changes only.
        private void ApplyReview(string reviewId, JArray picks)
        {
            if (reviewId == null || !_reviews.TryGetValue(reviewId, out var points)) return;
            var wanted = new HashSet<string>((picks ?? new JArray()).Where(t => t.Type == JTokenType.String).Select(t => (string)t), StringComparer.Ordinal);
            var chosen = points.Where(p => wanted.Contains(p.Id)).ToList();
            if (chosen.Count == 0) return;
            _ = RunDraftAsync(CustomApiChatService.ActionType.Custom, DraftCoaching.ApplyInstructions(chosen));
        }

        // -- the selection --------------------------------------------------

        private T OnUi<T>(Func<T> read, T failed)
        {
            try
            {
                return _hostContainer.InvokeRequired ? (T)_hostContainer.Invoke(read) : read();
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnUi failed: " + ex.GetType().Name, "ChatController");
                return failed;
            }
        }

        /// <summary>
        /// Runs one writing action (a chip, or Draft with typed instructions)
        /// and shows the result as a draft card. Nothing reaches the email
        /// until the user presses one of the card's buttons.
        /// </summary>
        private async Task RunDraftAsync(CustomApiChatService.ActionType action, string instructions, ComposeSelection selection = null)
        {
            if (_turnInFlight || !_isReady) return;

            // What the user typed, apart from their signature and the message being replied to.
            var compose = _surface?.GetCurrentComposeState(includeFullBody: true);

            // An edit with text selected in the user's own part works on that text only.
            var isEdit = action != CustomApiChatService.ActionType.Draft && action != CustomApiChatService.ActionType.Custom
                && !string.IsNullOrWhiteSpace(compose?.BodyPlaintext);
            if (isEdit && selection == null && _inspector != null)
            {
                selection = OnUi(() => ComposeDocument.ReadSelection(_inspector), null);
            }
            if (!isEdit) selection = null;
            if (selection != null) compose = ComposePrompt.WithOwnText(compose, selection.Text);
            var problem = ComposePrompt.Problem(action, compose);
            if (problem != null)
            {
                await RunScript("jameelcopilot.appendAuditRow(" + JsString(problem) + ");");
                return;
            }

            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var writesReply = ComposePrompt.WritesReply(action, compose);
            var hasOwnText = !string.IsNullOrWhiteSpace(compose?.BodyPlaintext);
            var id = "draft_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            var title = ComposeChips.DraftTitle(action, writesReply) + (selection != null ? " (selected text)" : "");
            await RunScript("jameelcopilot.beginDraft(" + JsString(id) + ", " + JsString(title) + ");");

            try
            {
                var ctx = CustomApiChatService.NewActionContext(action);
                ctx.SystemInstructions = ComposePrompt.Instructions(action, compose);
                ctx.Context = ComposePrompt.Context(compose, includeDraft: false);
                ctx.Temperature = _prefs.Temperature;
                ctx.Memory = _chats.Memory;
                var userMessage = ComposePrompt.UserMessage(action, compose, instructions ?? "");

                var result = await _chat.RunTurnAsync(ctx, userMessage, _toolHost, new WebViewSink(this, id), _activeCts.Token);
                var text = (result.FinalAssistantText ?? "").Trim();

                if (result.StopReason != StopReason.Completed || text.Length == 0)
                {
                    await RunScript(ComposeChips.FinalizeScript(id, ComposeChips.DiscardOnly(),
                        result.StopReason == StopReason.Cancelled ? "stopped" : "not finished", null, result.StopReason == StopReason.Error));
                    return;
                }

                _drafts[id] = new Draft { Action = action, Instructions = instructions, Text = text, Selection = selection };
                // The chat keeps what was asked and what Jameel Copilot wrote, so "make it shorter" has something to work on.
                _store.Append(HistoryMessage("user", ComposeChips.Asked(action, writesReply, instructions)));
                _store.Append(HistoryMessage("assistant", text));
                _chats.Save();
                await RunScript(ComposeChips.FinalizeScript(id,
                    selection != null ? ComposeChips.SelectionButtons() : ComposeChips.DraftButtons(hasOwnText), "not in the email yet",
                    writesReply ? ComposePrompt.WroteReply : null, false));
            }
            catch (OperationCanceledException)
            {
                await RunScript(ComposeChips.FinalizeScript(id, ComposeChips.DiscardOnly(), "stopped", null, false));
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunDraftAsync failed: " + ex.GetType().Name, "ChatController");
                await RunScript(ComposeChips.FinalizeScript(id, ComposeChips.DiscardOnly(), "not finished", null, true));
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("jameelcopilot.setComposerEnabled(true, false);");
            }
        }

        private static JObject HistoryMessage(string role, string text)
        {
            return new JObject(new JProperty("type", "message"), new JProperty("role", role), new JProperty("content", text));
        }

        // A button on a draft card.
        private void HandleDraftAction(string id, string what)
        {
            if (id == null || !_drafts.TryGetValue(id, out var draft)) return;
            switch (what)
            {
                case "insert":
                    MarkDraft(id, ChangeEmail(inspector => ComposeDocument.InsertAtTop(inspector, draft.Text)), "in the email");
                    break;
                case "replace":
                    MarkDraft(id, ChangeEmail(inspector => ComposeDocument.ReplaceOwnText(inspector, draft.Text)), "replaced your text");
                    break;
                case "replace_selection":
                    var replaced = ChangeEmail(inspector => ComposeDocument.ReplaceSelection(inspector, draft.Selection, draft.Text));
                    _ = RunScript(ComposeChips.MarkScript(id, replaced, "replaced the selected text",
                        "the selected text changed since; select it again and press the chip"));
                    break;
                case "regenerate":
                    _ = RunDraftAsync(draft.Action, draft.Instructions, draft.Selection);
                    break;
                case "discard":
                    _drafts.Remove(id);
                    break;
            }
        }

        private void MarkDraft(string id, bool done, string where)
        {
            _ = RunScript(ComposeChips.MarkScript(id, done, where, "could not reach the email; copy the text instead"));
        }

        // An answer in the chat that the user wants in the email as it is.
        private void InsertAnswer(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            var done = ChangeEmail(inspector => ComposeDocument.InsertAtTop(inspector, text));
            _ = RunScript("jameelcopilot.appendAuditRow(" +
                          JsString(done ? "Inserted at the top of the email." : "Could not reach the email. Copy the text instead.") + ");");
        }

        // Text goes into the email through its editor, on the UI thread, so the
        // signature and the quoted message stay exactly as they are.
        private bool ChangeEmail(Func<Outlook.Inspector, bool> change)
        {
            try
            {
                if (_inspector == null) return false;
                return _hostContainer.InvokeRequired
                    ? (bool)_hostContainer.Invoke(new Func<bool>(() => change(_inspector)))
                    : change(_inspector);
            }
            catch (Exception ex)
            {
                // A window that was closed, or an editor that is not there.
                TraceLog.Write("ChangeEmail failed: " + ex.GetType().Name, "ChatController");
                return false;
            }
        }

        // What the compose window holds right now. It goes to the model as context (data).
        private string BuildComposeContext()
        {
            try
            {
                var state = _surface?.GetCurrentComposeState(includeFullBody: false);
                return ComposePrompt.Context(state, includeDraft: true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("BuildComposeContext error: " + ex);
                return "";
            }
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                // EVERY WebView2 property/method access (including just
                // reading _webView.CoreWebView2) MUST happen on the UI thread.
                // Marshal first, THEN touch the control.
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null
                    && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
                {
                    await marshaller.RunAsync(() =>
                    {
                        if (_isDisposed) return;
                        var core = _webView?.CoreWebView2;
                        if (core == null) return;
                        _ = core.ExecuteScriptAsync(script);
                    }, CancellationToken.None).ConfigureAwait(false);
                    return;
                }
                // On UI thread: safe to touch the control directly.
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "ChatController");
                System.Diagnostics.Debug.WriteLine("ExecuteScriptAsync failed: " + ex);
            }
        }

        // JSON-encode a string for safe inlining inside an ExecuteScriptAsync
        // call. Newtonsoft handles quote / backslash / control characters.
        // A time picked on a meeting card: Outlook's meeting window, filled in, for the user to send.
        private void OpenMeeting(string callId, int slot)
        {
            var opened = false;
            try
            {
                opened = callId != null && _meetings.TryGetValue(callId, out var proposal) && _surface != null
                         && _surface.OpenMeetingRequest(proposal, slot);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OpenMeeting failed: " + ex.GetType().Name, "Jameel Copilot");
            }
            if (!opened)
            {
                _ = RunScript("jameelcopilot.appendAuditRow(" + JsString("The meeting request could not be opened.") + ");");
            }
        }

        // An email listed under an answer, clicked by the user.
        private void OpenMessage(string id)
        {
            var opened = false;
            try
            {
                opened = _surface != null && _surface.OpenMessage(id);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OpenMessage failed: " + ex.GetType().Name, "Jameel Copilot");
            }
            if (!opened)
            {
                _ = RunScript("jameelcopilot.appendAuditRow(" + JsString("That email could not be opened. It may have been moved or deleted.") + ");");
            }
        }

        private static string JsString(string s)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try { _webView?.Dispose(); } catch { }
        }

        /// <summary>
        /// Streams ChatEventSink callbacks back into the WebView2 surface via
        /// ExecuteScriptAsync. Fire-and-forget; failures are swallowed and
        /// logged because the chat loop should not stall on UI hiccups.
        /// </summary>
        private sealed class WebViewSink : ChatEventSink
        {
            private readonly Dictionary<string, string> _names = new Dictionary<string, string>(StringComparer.Ordinal);
            private readonly ChatController _owner;
            private readonly string _assistantId;
            public WebViewSink(ChatController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                TraceLog.Write("Sink.OnTokenDelta len=" + (delta?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript(
                    "jameelcopilot.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                _names[callId ?? ""] = name;
                TraceLog.Write("Sink.OnToolCallStart " + name + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "jameelcopilot.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                _names.TryGetValue(callId ?? "", out var name);
                if (ok && name == ProposeMeetingTool.ToolName)
                {
                    var proposal = MeetingProposal.FromResult(resultJson);
                    if (proposal != null) _owner._meetings[callId] = proposal;
                }
                _owner._rules.Offered(callId, name, ok, resultJson);
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "jameelcopilot.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
                if (IsWriteTool(callId) && ok)
                {
                    _ = _owner.RunScript(
                        "jameelcopilot.appendAuditRow(" + JsString("Wrote: " + summary) + ");");
                }
            }
            public override void OnError(string message)
            {
                TraceLog.Write("Sink.OnError: " + message, "WebViewSink");
                _ = _owner.RunScript("jameelcopilot.showError(" + JsString(message ?? "") + ");");
            }
            public override void OnAssistantMessageComplete(string text)
            {
                TraceLog.Write("Sink.OnAssistantMessageComplete len=" + (text?.Length ?? 0), "WebViewSink");
            }
            public override void OnRoundBoundary()
            {
                TraceLog.Write("Sink.OnRoundBoundary", "WebViewSink");
            }

            // The C# side doesn't currently track which callIds correspond
            // to write tools - the model decides at runtime. For now we
            // detect write-tool intent by name prefix as a best effort.
            private static bool IsWriteTool(string callId)
            {
                // CallId doesn't carry the tool name; this method is a stub
                // until we plumb the name through. Returning false keeps the
                // audit row off until the explicit plumbing is in place
                // (avoids false-positive audit rows for read tools).
                return false;
            }
        }
    }
}
