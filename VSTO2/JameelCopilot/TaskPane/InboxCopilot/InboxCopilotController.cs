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
using JameelCopilot.TaskPane.Chat;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JameelCopilot.TaskPane.InboxCopilot
{
    /// <summary>
    /// Drives the per-Explorer Inbox Copilot chat surface. Mirrors
    /// Phase 2's ChatController but is anchored to an Explorer instead
    /// of an Inspector. Builds a fresh system prompt + quick-action
    /// chip set on every selection change and on every turn.
    /// </summary>
    public sealed class InboxCopilotController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CustomApiChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Outlook.Explorer _explorer;
        private readonly ExportBridge _exportBridge;
        private readonly PrefsBridge _prefs;
        private readonly ChatsBridge _chats;
        private readonly RulesBridge _rules;
        private readonly Dictionary<string, MeetingProposal> _meetings = new Dictionary<string, MeetingProposal>(StringComparer.Ordinal);
        private readonly Dictionary<string, ReplyDraft> _drafts = new Dictionary<string, ReplyDraft>(StringComparer.Ordinal);

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public InboxCopilotController(
            Control hostContainer,
            CustomApiChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store,
            Outlook.Explorer explorer,
            UserPrefsService prefs = null,
            ChatArchive chats = null,
            UserMemoryService memory = null,
            UserMemoryService prompts = null)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            _explorer = explorer;
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
            _prefs = new PrefsBridge(prefs, RunScript);
            _chats = new ChatsBridge(chats, memory, _store, RunScript, ChatArchive.MailboxScope, null, () => _turnInFlight, null,
                prompts, () => { if (_isReady && !_isDisposed) PushContextStripAndChips(); });
            _rules = new RulesBridge(surface, RunScript);
        }

        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "InboxCopilot");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                ShowFallback("Jameel Copilot needs the WebView2 runtime. Contact IT.");
                return;
            }
            _webView = new WebView2 { Dock = DockStyle.Fill };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.Navigate(WebView2Bootstrap.PageUrl);
                if (_explorer != null)
                {
                    _explorer.SelectionChange += OnExplorerSelectionChange;
                    _explorer.FolderSwitch += OnExplorerFolderSwitch;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "InboxCopilot");
                ShowFallback("WebView2 failed to initialize: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null) { _fallbackLabel.Text = message; return; }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message,
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private System.Windows.Forms.Timer _selectionDebounce;
        private bool _mailboxOnlyMode;

        private void OnExplorerSelectionChange()
        {
            if (_isDisposed || !_isReady) return;
            _mailboxOnlyMode = false;
            ScheduleContextStripUpdate();
        }

        private void OnExplorerFolderSwitch()
        {
            if (_isDisposed || !_isReady) return;
            _mailboxOnlyMode = false;
            ScheduleContextStripUpdate();
        }

        private void ScheduleContextStripUpdate()
        {
            if (_selectionDebounce == null)
            {
                _selectionDebounce = new System.Windows.Forms.Timer { Interval = 180 };
                _selectionDebounce.Tick += (s, e) =>
                {
                    _selectionDebounce.Stop();
                    if (!_isDisposed && _isReady && !_turnInFlight) PushContextStripAndChips();
                };
            }
            _selectionDebounce.Stop();
            _selectionDebounce.Start();
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
                TraceLog.Write("WebMessageReceived: " + type, "InboxCopilot");
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "InboxCopilot");
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
                        OpenMeeting((string)payload?["call"],
                            payload?["slot"]?.Type == JTokenType.Integer ? (int)payload["slot"] : -1,
                            payload?["room_index"]?.Type == JTokenType.Integer ? (int)payload["room_index"] : -1);
                        break;
                    case "open_message":
                        OpenMessage((string)payload?["id"]);
                        break;
                    case "chip_action":
                        RunChip((string)payload?["id"]);
                        break;
                    case "set_scope":
                        _mailboxOnlyMode = string.Equals((string)payload?["scope"], "mailbox", StringComparison.OrdinalIgnoreCase);
                        PushContextStripAndChips();
                        break;
                    case "tidy_apply":
                        _ = ApplyTidyAsync((string)payload?["id"], payload?["picks"] as JArray);
                        break;
                    case "tidy_undo":
                        _ = UndoTidyAsync((string)payload?["id"]);
                        break;
                    case "priority_reply":
                        _ = ReplyToListedAsync((string)payload?["id"], followUp: false);
                        break;
                    case "followup":
                        _ = ReplyToListedAsync((string)payload?["id"], followUp: true);
                        break;
                    case "task_add":
                        _ = AddTaskAsync((string)payload?["id"], payload?["index"]?.Type == JTokenType.Integer ? (int)payload["index"] : -1);
                        break;
                    case "draft_action":
                        HandleDraftAction((string)payload?["id"], (string)payload?["action"]);
                        break;
                    case "copy":
                        var clip = (string)payload?["text"] ?? _store.LatestAssistantText() ?? _store.ExportForClipboard();
                        try { Clipboard.SetText(clip ?? ""); } catch { }
                        break;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("HandleHostMessageAsync EXCEPTION: " + ex, "InboxCopilot");
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "InboxCopilot");
            _isReady = true;
            _ = RunScript("jameelcopilot.applyTheme('light');");
            PushContextStripAndChips();
            TraceLog.Write("OnWebViewReady completed", "InboxCopilot");
        }

        private void PushContextStripAndChips()
        {
            try
            {
                CurrentSelectionResult sel = null;
                string folderName = "";
                int unreadCount = 0, totalCount = 0;
                try
                {
                    sel = _surface.GetCurrentSelection(includeFullBodies: false, maxItems: 1);
                    folderName = sel?.Folder ?? "";
                    Outlook.MAPIFolder folder = null;
                    try { folder = _explorer?.CurrentFolder; }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        try { folder = Globals.ThisAddIn?.Application?.ActiveExplorer()?.CurrentFolder; } catch { }
                    }
                    try
                    {
                        if (folder != null)
                        {
                            if (string.IsNullOrEmpty(folderName)) folderName = folder.Name ?? "";
                            unreadCount = folder.UnReadItemCount;
                            totalCount = folder.Items.Count;
                        }
                    }
                    catch { }
                }
                catch (Exception ex) { TraceLog.Write("PushContextStrip surface error: " + ex.Message, "InboxCopilot"); }
                if (string.IsNullOrWhiteSpace(folderName)) folderName = "Mailbox";

                var ctx = new JObject(
                    new JProperty("folder", folderName),
                    new JProperty("unread_count", unreadCount),
                    new JProperty("total_count", totalCount),
                    new JProperty("mailbox_only", _mailboxOnlyMode));
                var hasSelection = sel != null && sel.Count > 0 && sel.Messages != null && sel.Messages.Count > 0;
                if (hasSelection)
                {
                    var first = sel.Messages[0];
                    ctx.Add("selection", new JObject(
                        new JProperty("count", sel.Count),
                        new JProperty("subject", first.Subject ?? ""),
                        new JProperty("from", first.From ?? "")));
                }
                _ = RunScript("jameelcopilot.setContextStrip(" +
                    ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");

                var effectiveSelCount = _mailboxOnlyMode ? 0 : (sel?.Count ?? 0);
                var chips = QuickActionChip.ToJson(QuickActionChip.Groups(effectiveSelCount));
                if (effectiveSelCount == 0 && chips.Count > 0 && chips[0] is JObject mailboxGroup && mailboxGroup["chips"] is JArray mChips)
                {
                    var newMeetingChip = new JObject(
                        new JProperty("label", "+ New meeting"),
                        new JProperty("mode", "fill"),
                        new JProperty("prompt", "Schedule a 1-hour meeting with [names] about [topic] via [HO1 room / Microsoft Teams]."));
                    mChips.Insert(Math.Min(1, mChips.Count), newMeetingChip);
                }
                chips.Add(_chats.PromptsGroup());
                _ = RunScript("jameelcopilot.setQuickActions(" + chips.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStripAndChips error: " + ex.Message, "InboxCopilot");
            }
        }

        private async Task StartTurnAsync(string userText)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxCopilot");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxCopilot");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            var isTranslateChip = userText.StartsWith("Translate the main message body of the selected email", StringComparison.Ordinal);
            var isSummarizeChip = userText.StartsWith("Summarize the selected message and the rest of its conversation thread.", StringComparison.Ordinal);
            var isScheduleChip = userText.StartsWith("Help me schedule a meeting about the selected email", StringComparison.Ordinal);
            var shownUserText = isTranslateChip ? "Translate the selected email."
                : isSummarizeChip ? "Summarize this email."
                : isScheduleChip ? "Schedule a meeting for this email."
                : userText;
            await RunScript("jameelcopilot.appendUserMessage(" + JsString(shownUserText) + ");");
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                var initialSnapshot = _store.Snapshot();
                MessageDetail fastMsg = (isTranslateChip || isSummarizeChip)
                    ? await Task.Run(() => _surface?.GetCurrentSelection(includeFullBodies: true, maxItems: 1)?.Messages?.FirstOrDefault(), _activeCts.Token)
                    : null;

                TurnResult result;
                var sink = new WebViewSink(this, assistantId);
                if (isTranslateChip && fastMsg != null && !string.IsNullOrWhiteSpace(fastMsg.BodyPlaintext))
                {
                    var ownOnly = ComposeBody.FromPlainText(fastMsg.BodyPlaintext).Own;
                    var bodyText = (!string.IsNullOrWhiteSpace(ownOnly) ? ownOnly : ComposeBody.StripSignatureAndFooter(fastMsg.BodyPlaintext)).Trim();
                    var letters = bodyText.Count(char.IsLetter);
                    var arabicLetters = bodyText.Count(c => c >= '\u0600' && c <= '\u06FF');
                    var intoArabic = letters == 0 || ((double)arabicLetters / letters) < 0.35;
                    var action = intoArabic ? CustomApiChatService.ActionType.Arabic : CustomApiChatService.ActionType.English;
                    var ctx = CustomApiChatService.NewActionContext(action);
                    ctx.Temperature = _prefs.Temperature;
                    var prompt = "Action: translate into " + action + ". The text below is what the user wrote; work on this text only.\n\n" + bodyText;
                    result = await _chat.RunTurnAsync(ctx, prompt, _toolHost, sink, _activeCts.Token);
                    if (result.StopReason != StopReason.Error && !string.IsNullOrEmpty(result.FinalAssistantText))
                    {
                        _store.Append(ChipLine("Translate the selected message."));
                        _store.Append(HistoryMessage("assistant", result.FinalAssistantText));
                        _ = Task.Run(() => _chats.Save());
                    }
                }
                else if (isSummarizeChip && fastMsg != null)
                {
                    var language = StructuredAnswers.ReplyLanguage(initialSnapshot, "the language of the email");
                    var attSb = new System.Text.StringBuilder();
                    foreach (var a in fastMsg.Attachments ?? new AttachmentSummary[0])
                    {
                        attSb.Append("\nAttachment: ").Append(a.Filename ?? "");
                        if (!string.IsNullOrEmpty(a.ContentPlaintext))
                        {
                            attSb.Append("\nAttachment content:\n").Append(a.ContentPlaintext);
                        }
                    }
                    var emailData = "From: " + (fastMsg.From ?? "") + "\nTo: " + string.Join("; ", fastMsg.To ?? new string[0])
                        + "\nCc: " + string.Join("; ", fastMsg.Cc ?? new string[0])
                        + "\nReceived: " + fastMsg.ReceivedAt.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture)
                        + "\nSubject: " + (fastMsg.Subject ?? "")
                        + "\n\n" + (fastMsg.BodyPlaintext ?? "") + attSb;
                    var ctx = new ConversationContext
                    {
                        UseTools = false,
                        SystemInstructions = "Summarize the email (including its quoted conversation thread and any extracted attachment content) concisely in " + language + ". Highlight: (1) Key points, (2) Decisions or status, (3) Action items / asks. Use clean bullet points. Do not search other unrelated emails.",
                        Context = emailData,
                        Temperature = _prefs.Temperature,
                        Memory = await Task.Run(() => _chats.Memory, _activeCts.Token),
                    };
                    result = await _chat.RunTurnAsync(ctx, "Summarize this email.", _toolHost, sink, _activeCts.Token);
                    if (result.StopReason != StopReason.Error && !string.IsNullOrEmpty(result.FinalAssistantText))
                    {
                        _store.Append(ChipLine("Summarize the selected message."));
                        _store.Append(HistoryMessage("assistant", result.FinalAssistantText));
                        _ = Task.Run(() => _chats.Save());
                    }
                }
                else
                {
                    // Off the UI thread so the WebView2 UI never freezes while reading Outlook COM.
                var contextString = await Task.Run(() => BuildContextForCurrentState(), _activeCts.Token);
                var ctx = new ConversationContext
                {
                    // One pane answers questions and writes reports: both texts, every turn.
                    SystemInstructions = SystemPromptBuilder.Pane("inbox") + "\n\n" + SystemPromptBuilder.Pane("reports"),
                    Context = contextString,
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    Temperature = _prefs.Temperature,
                    Memory = await Task.Run(() => _chats.Memory, _activeCts.Token),
                };

                result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                if (result.StopReason != StopReason.Error)
                {
                    for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                    {
                        _store.Append(ctx.History[i]);
                    }
                    _ = Task.Run(() => _chats.Save());
                }
                }
                var finalText = (result.FinalAssistantText ?? "").Trim();
                if (result.StopReason == StopReason.Completed && finalText.Length > 0 && !_mailboxOnlyMode)
                {
                    var selMsg = await Task.Run(() => _surface?.GetCurrentSelection(includeFullBodies: false, maxItems: 1)?.Messages?.FirstOrDefault());
                    if (selMsg != null && !string.IsNullOrEmpty(selMsg.Id))
                    {
                        _drafts[assistantId] = new ReplyDraft { Message = selMsg, Text = finalText, FollowUp = false };
                    }
                }
                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TraceLog.Write("StartTurnAsync COMException: " + ex.Message, "InboxCopilot");
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(assistantId) + ", {});");
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
                TraceLog.Write("<< StartTurnAsync", "InboxCopilot");
            }
        }

        // A chip the pane runs itself.
        private void RunChip(string chipId)
        {
            switch (chipId)
            {
                case QuickActionChip.ReplyChipId: _ = RunReplyDraftAsync(null); break;
                case QuickActionChip.PrioritizeChipId: _ = RunPrioritizeAsync(); break;
                case QuickActionChip.TidyChipId: _ = RunTidyAsync(); break;
                case QuickActionChip.DayChipId: _ = RunDayAsync(); break;
                case QuickActionChip.TodosChipId: _ = RunTodosAsync(); break;
                case QuickActionChip.WaitingChipId: _ = RunWaitingAsync(); break;
            }
        }

        // -- Prioritize my inbox ----------------------------------------------

        // Structured answers are JSON for Jameel Copilot's screen: nothing of them streams to the page.
        private sealed class QuietSink : ChatEventSink { }

        /// <summary>
        /// Ranks the recent unread mail (high / medium / low, with a reason)
        /// and shows it as a card whose rows open the email.
        /// </summary>
        private async Task RunPrioritizeAsync()
        {
            if (_turnInFlight || !_isReady || _surface == null) return;
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var id = "prio_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            await RunScript("jameelcopilot.appendUserMessage(" + JsString("Prioritize my inbox.") + ");");
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(id) + ", " + JsString("Reading your recent unread mail...") + ");");

            try
            {
                var now = DateTimeOffset.Now;
                var ct = _activeCts.Token;
                // Off the UI thread, as the tools run: a search can wait for an Outlook event on the UI thread.
                var candidates = await Task.Run(() =>
                {
                    var unread = _surface.SearchMessages(InboxPriorities.UnreadArgs(now), ct)?.Messages;
                    var recent = (unread?.Count ?? 0) >= 10 ? null : _surface.SearchMessages(InboxPriorities.RecentArgs(now), ct)?.Messages;
                    return InboxPriorities.Candidates(unread, recent);
                }, ct);

                if (candidates.Count == 0)
                {
                    await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Nothing unread in the last 7 days.") + ");");
                    await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {});");
                    return;
                }

                var language = StructuredAnswers.ReplyLanguage(_store.Snapshot());
                var ctx = StructuredAnswers.NewContext(InboxPriorities.Instructions(language),
                    StructuredAnswers.MessageList(candidates), _chats.Memory);
                var result = await _chat.RunTurnAsync(ctx, "Sort these emails.", _toolHost, new QuietSink(), ct);
                var answer = StructuredAnswers.ParseObject(result.FinalAssistantText);

                if (result.StopReason != StopReason.Completed || answer == null)
                {
                    var why = result.StopReason == StopReason.Cancelled ? "Stopped."
                        : result.StopReason == StopReason.Error ? (result.ErrorMessage ?? "The AI service returned an error.")
                        : "Jameel Copilot could not sort the mail this time. Try again.";
                    await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString(why) + ");");
                    await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {error:true});");
                    return;
                }

                var items = InboxPriorities.Parse(answer, candidates);
                await RunScript("jameelcopilot.showPriorities(" + JsString(id) + ", " +
                                InboxPriorities.ToJson(items).ToString(Newtonsoft.Json.Formatting.None) + ");");
                _store.Append(ChipLine("Prioritize my inbox."));
                _store.Append(HistoryMessage("assistant", InboxPriorities.AsText(items)));
                _chats.Save();
            }
            catch (OperationCanceledException)
            {
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Stopped.") + ");");
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunPrioritizeAsync failed: " + ex.GetType().Name, "InboxCopilot");
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Jameel Copilot could not sort the mail this time. Try again.") + ");");
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

        // -- Cards from a chip: Start my day, Action items, Waiting for a reply --

        private const string DayFailed = "Jameel Copilot could not prepare your day this time. Try again.";
        private const string TodosFailed = "Jameel Copilot could not list the action items this time. Try again.";
        private const string WaitingFailed = "Jameel Copilot could not check your sent mail this time. Try again.";

        // The to-dos of each card and the ones made tasks. Kept here, never taken from the page.
        private readonly Dictionary<string, List<TodoItem>> _todos = new Dictionary<string, List<TodoItem>>(StringComparer.Ordinal);
        private readonly HashSet<string> _tasksAdded = new HashSet<string>(StringComparer.Ordinal);

        // Why a card shows words instead: nothing to show (not an error), or what went wrong.
        private sealed class CardProblem : Exception
        {
            public CardProblem(string message, bool isError = true) : base(message) { IsError = isError; }
            public bool IsError { get; }
        }

        /// <summary>
        /// One chip whose answer is a card: the turn's gate, the lines on the
        /// page while Jameel Copilot works, and what a failure shows. <paramref name="fill"/>
        /// shows the card or throws a <see cref="CardProblem"/>.
        /// </summary>
        private async Task RunCardAsync(string kind, string asked, string working, string failed, Func<string, CancellationToken, Task> fill)
        {
            if (_turnInFlight || !_isReady || _surface == null) return;
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var id = kind + "_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            await RunScript("jameelcopilot.appendUserMessage(" + JsString(asked) + ");");
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(id) + ", " + JsString(working) + ");");
            try
            {
                await fill(id, _activeCts.Token);
            }
            catch (OperationCanceledException)
            {
                await EndCardAsync(id, "Stopped.", "{stopped:true}");
            }
            catch (CardProblem problem)
            {
                await EndCardAsync(id, problem.Message, problem.IsError ? "{error:true}" : "{}");
            }
            catch (Exception ex)
            {
                TraceLog.Write("Card " + kind + " failed: " + ex.GetType().Name, "InboxCopilot");
                await EndCardAsync(id, failed, "{error:true}");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("jameelcopilot.setComposerEnabled(true, false);");
            }
        }

        private async Task EndCardAsync(string id, string text, string opts)
        {
            await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString(text) + ");");
            await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", " + opts + ");");
        }

        // The model's data for a card, or a CardProblem that says why there is none.
        private async Task<JObject> AskForCardAsync(string instructions, string context, string request, string failed, CancellationToken ct)
        {
            var ctx = StructuredAnswers.NewContext(instructions, context, _chats.Memory);
            var result = await _chat.RunTurnAsync(ctx, request, _toolHost, new QuietSink(), ct);
            if (result.StopReason == StopReason.Cancelled) throw new OperationCanceledException();
            if (result.StopReason == StopReason.Error) throw new CardProblem(result.ErrorMessage ?? "The AI service returned an error.");
            var answer = StructuredAnswers.ParseObject(result.FinalAssistantText);
            if (result.StopReason != StopReason.Completed || answer == null) throw new CardProblem(failed);
            return answer;
        }

        // What the chat keeps of a card, so a follow-up question knows it.
        private void KeepCard(string asked, string asText)
        {
            _store.Append(ChipLine(asked));
            _store.Append(HistoryMessage("assistant", asText));
            _chats.Save();
        }

        /// <summary>"Start my day": to-dos and topics from the recent unread mail, and today's meetings from the calendar.</summary>
        private Task RunDayAsync() => RunCardAsync("day", "Start my day.", "Reading your mail and your calendar...", DayFailed, async (id, ct) =>
        {
            var now = DateTimeOffset.Now;
            var today = now.LocalDateTime.Date;
            IReadOnlyList<CalendarEvent> events = null;   // stays null when the calendar cannot be read
            var candidates = await Task.Run(() =>
            {
                try
                {
                    events = _surface.ListCalendarEvents(today, today.AddDays(1), 20);
                }
                catch (Exception ex)
                {
                    TraceLog.Write("Day calendar failed: " + ex.GetType().Name, "InboxCopilot");
                }
                var unread = _surface.SearchMessages(InboxPriorities.UnreadArgs(now), ct)?.Messages;
                var recent = (unread?.Count ?? 0) >= 10 ? null : _surface.SearchMessages(InboxPriorities.RecentArgs(now), ct)?.Messages;
                return InboxPriorities.Candidates(unread, recent);
            }, ct);

            var todos = new List<TodoItem>();
            var topics = new List<DayTopic>();
            if (candidates.Count > 0)
            {
                var answer = await AskForCardAsync(DayBrief.Instructions(StructuredAnswers.ReplyLanguage(_store.Snapshot())),
                    StructuredAnswers.MessageList(candidates), "Prepare my day.", DayFailed, ct);
                todos = DayBrief.ParseTodos(answer, candidates);
                topics = DayBrief.ParseTopics(answer, candidates);
            }
            _todos[id] = todos;
            await RunScript("jameelcopilot.showDay(" + JsString(id) + ", " +
                            DayBrief.ToJson(today, todos, topics, events).ToString(Newtonsoft.Json.Formatting.None) + ");");
            KeepCard("Start my day.", DayBrief.AsText(todos, topics, events));
        });

        /// <summary>"Action items": what the selected email asks of the user, each row with "Add to tasks".</summary>
        private Task RunTodosAsync() => RunCardAsync("todos", "Action items of the selected email.", "Reading the selected email...", TodosFailed, async (id, ct) =>
        {
            var message = await Task.Run(() => _surface.GetCurrentSelection(includeFullBodies: true, maxItems: 1)?.Messages?.FirstOrDefault(), ct);
            if (message == null) throw new CardProblem("Select one email first, then press Action items.", isError: false);
            string ownerName = "", ownerEmail = "";
            await Task.Run(() => _surface.GetOwner(out ownerName, out ownerEmail), ct);
            var answer = await AskForCardAsync(
                DayBrief.ActionItemInstructions(StructuredAnswers.ReplyLanguage(_store.Snapshot(), "the language of the email")),
                DayBrief.EmailContext(message, ownerName, ownerEmail), "List what this email asks of me.", TodosFailed, ct);
            var source = new MessageSummary { Id = message.Id, From = message.From, Subject = message.Subject, ReceivedAt = message.ReceivedAt };
            var todos = DayBrief.ParseTodos(answer, null, source);
            _todos[id] = todos;
            await RunScript("jameelcopilot.showTodos(" + JsString(id) + ", " + DayBrief.TodosJson(todos).ToString(Newtonsoft.Json.Formatting.None) + ");");
            KeepCard("Action items of the selected email.", todos.Count == 0 ? "This email asks nothing of you." : DayBrief.AsText(todos, null, null));
        });

        /// <summary>"Waiting for a reply": what the user sent that nobody answered and that asked for something.</summary>
        private Task RunWaitingAsync() => RunCardAsync("wait", "Who has not replied to me?", "Looking through what you sent...", WaitingFailed, async (id, ct) =>
        {
            var now = DateTimeOffset.Now;
            var sent = await Task.Run(() => _surface.ListAwaitingReply(
                now.LocalDateTime.AddDays(-WaitingReplies.Days), now.LocalDateTime.AddDays(-WaitingReplies.GraceDays), WaitingReplies.MaxCandidates), ct);
            if (sent.Count == 0)
            {
                throw new CardProblem("Nothing you sent in the last " + WaitingReplies.Days + " days is still unanswered.", isError: false);
            }
            var answer = await AskForCardAsync(WaitingReplies.Instructions(StructuredAnswers.ReplyLanguage(_store.Snapshot())),
                WaitingReplies.Context(sent), "Which of these am I still waiting on?", WaitingFailed, ct);
            var items = WaitingReplies.Parse(answer, sent);
            if (items.Count == 0)
            {
                throw new CardProblem("You are not waiting on anyone: the unanswered emails you sent needed no reply.", isError: false);
            }
            await RunScript("jameelcopilot.showPriorities(" + JsString(id) + ", " + WaitingReplies.ToJson(items, now).ToString(Newtonsoft.Json.Formatting.None) + ");");
            KeepCard("Who has not replied to me?", WaitingReplies.AsText(items));
        });

        // "Add to tasks" on a to-do row: an Outlook task from the row Jameel Copilot kept, once.
        private async Task AddTaskAsync(string cardId, int index)
        {
            if (cardId == null || !_todos.TryGetValue(cardId, out var todos) || index < 0 || index >= todos.Count) return;
            var key = cardId + ":" + index;
            if (!_tasksAdded.Add(key)) return;
            var todo = todos[index];
            var added = false;
            try
            {
                added = await Task.Run(() => _surface != null && _surface.AddTask(todo.Text, todo.Due, DayBrief.TaskNote(todo)));
            }
            catch (Exception ex)
            {
                TraceLog.Write("AddTask failed: " + ex.GetType().Name, "InboxCopilot");
            }
            if (!added) _tasksAdded.Remove(key);
            await RunScript("jameelcopilot.markTodo(" + JsString(cardId) + ", " + index + ", " +
                            JsString(added ? "added to your Tasks" : "could not be added") + ", " + (added ? "true" : "false") + ");");
        }

        // -- Tidy my inbox ----------------------------------------------------

        // The proposals of each tidy card, and what was applied from it (for Undo). Kept here, never taken from the page.
        private readonly Dictionary<string, List<TidyItem>> _tidy = new Dictionary<string, List<TidyItem>>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<TidyItem>> _tidied = new Dictionary<string, List<TidyItem>>(StringComparer.Ordinal);

        /// <summary>Proposes "mark read" and "flag" for the recent unread, unflagged mail, as a card the user ticks.</summary>
        private async Task RunTidyAsync()
        {
            if (_turnInFlight || !_isReady || _surface == null) return;
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var id = "tidy_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            await RunScript("jameelcopilot.appendUserMessage(" + JsString("Tidy my inbox.") + ");");
            await RunScript("jameelcopilot.appendAssistantMessage(" + JsString(id) + ", " + JsString("Reading your unread mail...") + ");");
            try
            {
                var ct = _activeCts.Token;
                var now = DateTimeOffset.Now;
                var candidates = await Task.Run(() =>
                    InboxPriorities.Candidates(_surface.SearchMessages(InboxTidy.Args(now), ct)?.Messages, null), ct);
                if (candidates.Count == 0)
                {
                    await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Nothing unread and unflagged in the last 14 days.") + ");");
                    await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {});");
                    return;
                }

                var ctx = StructuredAnswers.NewContext(InboxTidy.Instructions(StructuredAnswers.ReplyLanguage(_store.Snapshot())),
                    StructuredAnswers.MessageList(candidates), _chats.Memory);
                var result = await _chat.RunTurnAsync(ctx, "Propose a tidy-up.", _toolHost, new QuietSink(), ct);
                var answer = StructuredAnswers.ParseObject(result.FinalAssistantText);
                if (result.StopReason != StopReason.Completed || answer == null)
                {
                    var why = result.StopReason == StopReason.Cancelled ? "Stopped."
                        : result.StopReason == StopReason.Error ? (result.ErrorMessage ?? "The AI service returned an error.")
                        : "Jameel Copilot could not propose a tidy-up this time. Try again.";
                    await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString(why) + ");");
                    await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {error:true});");
                    return;
                }

                var items = InboxTidy.Parse(answer, candidates);
                _tidy[id] = items;
                await RunScript("jameelcopilot.showTidy(" + JsString(id) + ", " + InboxTidy.ToJson(items).ToString(Newtonsoft.Json.Formatting.None) + ");");
                _store.Append(ChipLine("Tidy my inbox."));
                _store.Append(HistoryMessage("assistant", InboxTidy.AsText(items)));
                _chats.Save();
            }
            catch (OperationCanceledException)
            {
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Stopped.") + ");");
                await RunScript("jameelcopilot.finalizeAssistantMessage(" + JsString(id) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunTidyAsync failed: " + ex.GetType().Name, "InboxCopilot");
                await RunScript("jameelcopilot.replaceAssistantText(" + JsString(id) + ", " + JsString("Jameel Copilot could not propose a tidy-up this time. Try again.") + ");");
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

        // "Apply selected" on a tidy card: the ticked changes, from the host's own list.
        private async Task ApplyTidyAsync(string cardId, JArray picks)
        {
            if (cardId == null || !_tidy.TryGetValue(cardId, out var items) || _tidied.ContainsKey(cardId)) return;
            var wanted = new HashSet<string>((picks ?? new JArray()).Where(t => t.Type == JTokenType.String).Select(t => (string)t), StringComparer.Ordinal);
            var chosen = items.Where(i => wanted.Contains(i.Message.Id)).ToList();
            if (chosen.Count == 0) return;
            _tidied[cardId] = chosen;
            var done = await Task.Run(() => chosen.Count(i => Change(i, apply: true)));
            await RunScript("jameelcopilot.markTidy(" + JsString(cardId) + ", " + JsString(done + " of " + chosen.Count + " changed") + ", true);");
        }

        private async Task UndoTidyAsync(string cardId)
        {
            if (cardId == null || !_tidied.TryGetValue(cardId, out var applied)) return;
            _tidied.Remove(cardId);
            _tidy.Remove(cardId);
            var undone = await Task.Run(() => applied.Count(i => Change(i, apply: false)));
            await RunScript("jameelcopilot.markTidy(" + JsString(cardId) + ", " + JsString("undone: " + undone + " put back") + ", false);");
        }

        // One change, or its reverse. The emails were unread and unflagged when Jameel Copilot proposed it.
        private bool Change(TidyItem item, bool apply)
        {
            try
            {
                if (item.Action == "flag") _surface.FlagMessage(item.Message.Id, apply ? "todo" : "none");
                else _surface.MarkAsRead(item.Message.Id, apply);
                return true;
            }
            catch (Exception ex)
            {
                TraceLog.Write("Tidy change failed: " + ex.GetType().Name, "InboxCopilot");
                return false;
            }
        }

        // "Reply" on a ranked row: the same reply card as Draft a reply, for that email.
        private async Task ReplyToListedAsync(string messageId, bool followUp)
        {
            if (_turnInFlight || string.IsNullOrEmpty(messageId)) return;
            MessageDetail message = null;
            try
            {
                message = await Task.Run(() => _surface?.ReadMessage(messageId, includeFullBody: true));
            }
            catch (Exception ex)
            {
                TraceLog.Write("ReplyToListedAsync read failed: " + ex.GetType().Name, "InboxCopilot");
            }
            if (message == null)
            {
                await RunScript("jameelcopilot.appendAuditRow(" + JsString("That email could not be read. It may have been moved or deleted.") + ");");
                return;
            }
            await RunReplyDraftAsync(message, followUp);
        }

        // -- Draft a reply ---------------------------------------------------

        // A reply Jameel Copilot wrote for a message, shown as a card. The text and the
        // message it answers are kept here, not taken back from the page.
        private sealed class ReplyDraft
        {
            public MessageDetail Message;
            public string Text;
            public bool FollowUp;   // to the user's own sent email, which nobody answered
        }

        /// <summary>
        /// Writes the user's reply to a message (the selected one, or the one
        /// a card was written for) and shows it as a draft card. Nothing is
        /// created in Outlook until the user presses Open reply.
        /// </summary>
        private async Task RunReplyDraftAsync(MessageDetail message, bool followUp = false)
        {
            if (_turnInFlight || !_isReady || _surface == null) return;
            if (message == null)
            {
                message = await Task.Run(() => _surface.GetCurrentSelection(includeFullBodies: true, maxItems: 1)?.Messages?.FirstOrDefault());
            }
            if (message == null)
            {
                await RunScript("jameelcopilot.appendAuditRow(" + JsString("Select one email first, then press Draft a reply.") + ");");
                return;
            }

            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            var id = "draft_" + (++_nextMessageId);
            await RunScript("jameelcopilot.setComposerEnabled(false, true);");
            await RunScript("jameelcopilot.beginDraft(" + JsString(id) + ", " + JsString(followUp ? "Follow-up draft" : "Reply draft") + ");");

            try
            {
                // The same one-shot writing as in a reply window with nothing typed: no tools.
                const CustomApiChatService.ActionType action = CustomApiChatService.ActionType.Revise;
                string ownerName = "", ownerEmail = "";
                await Task.Run(() => _surface.GetOwner(out ownerName, out ownerEmail), _activeCts.Token);
                var compose = followUp ? ComposePrompt.FollowUpTo(message, ownerName, ownerEmail) : ComposePrompt.ReplyTo(message, ownerName, ownerEmail);
                var ctx = CustomApiChatService.NewActionContext(action);
                ctx.SystemInstructions = followUp ? ComposePrompt.FollowUpInstructions : ComposePrompt.Instructions(action, compose);
                ctx.Context = ComposePrompt.Context(compose, includeDraft: false, ownEarlier: followUp);
                ctx.Temperature = _prefs.Temperature;
                ctx.Memory = await Task.Run(() => _chats.Memory, _activeCts.Token);
                var userMessage = followUp ? ComposePrompt.FollowUpRequest : ComposePrompt.UserMessage(action, compose, "");

                var result = await _chat.RunTurnAsync(ctx, userMessage, _toolHost, new WebViewSink(this, id), _activeCts.Token);
                var text = (result.FinalAssistantText ?? "").Trim();
                if (result.StopReason != StopReason.Completed || text.Length == 0)
                {
                    await RunScript(ComposeChips.FinalizeScript(id, ComposeChips.DiscardOnly(),
                        result.StopReason == StopReason.Cancelled ? "stopped" : "not finished", null, result.StopReason == StopReason.Error));
                    return;
                }

                _drafts[id] = new ReplyDraft { Message = message, Text = text, FollowUp = followUp };
                _store.Append(ChipLine(followUp ? "Draft a follow-up to my unanswered email." : "Draft a reply to the selected email."));
                _store.Append(HistoryMessage("assistant", text));
                _ = Task.Run(() => _chats.Save());
                await RunScript(ComposeChips.FinalizeScript(id, followUp ? ComposeChips.FollowUpButtons() : ComposeChips.ReplyWindowButtons(), "not in an email yet",
                    "Text in [brackets] is for you to fill in. No signature is written: Outlook adds yours.", false));
            }
            catch (OperationCanceledException)
            {
                await RunScript(ComposeChips.FinalizeScript(id, ComposeChips.DiscardOnly(), "stopped", null, false));
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunReplyDraftAsync failed: " + ex.GetType().Name, "InboxCopilot");
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

        // The user's side of a chip, in Jameel Copilot's words: kept for the chat, not taken for the language the user writes in.
        private static JObject ChipLine(string text)
        {
            var line = HistoryMessage("user", text);
            line.Add(StructuredAnswers.ChipLineMark, true);
            return line;
        }

        // A button on a reply draft card.
        private void HandleDraftAction(string id, string what)
        {
            if (id == null || !_drafts.TryGetValue(id, out var draft)) return;
            switch (what)
            {
                case "reply":
                case "reply_all":
                    var opened = false;
                    try
                    {
                        // Opens Outlook's own reply window with the text in it. The user sends it, or does not.
                        opened = _surface.OpenReplyWithText(draft.Message.Id, what == "reply_all", draft.Text);
                    }
                    catch (Exception ex)
                    {
                        TraceLog.Write("OpenReplyWithText failed: " + ex.GetType().Name, "InboxCopilot");
                    }
                    _ = RunScript(ComposeChips.MarkScript(id, opened, "opened in a reply window", "could not open the reply; copy the text instead"));
                    break;
                case "regenerate":
                    _ = RunReplyDraftAsync(draft.Message, draft.FollowUp);
                    break;
                case "discard":
                    _drafts.Remove(id);
                    break;
            }
        }

        private string BuildContextForCurrentState()
        {
            try
            {
                var sel = _mailboxOnlyMode ? null : _surface?.GetCurrentSelection(includeFullBodies: false, maxItems: 1);
                int unreadCount = 0, totalCount = 0;
                string folderName = !string.IsNullOrWhiteSpace(sel?.Folder) ? sel.Folder : "Inbox";
                try
                {
                    Outlook.MAPIFolder folder = null;
                    try { folder = _explorer?.CurrentFolder; }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        try { folder = Globals.ThisAddIn?.Application?.ActiveExplorer()?.CurrentFolder; } catch { }
                    }
                    if (folder != null)
                    {
                        if (!string.IsNullOrWhiteSpace(folder.Name)) folderName = folder.Name;
                        unreadCount = folder.UnReadItemCount;
                        totalCount = folder.Items.Count;
                    }
                }
                catch { }
                var org = (_surface as IOrgSurface)?.GetOrgContext();
                return InboxCopilotPromptBuilder.BuildContext(folderName, unreadCount, totalCount, sel, org);
            }
            catch (Exception ex)
            {
                TraceLog.Write("BuildContext error: " + ex.GetType().Name, "InboxCopilot");
                return "";
            }
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
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
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "InboxCopilot");
            }
        }

        // A time picked on a meeting card: Outlook's meeting window, filled in, for the user to send.
        private void OpenMeeting(string callId, int slot, int roomIndex = -1)
        {
            var opened = false;
            try
            {
                if (callId != null && _meetings.TryGetValue(callId, out var proposal) && _surface != null)
                {
                    if (roomIndex >= 0 && proposal.RoomOptions != null && roomIndex < proposal.RoomOptions.Count)
                    {
                        proposal = new MeetingProposal
                        {
                            Subject = proposal.Subject,
                            Attendees = proposal.Attendees,
                            OptionalAttendees = proposal.OptionalAttendees,
                            Slots = proposal.Slots,
                            Agenda = proposal.Agenda,
                            Room = proposal.RoomOptions[roomIndex].Value,
                            RoomOptions = proposal.RoomOptions,
                        };
                    }
                    opened = _surface.OpenMeetingRequest(proposal, slot);
                }
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
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _selectionDebounce?.Stop(); _selectionDebounce?.Dispose(); } catch { }
            try { _activeCts?.Cancel(); } catch { }
            try
            {
                if (_explorer != null)
                {
                    _explorer.SelectionChange -= OnExplorerSelectionChange;
                    _explorer.FolderSwitch -= OnExplorerFolderSwitch;
                }
            }
            catch { }
            try { _webView?.Dispose(); } catch { }
        }

        private sealed class WebViewSink : ChatEventSink
        {
            private readonly Dictionary<string, string> _names = new Dictionary<string, string>(StringComparer.Ordinal);
            private readonly InboxCopilotController _owner;
            private readonly string _assistantId;
            public WebViewSink(InboxCopilotController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                _ = _owner.RunScript("jameelcopilot.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                _names[callId ?? ""] = name;
                TraceLog.Write("Sink.OnToolCallStart " + name + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript("jameelcopilot.appendToolCallCard(" +
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
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " resultLen=" + (resultJson?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript("jameelcopilot.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
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
        }
    }
}
