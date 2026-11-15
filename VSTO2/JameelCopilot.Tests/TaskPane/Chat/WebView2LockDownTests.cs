using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Tests.Helpers;
using JameelCopilot.TaskPane.Chat;
using Xunit;

namespace JameelCopilot.Tests.TaskPane.Chat
{
    /// <summary>
    /// The lock-down in a real WebView2, started the way the panes start it
    /// (<c>WebView2Bootstrap.InitializeAsync</c>, then the chat page). The
    /// other guards read source text; this one fails if the running view is
    /// not locked down. "blocked.invalid" never resolves, and nothing here
    /// leaves the PC. Skipped where the WebView2 Runtime is not installed.
    /// </summary>
    public class WebView2LockDownTests
    {
        private sealed class WebView2FactAttribute : FactAttribute
        {
            public WebView2FactAttribute()
            {
                if (!WebView2Bootstrap.IsRuntimeInstalled()) Skip = "The WebView2 Runtime is not installed on this machine.";
            }
        }

        // What a crafted email could get the model to write.
        private const string HostileMarkdown =
            "Hello [link](https://blocked.invalid/?q=secret) ![img](https://blocked.invalid/i.png)\n\n"
            + "<img src=x onerror=window.__fromMarkdown=1> <script>window.__fromMarkdown=2</script>\n";

        // What a script would try if the renderer ever let one through.
        private const string HostileScript =
            "(function(){"
            + " var s=document.createElement('script'); s.textContent='window.__inline=1'; document.body.appendChild(s);"
            + " var d=document.createElement('div'); d.innerHTML='<img src=\"data:image/gif;base64,R0lGODlhAQABAAAAACw=\" onload=\"window.__handler=1\" onerror=\"window.__handler=1\">'; document.body.appendChild(d);"
            + " try { fetch('https://blocked.invalid/f').then(function(){window.__fetch='sent'}).catch(function(){window.__fetch='blocked'}); } catch(e) { window.__fetch='blocked'; }"
            + " window.__popup = window.open('https://blocked.invalid/popup') ? 'opened' : 'none';"
            + " setTimeout(function(){ location.href='https://blocked.invalid/nav'; }, 0);"
            + " return 1; })()";

        private const string Observe =
            "JSON.stringify({"
            + " anchors: document.querySelectorAll('[data-message-id=\"a1\"] a').length,"
            + " images: document.querySelectorAll('[data-message-id=\"a1\"] img').length,"
            + " scripts: document.querySelectorAll('[data-message-id=\"a1\"] script').length,"
            + " shownText: document.querySelector('[data-message-id=\"a1\"] .msg-content').textContent,"
            + " fromMarkdown: window.__fromMarkdown || 0, inline: window.__inline || 0, handler: window.__handler || 0,"
            + " fetch: window.__fetch || 'pending', popup: window.__popup || 'pending', href: location.href })";

        [WebView2Fact]
        public void TheChatPage_StartedAsThePanesStartIt_IsLockedDown()
        {
            var root = Path.Combine(Path.GetTempPath(), "jc-lockdown-test-" + Guid.NewGuid().ToString("N"));
            var settings = new Dictionary<string, bool>();
            var messageSources = new List<string>();
            var popupsLeftUnhandled = 0;
            JObject page = null;

            try
            {
                WithWebView(async web =>
                {
                    await Within(WebView2Bootstrap.InitializeAsync(web, Path.Combine(root, "WebUI"), Path.Combine(root, "WebView2Data")), 60);
                    var core = web.CoreWebView2;
                    settings["dev tools"] = core.Settings.AreDevToolsEnabled;
                    settings["context menus"] = core.Settings.AreDefaultContextMenusEnabled;
                    settings["host objects"] = core.Settings.AreHostObjectsAllowed;
                    settings["script dialogs"] = core.Settings.AreDefaultScriptDialogsEnabled;
                    settings["autofill"] = core.Settings.IsGeneralAutofillEnabled;
                    settings["password save"] = core.Settings.IsPasswordAutosaveEnabled;
                    settings["reputation lookup"] = core.Settings.IsReputationCheckingRequired;
                    core.WebMessageReceived += (s, e) => messageSources.Add(e.Source);
                    core.NewWindowRequested += (s, e) => { if (!e.Handled) popupsLeftUnhandled++; };

                    var loaded = new TaskCompletionSource<bool>();
                    core.NavigationCompleted += (s, e) => loaded.TrySetResult(e.IsSuccess);
                    core.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                    Assert.True(await Within(loaded.Task, 30), "The chat page did not load under the lock-down.");

                    await core.ExecuteScriptAsync(
                        "jameelcopilot.appendAssistantMessage('a1',''); jameelcopilot.appendTextDelta('a1', "
                        + JsonConvert.SerializeObject(HostileMarkdown) + "); jameelcopilot.finalizeAssistantMessage('a1', {}); 1");
                    await core.ExecuteScriptAsync(HostileScript);
                    await Task.Delay(2000);   // time for the fetch, the pop-up and the navigation to be refused

                    page = JObject.Parse(JsonConvert.DeserializeObject<string>(await core.ExecuteScriptAsync(Observe)));
                });
            }
            finally
            {
                // The browser processes may still hold the profile; the folder is under %TEMP%.
                try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }

            Assert.All(settings, s => Assert.False(s.Value, s.Key + " is still on"));

            // Model output is shown as text: no link, image or script element, and nothing in it ran.
            Assert.Equal((0, 0, 0), ((int)page["anchors"], (int)page["images"], (int)page["scripts"]));
            Assert.Contains("https://blocked.invalid/?q=secret", (string)page["shownText"]);
            Assert.Equal(0, (int)page["fromMarkdown"]);

            // Behind the renderer: the page's policy stops injected script and every request that leaves.
            Assert.Equal(0, (int)page["inline"]);
            Assert.Equal(0, (int)page["handler"]);
            Assert.Equal("blocked", (string)page["fetch"]);

            // Behind the policy: the view opens no window and goes nowhere.
            Assert.Equal("none", (string)page["popup"]);
            Assert.Equal(0, popupsLeftUnhandled);
            Assert.StartsWith("https://" + WebView2Bootstrap.VirtualHost + "/", (string)page["href"]);

            // The page talks to the add-in from the one source the add-in listens to.
            Assert.NotEmpty(messageSources);
            Assert.All(messageSources, source => Assert.True(WebView2Bootstrap.IsTrustedSource(source), source));
        }

        // The response-style picker in the real page: hidden until the add-in
        // says which style is in use, and a change reaches the add-in as a
        // message from the page.
        [WebView2Fact]
        public void TheStylePicker_ShowsTheChoice_AndReportsAChange()
        {
            var root = Path.Combine(Path.GetTempPath(), "jc-lockdown-test-" + Guid.NewGuid().ToString("N"));
            var messages = new List<string>();
            var sources = new List<string>();
            var seen = new Dictionary<string, string>();

            try
            {
                WithWebView(async web =>
                {
                    await Within(WebView2Bootstrap.InitializeAsync(web, Path.Combine(root, "WebUI"), Path.Combine(root, "WebView2Data")), 60);
                    var core = web.CoreWebView2;
                    var loaded = new TaskCompletionSource<bool>();
                    core.NavigationCompleted += (s, e) => loaded.TrySetResult(e.IsSuccess);
                    core.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                    Assert.True(await Within(loaded.Task, 30), "The chat page did not load.");
                    core.WebMessageReceived += (s, e) =>
                    {
                        sources.Add(e.Source);
                        messages.Add(e.TryGetWebMessageAsString());
                    };

                    // The mark in the header is one of the page's own files, loaded as an image.
                    seen["mark element"] = await core.ExecuteScriptAsync(
                        "(function(){ var m = document.querySelector('header .app-mark'); return m ? m.getAttribute('src') : null; })()");
                    await core.ExecuteScriptAsync(
                        "(function(){ window.__mark = 'pending'; var i = new Image();"
                        + " i.onload = function(){ window.__mark = 'loaded'; }; i.onerror = function(){ window.__mark = 'error'; };"
                        + " i.src = 'jc-mark.svg'; return 1; })()");

                    const string hidden = "document.getElementById('stylePicker').hidden";
                    const string value = "document.getElementById('styleSelect').value";
                    seen["hidden at start"] = await core.ExecuteScriptAsync(hidden);
                    await core.ExecuteScriptAsync("jameelcopilot.setStyle('creative'); 1");
                    seen["hidden after setStyle"] = await core.ExecuteScriptAsync(hidden);
                    seen["value after setStyle"] = await core.ExecuteScriptAsync(value);
                    await core.ExecuteScriptAsync(
                        "(function(){ var s = document.getElementById('styleSelect'); s.value = 'precise';"
                        + " s.dispatchEvent(new Event('change')); return 1; })()");
                    await Task.Delay(500);
                    await core.ExecuteScriptAsync("jameelcopilot.setStyle(null); 1");
                    seen["hidden after setStyle(null)"] = await core.ExecuteScriptAsync(hidden);
                    seen["mark"] = await core.ExecuteScriptAsync("window.__mark");
                });
            }
            finally
            {
                try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }

            Assert.Equal("\"jc-mark.svg\"", seen["mark element"]);
            Assert.Equal("\"loaded\"", seen["mark"]);
            Assert.Equal("true", seen["hidden at start"]);
            Assert.Equal("false", seen["hidden after setStyle"]);
            Assert.Equal("\"creative\"", seen["value after setStyle"]);
            Assert.Equal("true", seen["hidden after setStyle(null)"]);

            var styleMessages = messages.Where(m => m != null && m.Contains("set_style")).ToList();
            Assert.Equal(new[] { "{\"type\":\"set_style\",\"payload\":{\"style\":\"precise\"}}" }, styleMessages);
            Assert.All(sources, source => Assert.True(WebView2Bootstrap.IsTrustedSource(source), source));
        }

        // The one pane: saved chats, memory, chip groups and draft cards, driven
        // the way the host drives them. Titles, notes and drafts are text from
        // the mailbox or the model: they must be shown as text, never parsed.
        private const string DriveThePane = @"
(function(){
  var o = jameelcopilot, r = {};
  var hostile = '<img src=x onerror=window.__pane=1><b>bold</b>';
  function texts(sel) { return [].map.call(document.querySelectorAll(sel), function(n){ return n.textContent; }).join('|'); }

  o.setChats([{id:'c1', title: hostile, thread:true, when:'Today 09:00'}, {id:'c2', title:'Second', thread:false, when:'Yesterday'}], 'c2');
  r.chatRows = document.querySelectorAll('#chatList .chat-row').length;
  r.chatTitle = document.querySelector('#chatList .chat-title').textContent;
  r.chatParsed = document.querySelectorAll('#chatList img, #chatList b').length;
  r.current = document.querySelector('#chatList .chat-row.current .chat-title').textContent;
  document.querySelector('#chatList .chat-open').click();
  var del = document.querySelectorAll('#chatList .chat-del')[1];
  del.click();
  r.delAsks = del.textContent;
  del.click();
  var clearAll = document.getElementById('btnClearChats');
  clearAll.click();
  r.clearAsks = clearAll.textContent;

  o.setMemory([{id:'m1', text: hostile}]);
  r.memText = document.querySelector('#memoryList .mem-text').textContent;
  r.memParsed = document.querySelectorAll('#memoryList img, #memoryList b').length;
  var memDel = document.querySelectorAll('#memoryList .mem-actions button')[1];
  memDel.click();
  r.memDelAsks = memDel.textContent;
  memDel.click();
  document.getElementById('memoryInput').value = '  Keep replies short.  ';
  document.getElementById('btnMemoryAdd').click();

  o.setQuickActions([
    {title:'This email', chips:[{label:'Draft a reply', mode:'action', id:'reply'}, {label:'Summarize', prompt:'Summarize it.'}]},
    {title:'Reports', chips:[{label:'With a person', prompt:'Summarize my mail with [name].', mode:'fill'}]}]);
  r.titles = texts('#quickActions .qa-title');
  var chips = document.querySelectorAll('#quickActions .qa-chip:not(.qa-more)');
  chips[0].click();
  chips[2].click();
  r.filled = document.getElementById('composerInput').value;
  chips[1].click();
  r.afterSend = document.getElementById('composerInput').value;

  o.loadConversation([{role:'user', text:'Hello'}, {role:'assistant', text:'Hi there'}]);
  r.loaded = texts('#messages .msg .msg-content');

  o.setComposeMode(true);
  o.beginDraft('d1', 'Reply draft');
  o.appendTextDelta('d1', 'Dear Sara, ' + hostile);
  o.finalizeDraft('d1', [{label:'Insert', action:'insert', primary:true}, {label:'Discard', action:'discard'}], {status:'not in the email yet', note:'Fill the brackets.'});
  r.draftText = document.querySelector('[data-message-id=d1] .msg-content').textContent.trim();
  r.draftParsed = document.querySelectorAll('[data-message-id=d1] img, [data-message-id=d1] b').length;
  r.draftButtons = texts('[data-message-id=d1] .draft-foot button');
  r.draftNote = texts('[data-message-id=d1] .draft-note');
  document.querySelector('[data-message-id=d1] .draft-foot button').click();
  o.markDraft('d1', 'in the email', true);
  r.draftStatus = texts('[data-message-id=d1] .draft-status');
  r.usedCardButtons = texts('[data-message-id=d1] .draft-foot button');
  document.querySelector('[data-message-id=d1] .draft-foot button').click();
  r.draftGone = document.querySelector('[data-message-id=d1]') === null;

  r.sendLabel = document.getElementById('btnSend').textContent;
  r.draftButtonHidden = document.getElementById('btnDraft').hidden;
  document.getElementById('composerInput').value = 'say yes';
  document.getElementById('btnDraft').click();
  o.appendAssistantMessage('a9', 'Answer text');
  o.finalizeAssistantMessage('a9', {});
  document.querySelector('[data-message-id=a9] .msg-foot button').click();
  o.appendToolCallCard('t1', 'jc_remember', '{}');
  o.updateToolCallCard('t1', true, 'ok', JSON.stringify({result_type:'memory_proposal', note: hostile}));

  r.proposal = texts('.msg-draft .draft-title') + ' / ' + document.querySelector('.msg-draft:last-of-type .msg-content').textContent;
  r.proposalParsed = document.querySelectorAll('.msg-draft:last-of-type img, .msg-draft:last-of-type b').length;
  document.querySelector('.msg-draft:last-of-type .draft-foot button').click();
  // A meeting proposal during an answer lands inside that answer, after its text.
  o.appendAssistantMessage('m5', 'Here are the times.');
  o.appendToolCallCard('t5', 'jc_propose_meeting', '{}');
  o.updateToolCallCard('t5', true, 'ok', JSON.stringify({result_type:'meeting_proposal', subject:'Budget', attendees:['sara@example.com'], agenda:'', slots:[{start:'2026-11-16T10:00', end:'2026-11-16T10:30'}]}));
  o.appendToolCallCard('t6', 'outlook_read_message', '{}');
  o.updateToolCallCard('t6', false, '{""error"":{""code"":""not_found"",""message"":""Message X not found""}}', '');
  var errs = document.querySelectorAll('.tool-status-err'); r.toolError = errs[errs.length - 1].textContent;
  o.finalizeAssistantMessage('m5', {});
  r.meetingInside = document.querySelectorAll('[data-message-id=m5] .msg-meeting').length;
  r.meetingSlots = texts('[data-message-id=m5] .msg-meeting .draft-foot button');

  // The emails an answer used are listed under it; a row opens its email.
  o.appendAssistantMessage('a20', 'Two emails.');
  o.appendToolCallCard('t2', 'outlook_search_messages', '{}');
  o.updateToolCallCard('t2', true, '2 found', JSON.stringify({messages: [
    {id:'k1', subject: hostile, from:'Sara Example', received_at:'2026-11-12T09:00:00Z'},
    {id:'k1', subject: 'again', from:'Sara Example', received_at:'2026-11-12T09:00:00Z'},
    {id:'k2', subject: 'Budget', from:'Omar Example', received_at:'2026-11-10T09:00:00Z'}]}));
  o.appendToolCallCard('t3', 'outlook_read_message', '{}');
  o.updateToolCallCard('t3', true, 'read', JSON.stringify({id:'k3', subject:'Read one', from:'Lina', received_at:'bad date'}));
  o.finalizeAssistantMessage('a20', {});
  r.sources = texts('[data-message-id=a20] .source-subject');
  r.sourcesParsed = document.querySelectorAll('[data-message-id=a20] .msg-sources img, [data-message-id=a20] .msg-sources b').length;
  document.querySelector('[data-message-id=a20] .source-row').click();
  o.appendUserMessage('next question');
  o.appendAssistantMessage('a21', 'No tools this time.');
  o.finalizeAssistantMessage('a21', {});
  r.noSources = document.querySelectorAll('[data-message-id=a21] .msg-sources').length;

  // Prioritize my inbox: rows open the email, Reply asks for a reply card, For information starts folded.
  o.appendAssistantMessage('p1', 'Reading...');
  o.showPriorities('p1', [
    {level:'high', label:'Needs you', items:[{id:'h1', from:'Sara', subject: hostile, when:'2026-11-13T09:00:00Z', reason: hostile}]},
    {level:'medium', label:'Can wait', items:[]},
    {level:'low', label:'For information', items:[{id:'l1', from:'News', subject:'Weekly', when:'2026-11-12T09:00:00Z', reason:''}]}]);
  r.prioHeads = texts('[data-message-id=p1] .prio-head');
  r.prioReason = document.querySelector('[data-message-id=p1] .prio-reason').textContent;
  r.prioParsed = document.querySelectorAll('[data-message-id=p1] img, [data-message-id=p1] b').length;
  r.lowFolded = document.querySelector('[data-message-id=p1] .prio-low .prio-list').hidden;
  r.lowReply = document.querySelectorAll('[data-message-id=p1] .prio-low .prio-reply').length;
  document.querySelector('[data-message-id=p1] .prio-high .prio-open').click();
  document.querySelector('[data-message-id=p1] .prio-high .prio-reply').click();
  o.replaceAssistantText('a21', hostile);
  r.replaced = document.querySelector('[data-message-id=a21] .msg-content').textContent;

  // Review: the summary, ticked points, Apply sends only the ticked ones.
  o.appendAssistantMessage('rv1', 'Reviewing...');
  o.showReview('rv1', hostile, [{id:'1', issue:'Unanswered', suggestion: hostile}, {id:'2', issue:'', suggestion:'Add a date.'}]);
  r.reviewSummary = document.querySelector('[data-message-id=rv1] .review-summary').textContent;
  r.reviewParsed = document.querySelectorAll('[data-message-id=rv1] img, [data-message-id=rv1] b').length;
  var boxes = document.querySelectorAll('[data-message-id=rv1] .review-point input');
  r.reviewChecked = [].filter.call(boxes, function(b){ return b.checked; }).length;
  boxes[0].checked = false;
  document.querySelector('[data-message-id=rv1] .draft-foot .btn-primary').click();
  o.appendAssistantMessage('rv2', 'Reviewing...');
  o.showReview('rv2', 'Fine.', []);
  r.reviewOk = document.querySelector('[data-message-id=rv2] .review-ok').textContent;

  // Tidy: ticked changes, Apply sends the ticked ids, Undo appears once applied.
  o.appendAssistantMessage('t9', 'Reading...');
  o.showTidy('t9', [{action:'flag', label:'Flag for follow-up', items:[{id:'f1', from:'Sara', subject: hostile, when:'', reason:'Asks.'}]},
                    {action:'mark_read', label:'Mark as read', items:[{id:'m1', from:'News', subject:'Weekly', when:'', reason:''}]}]);
  r.tidyHeads = texts('[data-message-id=t9] .prio-head');
  r.tidyParsed = document.querySelectorAll('[data-message-id=t9] img, [data-message-id=t9] b').length;
  var tb = document.querySelectorAll('[data-message-id=t9] .review-point input');
  tb[1].checked = false;
  document.querySelector('[data-message-id=t9] .draft-foot .btn-primary').click();
  o.markTidy('t9', '1 of 1 changed', true);
  r.tidyStatus = document.querySelector('[data-message-id=t9] .draft-note').textContent;
  r.tidyButtons = texts('[data-message-id=t9] .draft-foot button');
  document.querySelector('[data-message-id=t9] .draft-foot button').click();
  o.markTidy('t9', 'undone', false);
  r.tidyAfterUndo = document.querySelectorAll('[data-message-id=t9] .draft-foot button').length;

  document.getElementById('btnNewChat').click();
  r.pane = window.__pane || 0;
  return JSON.stringify(r);
})()";

        [WebView2Fact]
        public void TheOnePane_ShowsChatsMemoryChipsAndDrafts_AsText_AndReportsEachClick()
        {
            var root = Path.Combine(Path.GetTempPath(), "jc-lockdown-test-" + Guid.NewGuid().ToString("N"));
            var messages = new List<string>();
            JObject page = null;
            string compact = null;
            const string hostile = "<img src=x onerror=window.__pane=1><b>bold</b>";

            try
            {
                WithWebView(async web =>
                {
                    await Within(WebView2Bootstrap.InitializeAsync(web, Path.Combine(root, "WebUI"), Path.Combine(root, "WebView2Data")), 60);
                    var core = web.CoreWebView2;
                    var loaded = new TaskCompletionSource<bool>();
                    core.NavigationCompleted += (s, e) => loaded.TrySetResult(e.IsSuccess);
                    core.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                    Assert.True(await Within(loaded.Task, 30), "The chat page did not load.");
                    core.WebMessageReceived += (s, e) => messages.Add(e.TryGetWebMessageAsString());

                    page = JObject.Parse(JsonConvert.DeserializeObject<string>(await core.ExecuteScriptAsync(DriveThePane)));
                    await Task.Delay(500);
                    compact = await core.ExecuteScriptAsync("document.getElementById('quickActions').classList.contains('compact')");
                });
            }
            finally
            {
                try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }

            // Shown as text: nothing a title, a note or a draft holds becomes an element or runs.
            Assert.Equal(2, (int)page["chatRows"]);
            Assert.Equal(hostile, (string)page["chatTitle"]);
            Assert.Equal("Second", (string)page["current"]);
            Assert.Equal(hostile, (string)page["memText"]);
            Assert.Equal("Dear Sara, " + hostile, (string)page["draftText"]);
            Assert.Equal(0, (int)page["chatParsed"] + (int)page["memParsed"] + (int)page["draftParsed"]);
            Assert.Equal(0, (int)page["pane"]);

            // The emails an answer used: one row each, once, shown as text; none when no tool ran.
            Assert.Equal(hostile + "|Budget|Read one", (string)page["sources"]);
            Assert.Equal(0, (int)page["sourcesParsed"]);
            Assert.Equal(0, (int)page["noSources"]);

            // Prioritize my inbox: groups with counts (an empty group is left out), text only, low folded, no Reply on low.
            Assert.Equal("Needs you (1)|For information (1)", (string)page["prioHeads"]);
            Assert.Equal(hostile, (string)page["prioReason"]);
            Assert.Equal(0, (int)page["prioParsed"]);
            Assert.True((bool)page["lowFolded"]);
            Assert.Equal(0, (int)page["lowReply"]);
            Assert.Equal(hostile, (string)page["replaced"]);

            // Review: shown as text, every point ticked at first; a clean draft says so.
            Assert.Equal(hostile, (string)page["reviewSummary"]);
            Assert.Equal(0, (int)page["reviewParsed"]);
            Assert.Equal(2, (int)page["reviewChecked"]);
            Assert.Equal("Nothing to change. It is ready to send.", (string)page["reviewOk"]);

            // The meeting card sits in its answer; a tool's raw error is never shown.
            Assert.Equal(1, (int)page["meetingInside"]);
            Assert.Equal("Mon 16 Nov  10:00 - 10:30", (string)page["meetingSlots"]);
            Assert.Equal("\u26A0 Reading message did not work; Jameel Copilot carried on without it.", (string)page["toolError"]);

            // Tidy: groups with counts, text only; after Apply only Undo; after Undo nothing.
            Assert.Equal("Flag for follow-up (1)|Mark as read (1)", (string)page["tidyHeads"]);
            Assert.Equal(0, (int)page["tidyParsed"]);
            Assert.Equal("1 of 1 changed", (string)page["tidyStatus"]);
            Assert.Equal("Undo", (string)page["tidyButtons"]);
            Assert.Equal(0, (int)page["tidyAfterUndo"]);

            // "Remember this?": the model's proposal is a card, shown as text; Save is the user's.
            Assert.Equal("Remember this? / " + hostile, (string)page["proposal"]);
            Assert.Equal(0, (int)page["proposalParsed"]);

            // Chips: a group title each; "fill" waits for the user, "send" goes.
            Assert.Equal("This email|Reports", (string)page["titles"]);
            Assert.Equal("Summarize my mail with [name].", (string)page["filled"]);
            Assert.Equal("", (string)page["afterSend"]);
            Assert.Equal("true", compact);

            Assert.Equal("Hello|Hi there", ((string)page["loaded"]).Replace("\n", ""));

            // The draft card and the compose window's input.
            Assert.Equal("Insert|Discard", (string)page["draftButtons"]);
            Assert.Equal("Fill the brackets.", (string)page["draftNote"]);
            Assert.Equal("in the email", (string)page["draftStatus"]);
            // A card that was used cannot put its text in a second time.
            Assert.Equal("Dismiss", (string)page["usedCardButtons"]);

            // Deleting takes two clicks: the first one only asks. (Clear all was pressed once: no message below.)
            Assert.Equal("Delete?", (string)page["delAsks"]);
            Assert.Equal("Delete?", (string)page["memDelAsks"]);
            Assert.Equal("Delete all chats?", (string)page["clearAsks"]);
            Assert.True((bool)page["draftGone"]);
            Assert.Equal("Ask", (string)page["sendLabel"]);
            Assert.False((bool)page["draftButtonHidden"]);

            // Every click reaches the host as one message.
            var expected = new[]
            {
                "{\"type\":\"open_chat\",\"payload\":{\"id\":\"c1\"}}",
                "{\"type\":\"delete_chat\",\"payload\":{\"id\":\"c2\"}}",
                "{\"type\":\"memory_remove\",\"payload\":{\"id\":\"m1\"}}",
                "{\"type\":\"memory_add\",\"payload\":{\"text\":\"Keep replies short.\"}}",
                "{\"type\":\"chip_action\",\"payload\":{\"id\":\"reply\"}}",
                "{\"type\":\"send\",\"payload\":{\"text\":\"Summarize it.\"}}",
                "{\"type\":\"draft_action\",\"payload\":{\"id\":\"d1\",\"action\":\"insert\"}}",
                "{\"type\":\"draft_action\",\"payload\":{\"id\":\"d1\",\"action\":\"discard\"}}",
                "{\"type\":\"draft\",\"payload\":{\"text\":\"say yes\"}}",
                "{\"type\":\"insert_text\",\"payload\":{\"text\":\"Answer text\"}}",
                "{\"type\":\"memory_add\",\"payload\":{\"text\":\"<img src=x onerror=window.__pane=1><b>bold</b>\"}}",
                "{\"type\":\"open_message\",\"payload\":{\"id\":\"k1\"}}",
                "{\"type\":\"open_message\",\"payload\":{\"id\":\"h1\"}}",
                "{\"type\":\"priority_reply\",\"payload\":{\"id\":\"h1\"}}",
                "{\"type\":\"review_apply\",\"payload\":{\"id\":\"rv1\",\"picks\":[\"2\"]}}",
                "{\"type\":\"tidy_apply\",\"payload\":{\"id\":\"t9\",\"picks\":[\"f1\"]}}",
                "{\"type\":\"tidy_undo\",\"payload\":{\"id\":\"t9\"}}",
                "{\"type\":\"new_chat\"}",
            };
            Assert.Equal(expected, messages.Where(m => m != null && !m.Contains("\"ready\"")).ToArray());
        }

        // The cards of 1.4, driven the way the host drives them. No double quote in this script.
        private const string DriveTheNewCards = @"
(function () {
  var o = window.jameelcopilot, r = {};
  var HOSTILE = '<img src=x onerror=window.__pane=1><b>bold</b>';
  function texts(sel) { return [].map.call(document.querySelectorAll(sel), function (e) { return e.textContent; }); }

  // The user's own prompts: chips, and a view to add, change and delete them.
  o.setQuickActions([{ title: 'My prompts', chips: [{ label: 'Weekly', prompt: 'Summarize my week.', mode: 'send' }, { label: '+ Add', mode: 'prompts' }] }]);
  document.querySelectorAll('.qa-chip')[1].click();
  r.promptsShown = !document.getElementById('viewPrompts').hidden && document.getElementById('viewChat').hidden;
  o.setPrompts([{ id: 'p1', text: HOSTILE }]);
  r.promptText = texts('#promptList .mem-text')[0];
  document.getElementById('promptInput').value = 'Summarize what [name] sent me.';
  document.getElementById('btnPromptAdd').click();
  document.getElementById('btnPromptsBack').click();
  r.backInChat = !document.getElementById('viewChat').hidden;

  // Start my day.
  o.appendAssistantMessage('day1', 'Reading...');
  o.showDay('day1', {
    date: '2026-11-15',
    todos: [{ text: 'Send the budget', due: '2026-11-17', id: 'e1', from: 'Sara Example', subject: 'Budget' },
            { text: HOSTILE, due: '', id: 'e2', from: 'A', subject: 'B' }],
    topics: [{ title: 'Office move', summary: 'It is on Thursday.', emails: [{ id: 'e3', from: 'Facilities', subject: 'Move', when: '2026-11-14T09:00:00+03:00' }] }],
    meetings: [{ subject: 'Budget review', start: '2026-11-15T10:00', end: '2026-11-15T10:30', all_day: false, location: 'Room 4' }]
  });
  r.dayTitle = texts('[data-message-id=day1] .day-title')[0];
  r.daySections = texts('[data-message-id=day1] .day-section');
  r.todoTexts = texts('[data-message-id=day1] .todo-text');
  r.due = texts('[data-message-id=day1] .todo-due');
  r.meeting = [].map.call(document.querySelector('[data-message-id=day1] .meet-row').children, function (e) { return e.textContent; });
  document.querySelector('[data-message-id=day1] .todo-source').click();
  document.querySelector('[data-message-id=day1] .topic .source-row').click();
  var adds = document.querySelectorAll('[data-message-id=day1] .todo-add');
  adds[0].click();
  o.markTodo('day1', 0, 'added to your Tasks', true);
  adds[1].click();
  o.markTodo('day1', 1, 'could not be added', false);
  r.todoDone = texts('[data-message-id=day1] .todo-done');
  r.todoRetry = texts('[data-message-id=day1] .todo-add');
  r.dayParsed = document.querySelectorAll('[data-message-id=day1] img, [data-message-id=day1] b, #promptList img, #promptList b').length;

  o.appendAssistantMessage('day2', '');
  o.showDay('day2', { date: '2026-11-15', todos: [], topics: [], meetings: null });
  r.dayEmpty = texts('[data-message-id=day2] .day-empty');
  r.dayEmptySections = texts('[data-message-id=day2] .day-section');

  // Action items of an email that asks nothing.
  o.appendAssistantMessage('td1', '');
  o.showTodos('td1', []);
  r.noTodos = texts('[data-message-id=td1] .review-ok')[0];

  // Waiting for a reply: the row's button asks for a follow-up, not a reply.
  o.appendAssistantMessage('w1', '');
  o.showPriorities('w1', [{ level: 'medium', label: 'Waiting for a reply', button: { label: 'Follow up', type: 'followup' },
    items: [{ id: 's1', from: 'To: Sara Example', subject: 'Figures', when: '2026-11-10T09:00:00+03:00', reason: 'You asked for the figures. (sent 5 days ago)' }] }]);
  r.waitButton = texts('[data-message-id=w1] .prio-reply')[0];
  document.querySelector('[data-message-id=w1] .prio-reply').click();

  // A rule proposal: shown as text; Create rule names the card only.
  o.appendToolCallCard('r1', 'jc_propose_rule', '{}');
  o.updateToolCallCard('r1', true, '', JSON.stringify({ result_type: 'rule_proposal', name: 'Vendor invoices', from: ['billing@vendor.example'],
    subject_contains: ['invoice'], folder_id: 'f1', folder: HOSTILE, category: 'Finance' }));
  r.ruleLines = texts('.msg-rule .rule-line');
  r.ruleParsed = document.querySelectorAll('.msg-rule img, .msg-rule b').length;
  document.querySelector('.msg-rule .draft-foot button').click();
  o.markRule('r1', 'created - see Rules in Outlook', true);
  r.ruleStatus = texts('.msg-rule .draft-status')[0];
  r.ruleButtons = document.querySelectorAll('.msg-rule .draft-foot button').length;

  // A meeting proposal names its room.
  o.appendToolCallCard('m1', 'jc_propose_meeting', '{}');
  o.updateToolCallCard('m1', true, '', JSON.stringify({ result_type: 'meeting_proposal', subject: 'Sync', attendees: ['sara@example.com'], room: 'Room 4',
    slots: [{ start: '2026-11-15T10:00', end: '2026-11-15T10:30' }], agenda: '' }));
  r.meetingWho = texts('.msg-meeting .meeting-who');
  r.pane = window.__pane || 0;
  return JSON.stringify(r);
})()";

        [WebView2Fact]
        public void TheCardsOf14_ShowDayTodosWaitingRulesAndPrompts_AsText_AndReportEachClick()
        {
            var root = Path.Combine(Path.GetTempPath(), "jc-lockdown-test-" + Guid.NewGuid().ToString("N"));
            var messages = new List<string>();
            JObject page = null;
            const string hostile = "<img src=x onerror=window.__pane=1><b>bold</b>";

            try
            {
                WithWebView(async web =>
                {
                    await Within(WebView2Bootstrap.InitializeAsync(web, Path.Combine(root, "WebUI"), Path.Combine(root, "WebView2Data")), 60);
                    var core = web.CoreWebView2;
                    var loaded = new TaskCompletionSource<bool>();
                    core.NavigationCompleted += (s, e) => loaded.TrySetResult(e.IsSuccess);
                    core.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                    Assert.True(await Within(loaded.Task, 30), "The chat page did not load.");
                    core.WebMessageReceived += (s, e) => messages.Add(e.TryGetWebMessageAsString());

                    page = JObject.Parse(JsonConvert.DeserializeObject<string>(await core.ExecuteScriptAsync(DriveTheNewCards)));
                    await Task.Delay(500);
                });
            }
            finally
            {
                try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }

            // My prompts: the last chip opens their view; a prompt is shown as text.
            Assert.True((bool)page["promptsShown"]);
            Assert.Equal(hostile, (string)page["promptText"]);
            Assert.True((bool)page["backInChat"]);

            // Start my day: three sections with counts; to-dos with a due date; the calendar's meeting.
            Assert.Equal("Your day - Sun 15 Nov", (string)page["dayTitle"]);
            Assert.Equal(new[] { "To do (2)", "Catch up (1)", "Today's meetings (1)" }, page["daySections"].Select(t => (string)t));
            Assert.Equal(new[] { "Send the budget", hostile }, page["todoTexts"].Select(t => (string)t));
            Assert.Equal(new[] { "Due 17 Nov" }, page["due"].Select(t => (string)t));
            Assert.Equal(new[] { "10:00 - 10:30", "Budget review", "Room 4" }, page["meeting"].Select(t => (string)t));
            Assert.Equal(new[] { "added to your Tasks" }, page["todoDone"].Select(t => (string)t));
            Assert.Equal(new[] { "Try again" }, page["todoRetry"].Select(t => (string)t));
            Assert.Equal(0, (int)page["dayParsed"]);

            // Nothing asked and an unreadable calendar are said in words; no empty Catch up section.
            Assert.Equal(new[] { "Nothing in your recent mail asks something of you.", "Your calendar could not be read." }, page["dayEmpty"].Select(t => (string)t));
            Assert.Equal(new[] { "To do (0)", "Today's meetings" }, page["dayEmptySections"].Select(t => (string)t));
            Assert.Equal("This email asks nothing of you.", (string)page["noTodos"]);

            Assert.Equal("Follow up", (string)page["waitButton"]);

            // A rule: what it does in plain words, the folder's name as text, and no button left once created.
            Assert.Equal(new[]
            {
                "When new mail is frombilling@vendor.example", "and its subject containsinvoice",
                "move it to the folder" + hostile, "and give it the categoryFinance",
            }, page["ruleLines"].Select(t => (string)t));
            Assert.Equal(0, (int)page["ruleParsed"]);
            Assert.Equal("created - see Rules in Outlook", (string)page["ruleStatus"]);
            Assert.Equal(0, (int)page["ruleButtons"]);

            Assert.Equal(new[] { "With: sara@example.com", "Room: Room 4" }, page["meetingWho"].Select(t => (string)t));
            Assert.Equal(0, (int)page["pane"]);

            // Every click reaches the host as one message that names a card or a row, never its text.
            var expected = new[]
            {
                "{\"type\":\"show_prompts\"}",
                "{\"type\":\"prompt_add\",\"payload\":{\"text\":\"Summarize what [name] sent me.\"}}",
                "{\"type\":\"open_message\",\"payload\":{\"id\":\"e1\"}}",
                "{\"type\":\"open_message\",\"payload\":{\"id\":\"e3\"}}",
                "{\"type\":\"task_add\",\"payload\":{\"id\":\"day1\",\"index\":0}}",
                "{\"type\":\"task_add\",\"payload\":{\"id\":\"day1\",\"index\":1}}",
                "{\"type\":\"followup\",\"payload\":{\"id\":\"s1\"}}",
                "{\"type\":\"rule_create\",\"payload\":{\"call\":\"r1\"}}",
            };
            Assert.Equal(expected, messages.Where(m => m != null && !m.Contains("\"ready\"")).ToArray());
        }

        // Runs the body on an STA thread with a message loop, in an off-screen window.
        private static void WithWebView(Func<WebView2, Task> body)
        {
            Sta.Run(() =>
            {
                Exception failure = null;
                using (var form = new Form())
                using (var web = new WebView2())
                {
                    form.ShowInTaskbar = false;
                    form.StartPosition = FormStartPosition.Manual;
                    form.Location = new Point(-32000, -32000);
                    form.Size = new Size(800, 600);
                    web.Dock = DockStyle.Fill;
                    form.Controls.Add(web);
                    form.Load += async (sender, e) =>
                    {
                        try { await body(web); }
                        catch (Exception ex) { failure = ex; }
                        finally { form.Close(); }
                    };
                    Application.Run(form);
                }
                if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            });
        }

        private static async Task Within(Task task, int seconds)
        {
            if (await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(seconds))) != task)
            {
                throw new TimeoutException("WebView2 did not answer within " + seconds + " s.");
            }
            await task;
        }

        private static async Task<T> Within<T>(Task<T> task, int seconds)
        {
            await Within((Task)task, seconds);
            return task.Result;
        }
    }
}
