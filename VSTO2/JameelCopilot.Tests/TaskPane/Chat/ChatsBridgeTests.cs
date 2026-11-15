using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Prefs;
using JameelCopilot.TaskPane.Chat;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.TaskPane.Chat
{
    /// <summary>
    /// A pane keeps its conversation after every turn and has it back when it
    /// opens again; the Chats and Memory views are served from the mailbox.
    /// </summary>
    public class ChatsBridgeTests
    {
        private readonly FakeUserStore _mailbox = new FakeUserStore();
        private readonly List<string> _scripts = new List<string>();
        private DateTimeOffset _now = new DateTimeOffset(2026, 11, 14, 9, 30, 0, TimeSpan.FromHours(3));
        private bool _busy;

        private ChatArchive Archive() => new ChatArchive(_mailbox, () => _now);

        private ChatsBridge Pane(ConversationStore store, string scope = null, string title = null, ChatArchive archive = null)
        {
            return new ChatsBridge(archive ?? Archive(), new UserMemoryService(_mailbox), store,
                script => { _scripts.Add(script); return Task.CompletedTask; }, scope, title, () => _busy, () => _now);
        }

        private static ConversationStore Said(params string[] turns)
        {
            var store = new ConversationStore();
            Say(store, turns);
            return store;
        }

        private static void Say(ConversationStore store, params string[] turns)
        {
            for (var i = 0; i < turns.Length; i++)
            {
                store.Append(new JObject(new JProperty("type", "message"),
                    new JProperty("role", i % 2 == 0 ? "user" : "assistant"), new JProperty("content", turns[i])));
            }
        }

        private static JObject Payload(string name, string value) => new JObject(new JProperty(name, value));

        private static string[] Texts(ConversationStore store) => store.Snapshot().Select(m => (string)m["content"]).ToArray();

        [Fact]
        public async Task ThreadPane_GetsItsChatBack_WithItsHistory()
        {
            var first = Said("What does Sara ask for?", "The budget file.");
            Pane(first, "thread:CONV1", "RE: Budget review").Save();

            // The reply window is closed; later a new one opens on the same thread.
            _now = _now.AddDays(3);
            var reopened = new ConversationStore();
            await Pane(reopened, "thread:CONV1", "RE: Budget review").OpenAsync();

            Assert.Equal(new[] { "What does Sara ask for?", "The budget file." }, Texts(reopened));
            Assert.Equal(
                "jameelcopilot.loadConversation([{\"role\":\"user\",\"text\":\"What does Sara ask for?\"},{\"role\":\"assistant\",\"text\":\"The budget file.\"}]);",
                Assert.Single(_scripts));
            Assert.Equal("RE: Budget review", Assert.Single(Archive().List()).Title);
        }

        [Fact]
        public async Task MailboxPane_ReopensTodaysChat_NotAnOlderOne()
        {
            Pane(Said("What needs my attention?", "Three messages.")).Save();

            var sameDay = new ConversationStore();
            _now = _now.AddHours(5);
            await Pane(sameDay).OpenAsync();
            Assert.Equal(2, sameDay.Count);

            var nextDay = new ConversationStore();
            _now = _now.AddDays(1);
            _scripts.Clear();
            await Pane(nextDay).OpenAsync();
            Assert.Equal(0, nextDay.Count);
            Assert.Empty(_scripts);
        }

        [Fact]
        public async Task NewEmailWindow_OpensEmpty_AndAnotherThreadDoesNotGetThisChat()
        {
            Pane(Said("q", "a"), "thread:CONV1").Save();
            Pane(Said("draft me a note", "done"), ChatsBridge.ComposeScope).Save();

            var otherThread = new ConversationStore();
            var newEmail = new ConversationStore();
            await Pane(otherThread, "thread:CONV2").OpenAsync();
            await Pane(newEmail, ChatsBridge.ComposeScope).OpenAsync();

            Assert.Equal(0, otherThread.Count);
            Assert.Equal(0, newEmail.Count);
        }

        [Fact]
        public void Save_KeepsOneChatPerConversation_AndNewChatStartsAnother()
        {
            var store = Said("first question", "first answer");
            var pane = Pane(store);

            pane.Save();
            Say(store, "second question", "second answer");
            pane.Save();
            Assert.Equal(4, Assert.Single(Archive().List()).Count);

            Assert.True(pane.Handle("new_chat", null));
            Assert.Equal(0, store.Count);
            Assert.Equal("jameelcopilot.clear();", _scripts.Last());

            Say(store, "another topic", "another answer");
            pane.Save();
            Assert.Equal(new[] { "another topic", "first question" }, Archive().List().Select(c => c.Title));
        }

        [Fact]
        public void Save_WithNothingSaid_KeepsNothing()
        {
            Pane(new ConversationStore()).Save();

            Assert.Empty(Archive().List());
        }

        // Review Focus 2: the Inbox pane and a reply window are open together.
        [Fact]
        public void TwoPanes_OnOneArchive_KeepTheirOwnChats()
        {
            var archive = Archive();
            var inbox = Said("inbox question", "inbox answer");
            var reply = Said("reply question", "reply answer");
            var inboxPane = Pane(inbox, null, null, archive);
            var replyPane = Pane(reply, "thread:CONV1", "RE: Budget", archive);

            inboxPane.Save();
            replyPane.Save();
            Say(inbox, "inbox again", "inbox more");
            inboxPane.Save();

            var chats = Archive().List();
            Assert.Equal(2, chats.Count);
            Assert.Equal(4, chats.Single(c => c.Scope == ChatArchive.MailboxScope).Count);
            Assert.Equal(2, chats.Single(c => c.Scope == "thread:CONV1").Count);
        }

        [Fact]
        public void OpenChat_ShowsIt_AndGoingOnSavesIntoThatChat()
        {
            var thread = Said("thread question", "thread answer");
            Pane(thread, "thread:CONV1", "RE: Budget").Save();
            var threadChat = Archive().List().Single();

            // Opened from the Chats list in the Inbox pane, then continued there.
            var inbox = new ConversationStore();
            var pane = Pane(inbox);
            pane.Handle("open_chat", Payload("id", threadChat.Id));
            Assert.Equal(2, inbox.Count);
            Assert.StartsWith("jameelcopilot.loadConversation(", _scripts.Last());

            Say(inbox, "one more", "sure");
            pane.Save();

            var saved = Assert.Single(Archive().List());
            Assert.Equal(threadChat.Id, saved.Id);
            Assert.Equal("thread:CONV1", saved.Scope);
            Assert.Equal("RE: Budget", saved.Title);
            Assert.Equal(4, saved.Count);
        }

        [Fact]
        public void ShowChats_ListsThemNewestFirst_WithTheOpenOneMarked()
        {
            Pane(Said("older", "a"), "thread:CONV1", "RE: Budget").Save();
            _now = _now.AddMinutes(10);
            var pane = Pane(Said("What needs my attention?", "b"));
            pane.Save();
            var newest = Archive().List()[0];

            pane.Handle("show_chats", null);

            var script = _scripts.Last();
            Assert.StartsWith("jameelcopilot.setChats([{\"id\":\"" + newest.Id + "\",\"title\":\"What needs my attention?\",\"thread\":false,\"when\":\"Today 09:40\"},", script);
            Assert.Contains("\"title\":\"RE: Budget\",\"thread\":true,\"when\":\"Today 09:30\"}", script);
            Assert.EndsWith("], \"" + newest.Id + "\");", script);
        }

        [Fact]
        public void DeleteTheOpenChat_EmptiesThePane_AndClearAllEmptiesTheList()
        {
            var store = Said("question", "answer");
            var pane = Pane(store);
            pane.Save();
            var id = Archive().List().Single().Id;

            pane.Handle("delete_chat", Payload("id", id));

            Assert.Equal(0, store.Count);
            Assert.Empty(Archive().List());
            Assert.Contains("jameelcopilot.clear();", _scripts);
            Assert.Equal("jameelcopilot.setChats([], \"\");", _scripts.Last());

            Say(store, "again", "ok");
            pane.Save();
            pane.Handle("clear_chats", null);
            Assert.Empty(Archive().List());
            Assert.Equal(0, store.Count);
        }

        [Fact]
        public void OpenChat_ThatIsGone_RefreshesTheList()
        {
            var store = Said("mine", "kept");
            var pane = Pane(store);

            pane.Handle("open_chat", Payload("id", "no-such-chat"));
            pane.Handle("open_chat", null);

            Assert.Equal(2, store.Count);   // what is on screen stays
            Assert.All(_scripts, s => Assert.StartsWith("jameelcopilot.setChats(", s));
        }

        // A turn appends to the conversation when it ends: it must not be swapped meanwhile.
        [Fact]
        public void WhileATurnRuns_TheConversationIsNotSwapped()
        {
            var store = Said("question", "answer");
            var pane = Pane(store);
            pane.Save();
            var id = Archive().List().Single().Id;
            _busy = true;

            Assert.True(pane.Handle("new_chat", null));
            Assert.True(pane.Handle("open_chat", Payload("id", id)));
            Assert.True(pane.Handle("delete_chat", Payload("id", id)));
            Assert.True(pane.Handle("clear_chats", null));

            Assert.Equal(2, store.Count);
            Assert.Single(Archive().List());
            Assert.Empty(_scripts);
        }

        [Fact]
        public void Memory_IsAddedChangedAndRemovedFromThePage_AndGoesIntoThePrompt()
        {
            var pane = Pane(new ConversationStore());

            pane.Handle("memory_add", Payload("text", "Reply to HR in Arabic."));
            var note = new UserMemoryService(_mailbox).Notes.Single();
            Assert.Equal("jameelcopilot.setMemory([{\"id\":\"" + note.Id + "\",\"text\":\"Reply to HR in Arabic.\"}]);", _scripts.Last());
            Assert.Equal(UserMemoryService.PromptHeader + "\n- Reply to HR in Arabic.", pane.Memory);

            pane.Handle("memory_update", new JObject(new JProperty("id", note.Id), new JProperty("text", "Reply to HR in English.")));
            Assert.Contains("Reply to HR in English.", _scripts.Last());

            pane.Handle("memory_remove", Payload("id", note.Id));
            Assert.Equal("jameelcopilot.setMemory([]);", _scripts.Last());
            Assert.Equal("", pane.Memory);

            pane.Handle("show_memory", null);
            pane.Handle("memory_add", new JObject(new JProperty("text", 5)));
            Assert.Equal("jameelcopilot.setMemory([]);", _scripts.Last());
        }

        [Fact]
        public void OtherMessages_AreNotTheBridges()
        {
            var pane = Pane(new ConversationStore());

            Assert.False(pane.Handle("send", Payload("text", "hello")));
            Assert.False(pane.Handle("set_style", null));
            Assert.Empty(_scripts);
        }

        [Fact]
        public void NoArchiveAndNoMemory_StillWorkForTheSession()
        {
            var store = Said("question", "answer");
            var pane = new ChatsBridge(null, null, store, script => { _scripts.Add(script); return Task.CompletedTask; });

            pane.Save();
            pane.Handle("show_chats", null);
            pane.Handle("memory_add", Payload("text", "Keep it short."));

            Assert.Contains("\"title\":\"question\"", _scripts[0]);
            Assert.Contains("Keep it short.", pane.Memory);
        }

        [Theory]
        [InlineData("2026-11-14T09:18:00+03:00", "Today 09:18")]
        [InlineData("2026-11-13T17:40:00+03:00", "Yesterday 17:40")]
        [InlineData("2026-11-13T23:30:00+00:00", "Today 02:30")]      // another zone's evening is this morning here
        [InlineData("2026-11-10T08:31:00+03:00", "10 Nov 08:31")]
        [InlineData("2025-12-30T08:05:00+03:00", "30 Dec 2025 08:05")]
        public void When_IsShortAndInTheUsersTime(string updated, string expected)
        {
            Assert.Equal(expected, ChatsBridge.When(DateTimeOffset.Parse(updated, System.Globalization.CultureInfo.InvariantCulture), _now));
        }
        // The user's own prompts: chips, kept apart from the memory, and the pane shows its chips again after a change.
        [Fact]
        public void MyPrompts_AreChips_AndChangingThemShowsTheChipsAgain()
        {
            var changed = 0;
            var prompts = new UserMemoryService(_mailbox, "Prompts");
            var pane = new ChatsBridge(Archive(), new UserMemoryService(_mailbox), new ConversationStore(),
                script => { _scripts.Add(script); return Task.CompletedTask; }, null, null, () => _busy, () => _now, prompts, () => changed++);

            Assert.Equal("+ Add your own prompt", (string)pane.PromptsGroup()["chips"].Single()["label"]);

            Assert.True(pane.Handle("prompt_add", new JObject(new JProperty("text", "Summarize what [name] sent me this week, please."))));
            Assert.True(pane.Handle("prompt_add", new JObject(new JProperty("text", "My week."))));

            var chips = (JArray)pane.PromptsGroup()["chips"];
            Assert.Equal("My prompts", (string)pane.PromptsGroup()["title"]);
            Assert.Equal(new[] { "Summarize what [name] sent...", "My week.", "+ Add" }, chips.Select(c => (string)c["label"]));
            Assert.Equal(new[] { "fill", "send", "prompts" }, chips.Select(c => (string)c["mode"]));
            Assert.Equal("Summarize what [name] sent me this week, please.", (string)chips[0]["prompt"]);
            Assert.Equal(2, changed);
            Assert.StartsWith("jameelcopilot.setPrompts([", _scripts.Last());

            Assert.True(pane.Handle("prompt_remove", new JObject(new JProperty("id", prompts.Notes[1].Id))));
            Assert.True(pane.Handle("show_prompts", null));
            Assert.Equal(3, changed);                                            // showing them changes nothing
            Assert.Equal(2, ((JArray)pane.PromptsGroup()["chips"]).Count);
            Assert.Equal("", new UserMemoryService(_mailbox).PromptBlock());   // a prompt is not a memory note: it is not sent with every question
        }
    }
}
