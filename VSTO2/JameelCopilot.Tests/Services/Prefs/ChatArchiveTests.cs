using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Prefs;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services.Prefs
{
    /// <summary>
    /// The owner's report: "chats are not saved". A chat is kept in the
    /// user's own mailbox and is there again after Outlook restarts.
    /// </summary>
    public class ChatArchiveTests
    {
        private readonly FakeUserStore _store = new FakeUserStore();
        private DateTimeOffset _now = new DateTimeOffset(2026, 11, 14, 9, 0, 0, TimeSpan.FromHours(3));

        private ChatArchive NewArchive() => new ChatArchive(_store, () => _now);

        private static SavedChat Chat(string scope, params string[] turns)
        {
            var chat = new SavedChat { Scope = scope };
            for (var i = 0; i < turns.Length; i++)
            {
                chat.Messages.Add(new SavedMessage { Role = i % 2 == 0 ? "user" : "assistant", Text = turns[i] });
            }
            return chat;
        }

        [Fact]
        public void SavedChat_IsThereAfterARestart()
        {
            var chat = Chat(ChatArchive.MailboxScope, "What needs my attention?", "Three messages need a reply.");
            NewArchive().Save(chat);

            var afterRestart = NewArchive();   // a new instance reads the store, as a new Outlook session does
            var summary = Assert.Single(afterRestart.List());
            Assert.Equal("What needs my attention?", summary.Title);
            Assert.Equal(2, summary.Count);
            Assert.Equal(_now, summary.Updated);

            var loaded = afterRestart.Load(summary.Id);
            Assert.Equal(new[] { "user", "assistant" }, loaded.Messages.Select(m => m.Role));
            Assert.Equal("Three messages need a reply.", loaded.Messages[1].Text);
        }

        // Json.NET reads text that looks like a date as a date. A message is text whatever it looks like.
        [Fact]
        public void MessageThatIsOnlyADate_IsStillThereAfterARestart()
        {
            var chat = Chat(ChatArchive.MailboxScope, "When is the meeting?", "2026-11-18T10:00:00");
            NewArchive().Save(chat);

            Assert.Equal("2026-11-18T10:00:00", NewArchive().Load(chat.Id).Messages[1].Text);
        }

        [Fact]
        public void Save_Again_UpdatesTheSameChat_AndPutsItFirst()
        {
            var archive = NewArchive();
            var first = Chat(ChatArchive.MailboxScope, "one", "1");
            var second = Chat(ChatArchive.MailboxScope, "two", "2");
            archive.Save(first);
            _now = _now.AddMinutes(5);
            archive.Save(second);
            _now = _now.AddMinutes(5);

            first.Messages.Add(new SavedMessage { Role = "user", Text = "more" });
            first.Messages.Add(new SavedMessage { Role = "assistant", Text = "answer" });
            archive.Save(first);

            Assert.Equal(new[] { first.Id, second.Id }, NewArchive().List().Select(c => c.Id));
            Assert.Equal(4, NewArchive().Load(first.Id).Messages.Count);
        }

        // Two panes open at once each save their own chat; neither overwrites the other.
        [Fact]
        public void TwoChats_KeepTheirOwnMessages()
        {
            var archive = NewArchive();
            var mailbox = Chat(ChatArchive.MailboxScope, "inbox question", "inbox answer");
            var thread = Chat(ChatArchive.ThreadScope("CONV1"), "thread question", "thread answer");
            archive.Save(mailbox);
            archive.Save(thread);

            var again = NewArchive();
            Assert.Equal("inbox answer", again.Load(mailbox.Id).Messages[1].Text);
            Assert.Equal("thread answer", again.Load(thread.Id).Messages[1].Text);
            Assert.Equal(thread.Id, again.Latest("thread:CONV1").Id);
            Assert.Equal(mailbox.Id, again.Latest(ChatArchive.MailboxScope).Id);
            Assert.Null(again.Latest("thread:OTHER"));
        }

        // One mailbox open in two Outlook sessions (the PC and a laptop): each
        // read the index when it started. Neither may write its old copy over
        // what the other has saved since.
        [Fact]
        public void TwoSessionsOnOneMailbox_KeepEachOthersChats()
        {
            var pc = NewArchive();
            var laptop = NewArchive();
            Assert.Empty(pc.List());
            Assert.Empty(laptop.List());
            var onPc = Chat(ChatArchive.MailboxScope, "asked on the PC", "a");
            var onLaptop = Chat(ChatArchive.MailboxScope, "asked on the laptop", "b");

            pc.Save(onPc);
            _now = _now.AddMinutes(1);
            laptop.Save(onLaptop);

            Assert.Equal(new[] { onLaptop.Id, onPc.Id }, NewArchive().List().Select(c => c.Id));
            Assert.Equal(new[] { onLaptop.Id, onPc.Id }, pc.List().Select(c => c.Id));
            Assert.Equal("a", laptop.Load(onPc.Id).Messages[1].Text);

            laptop.Delete(onPc.Id);
            Assert.Equal(onLaptop.Id, Assert.Single(pc.List()).Id);
            Assert.Equal(onLaptop.Id, Assert.Single(NewArchive().List()).Id);
        }

        // A read that fails says nothing about what is stored. Writing the list
        // after it would replace every saved chat with the one just made.
        [Fact]
        public void AReadThatFails_NeverReplacesTheStoredList()
        {
            var earlier = NewArchive();
            var old1 = Chat(ChatArchive.MailboxScope, "old one", "1");
            var old2 = Chat(ChatArchive.MailboxScope, "old two", "2");
            earlier.Save(old1);
            _now = _now.AddMinutes(1);
            earlier.Save(old2);

            var archive = NewArchive();
            _store.FailNextLoads = 1;
            _now = _now.AddMinutes(1);
            var fresh = Chat(ChatArchive.MailboxScope, "new", "n");
            archive.Save(fresh);

            Assert.Equal(new[] { old2.Id, old1.Id }, NewArchive().List().Select(c => c.Id));   // nothing was replaced
            Assert.Equal(new[] { fresh.Id, old2.Id, old1.Id }, archive.List().Select(c => c.Id));   // this session has all three

            // The next save reaches the mailbox with the chat that was waiting.
            _now = _now.AddMinutes(1);
            fresh.Messages.Add(new SavedMessage { Role = "user", Text = "more" });
            fresh.Messages.Add(new SavedMessage { Role = "assistant", Text = "ok" });
            archive.Save(fresh);

            Assert.Equal(new[] { fresh.Id, old2.Id, old1.Id }, NewArchive().List().Select(c => c.Id));
            Assert.Equal(4, NewArchive().Load(fresh.Id).Messages.Count);
        }

        [Fact]
        public void ADeleteWhileTheMailboxCannotBeRead_IsDoneAtTheNextWrite()
        {
            var earlier = NewArchive();
            var keep = Chat(ChatArchive.MailboxScope, "keep", "k");
            var drop = Chat(ChatArchive.MailboxScope, "drop", "d");
            earlier.Save(keep);
            earlier.Save(drop);

            var archive = NewArchive();
            Assert.Equal(2, archive.List().Count);
            _store.FailNextLoads = 1;
            archive.Delete(drop.Id);

            Assert.DoesNotContain(archive.List(), c => c.Id == drop.Id);
            _now = _now.AddMinutes(1);
            archive.Save(Chat(ChatArchive.MailboxScope, "next", "n"));
            Assert.Equal(new[] { "next", "keep" }, NewArchive().List().Select(c => c.Title));
        }

        [Fact]
        public void ThreadScope_WithoutAConversation_IsTheMailbox()
        {
            Assert.Equal("thread:ABC", ChatArchive.ThreadScope(" ABC "));
            Assert.Equal(ChatArchive.MailboxScope, ChatArchive.ThreadScope(null));
            Assert.Equal(ChatArchive.MailboxScope, ChatArchive.ThreadScope("  "));
        }

        [Fact]
        public void EmptyChat_IsNotKept()
        {
            var archive = NewArchive();

            archive.Save(new SavedChat { Scope = ChatArchive.MailboxScope });
            archive.Save(Chat(ChatArchive.MailboxScope, "   ", ""));
            archive.Save(null);

            Assert.Empty(archive.List());
            Assert.Empty(_store.Items);
        }

        [Fact]
        public void Delete_RemovesTheChatAndItsItem()
        {
            var archive = NewArchive();
            var keep = Chat(ChatArchive.MailboxScope, "keep", "k");
            var drop = Chat(ChatArchive.MailboxScope, "drop", "d");
            archive.Save(keep);
            archive.Save(drop);

            archive.Delete(drop.Id);
            archive.Delete("no-such-chat");

            Assert.Equal(keep.Id, Assert.Single(NewArchive().List()).Id);
            Assert.Null(NewArchive().Load(drop.Id));
            Assert.DoesNotContain("Chat." + drop.Id, _store.Items.Keys);
        }

        [Fact]
        public void Clear_LeavesOnlyAnEmptyIndex()
        {
            var archive = NewArchive();
            archive.Save(Chat(ChatArchive.MailboxScope, "a", "1"));
            archive.Save(Chat(ChatArchive.MailboxScope, "b", "2"));

            archive.Clear();

            Assert.Empty(NewArchive().List());
            Assert.Equal(new[] { "Chats" }, _store.Items.Keys);
        }

        [Fact]
        public void MoreThanTheLimit_DropsTheOldestChatAndItsItem()
        {
            var archive = NewArchive();
            var chats = new List<SavedChat>();
            for (var i = 0; i < ChatArchive.MaxChats + 2; i++)
            {
                _now = _now.AddMinutes(1);
                var chat = Chat(ChatArchive.MailboxScope, "question " + i, "answer " + i);
                archive.Save(chat);
                chats.Add(chat);
            }

            var list = NewArchive().List();
            Assert.Equal(ChatArchive.MaxChats, list.Count);
            Assert.Equal(chats.Last().Id, list[0].Id);
            Assert.DoesNotContain(list, c => c.Id == chats[0].Id || c.Id == chats[1].Id);
            Assert.Equal(ChatArchive.MaxChats + 1, _store.Items.Count);   // the chats and the index
        }

        [Fact]
        public void LongChat_KeepsTheLastMessages_AndCutsAVeryLongOne()
        {
            var chat = new SavedChat { Scope = ChatArchive.MailboxScope };
            for (var i = 0; i < ChatArchive.MaxMessages + 10; i++)
            {
                chat.Messages.Add(new SavedMessage { Role = i % 2 == 0 ? "user" : "assistant", Text = "m" + i });
            }
            chat.Messages.Add(new SavedMessage { Role = "assistant", Text = new string('x', ChatArchive.MaxMessageChars + 500) });
            var archive = NewArchive();

            archive.Save(chat);

            var loaded = NewArchive().Load(chat.Id);
            Assert.Equal(ChatArchive.MaxMessages, loaded.Messages.Count);
            Assert.Equal(ChatArchive.MaxMessageChars, loaded.Messages.Last().Text.Length);
            Assert.Equal("m11", loaded.Messages[0].Text);
        }

        [Theory]
        [InlineData("  Summarize\r\n this   thread ", "Summarize this thread")]
        [InlineData("", "Chat")]
        [InlineData(null, "Chat")]
        public void TitleFrom_IsOneShortLine(string text, string expected)
        {
            Assert.Equal(expected, ChatArchive.TitleFrom(text));
            Assert.Equal(ChatArchive.MaxTitleChars, ChatArchive.TitleFrom(new string('a', 200)).Length);
        }

        // Only what was said is kept: no tool call, no tool result (that is mail content).
        [Fact]
        public void FromHistory_KeepsWhatWasSaid_NotWhatTheToolsRead()
        {
            var history = new[]
            {
                new JObject(new JProperty("type", "message"), new JProperty("role", "user"), new JProperty("content", "How many unread?")),
                new JObject(new JProperty("type", "function_call"), new JProperty("name", "outlook_count_messages"), new JProperty("arguments", "{}")),
                new JObject(new JProperty("type", "function_call_output"), new JProperty("output", "{\"secret mail text\":1}")),
                new JObject(new JProperty("type", "message"), new JProperty("role", "assistant"), new JProperty("content", "12 unread.")),
                new JObject(new JProperty("type", "message"), new JProperty("role", "assistant"), new JProperty("content", "  ")),
                new JObject(new JProperty("type", "message"), new JProperty("role", "system"), new JProperty("content", "rules")),
            };

            var saved = ChatArchive.FromHistory(history);

            Assert.Equal(new[] { "user:How many unread?", "assistant:12 unread." }, saved.Select(m => m.Role + ":" + m.Text));
        }

        [Fact]
        public void ToHistory_GivesTheNextQuestionItsContext()
        {
            var history = ChatArchive.ToHistory(Chat(ChatArchive.MailboxScope, "How many unread?", "12 unread."));

            Assert.Equal(2, history.Count);
            Assert.Equal("message", (string)history[0]["type"]);
            Assert.Equal("user", (string)history[0]["role"]);
            Assert.Equal("12 unread.", (string)history[1]["content"]);
            Assert.Empty(ChatArchive.ToHistory(null));
        }

        // The mailbox store can be unavailable (and a faulty one could throw): the chat still holds for the session.
        [Theory]
        [InlineData(true, false)]
        [InlineData(false, true)]
        public void StoreUnavailable_ChatsHoldForTheSession(bool unavailable, bool throws)
        {
            _store.Unavailable = unavailable;
            _store.Throws = throws;
            var archive = NewArchive();
            var chat = Chat(ChatArchive.MailboxScope, "question", "answer");

            archive.Save(chat);

            Assert.Equal(chat.Id, Assert.Single(archive.List()).Id);
            Assert.Equal("answer", archive.Load(chat.Id).Messages[1].Text);
            archive.Delete(chat.Id);
            archive.Clear();
            Assert.Empty(archive.List());
        }

        [Fact]
        public void NoStore_AndDamagedItems_AreAnEmptyArchive()
        {
            Assert.Empty(new ChatArchive(null).List());

            _store.Items["Chats"] = "{not json";
            Assert.Empty(NewArchive().List());

            _store.Items["Chats"] = "{\"chats\":[{\"id\":\"a1\",\"title\":\"t\"},{\"id\":5},\"x\",{\"id\":\"a1\"}]}";
            _store.Items["Chat.a1"] = "[]";
            var archive = NewArchive();
            Assert.Equal("a1", Assert.Single(archive.List()).Id);
            Assert.Null(archive.Load("a1"));
            Assert.Null(archive.Load(null));
        }
    }
}
