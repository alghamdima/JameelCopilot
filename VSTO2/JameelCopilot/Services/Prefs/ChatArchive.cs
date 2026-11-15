using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Prefs
{
    /// <summary>
    /// Reads a stored item. Json.NET would otherwise turn any text that looks
    /// like a date into a date, and a message or note that is only a date
    /// would no longer be text.
    /// </summary>
    internal static class StoredJson
    {
        /// <summary>The object, or null when the text holds something else. Throws <see cref="JsonException"/> on damaged text.</summary>
        public static JObject Parse(string json)
        {
            using (var reader = new JsonTextReader(new StringReader(json ?? "")) { DateParseHandling = DateParseHandling.None })
            {
                return JToken.ReadFrom(reader) as JObject;
            }
        }
    }

    public sealed class SavedMessage
    {
        /// <summary>"user" or "assistant".</summary>
        public string Role { get; set; }
        public string Text { get; set; }
    }

    public sealed class ChatSummary
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Scope { get; set; }
        public DateTimeOffset Updated { get; set; }
        public int Count { get; set; }
    }

    public sealed class SavedChat
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Scope { get; set; }
        public DateTimeOffset Updated { get; set; }
        public List<SavedMessage> Messages { get; set; } = new List<SavedMessage>();
    }

    /// <summary>
    /// The user's Jameel Copilot conversations, kept in their own mailbox: an index item
    /// and one item per chat. Only what was said is kept (the user's text and
    /// Jameel Copilot's answers), not the mail Jameel Copilot read on the way. A store that is missing
    /// or fails leaves the chats in memory for the session; nothing here
    /// throws.
    /// <para>The mailbox can be open in two Outlook sessions at once, and a
    /// read can fail. So the index is read again before every change, the
    /// change is applied to what is stored now, and after a read that failed
    /// the index is not written: this session's changes wait for the next
    /// write that follows a good read.</para>
    /// </summary>
    public sealed class ChatArchive
    {
        public const int MaxChats = 40;
        public const int MaxMessages = 80;
        public const int MaxMessageChars = 12000;
        public const int MaxTitleChars = 60;

        /// <summary>The scope of a chat about the mailbox as a whole.</summary>
        public const string MailboxScope = "mailbox";

        /// <summary>The scope of a chat that belongs to one email thread.</summary>
        public static string ThreadScope(string conversationId)
        {
            return string.IsNullOrWhiteSpace(conversationId) ? MailboxScope : "thread:" + conversationId.Trim();
        }

        private const string IndexName = "Chats";
        private const string ChatPrefix = "Chat.";

        private readonly IUserStore _store;
        private readonly Func<DateTimeOffset> _now;
        private readonly object _lock = new object();
        private readonly Dictionary<string, SavedChat> _chats = new Dictionary<string, SavedChat>(StringComparer.Ordinal);

        // What this session knows, newest first.
        private readonly List<ChatSummary> _index = new List<ChatSummary>();

        // Chats saved or deleted here whose change is not in the stored index yet.
        private readonly HashSet<string> _waiting = new HashSet<string>(StringComparer.Ordinal);

        /// <param name="store">Null keeps the chats for this session only.</param>
        public ChatArchive(IUserStore store, Func<DateTimeOffset> now = null)
        {
            _store = store;
            _now = now ?? (() => DateTimeOffset.Now);
        }

        /// <summary>Every saved chat, newest first.</summary>
        public IReadOnlyList<ChatSummary> List()
        {
            lock (_lock)
            {
                Refresh();
                return _index.ToArray();
            }
        }

        /// <summary>The newest chat of that scope, or null.</summary>
        public ChatSummary Latest(string scope)
        {
            lock (_lock)
            {
                Refresh();
                return _index.FirstOrDefault(c => string.Equals(c.Scope, scope, StringComparison.Ordinal));
            }
        }

        /// <summary>The chat, or null when it is not there.</summary>
        public SavedChat Load(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            lock (_lock)
            {
                if (_chats.TryGetValue(id, out var cached)) return cached;
                Refresh();
                if (_index.All(c => c.Id != id)) return null;
                var chat = ParseChat(Read(ChatPrefix + id));
                if (chat != null) _chats[id] = chat;
                return chat;
            }
        }

        /// <summary>Saves the chat as the newest. A chat with nothing said is not kept.</summary>
        public void Save(SavedChat chat)
        {
            if (chat == null || chat.Messages == null) return;
            var messages = chat.Messages
                .Where(m => m != null && !string.IsNullOrWhiteSpace(m.Text) && (m.Role == "user" || m.Role == "assistant"))
                .Select(m => new SavedMessage { Role = m.Role, Text = Cut(m.Text, MaxMessageChars) })
                .ToList();
            if (messages.Count == 0) return;
            if (messages.Count > MaxMessages) messages = messages.Skip(messages.Count - MaxMessages).ToList();

            lock (_lock)
            {
                var read = Refresh();
                if (string.IsNullOrEmpty(chat.Id)) chat.Id = Guid.NewGuid().ToString("N");
                if (string.IsNullOrWhiteSpace(chat.Title))
                {
                    chat.Title = TitleFrom(messages.FirstOrDefault(m => m.Role == "user")?.Text);
                }
                chat.Title = TitleFrom(chat.Title);
                if (string.IsNullOrWhiteSpace(chat.Scope)) chat.Scope = MailboxScope;
                chat.Updated = _now();
                chat.Messages = messages;

                _index.RemoveAll(c => c.Id == chat.Id);
                _index.Insert(0, new ChatSummary
                {
                    Id = chat.Id, Title = chat.Title, Scope = chat.Scope, Updated = chat.Updated, Count = messages.Count,
                });
                while (_index.Count > MaxChats)
                {
                    var oldest = _index[_index.Count - 1];
                    _index.RemoveAt(_index.Count - 1);
                    _chats.Remove(oldest.Id);
                    _waiting.Add(oldest.Id);
                    Remove(ChatPrefix + oldest.Id);
                }
                _chats[chat.Id] = chat;
                _waiting.Add(chat.Id);

                Write(ChatPrefix + chat.Id, ChatJson(chat));
                WriteIndex(read);
            }
        }

        public void Delete(string id)
        {
            if (string.IsNullOrEmpty(id)) return;
            lock (_lock)
            {
                var read = Refresh();
                if (_index.RemoveAll(c => c.Id == id) == 0) return;
                _chats.Remove(id);
                _waiting.Add(id);
                Remove(ChatPrefix + id);
                WriteIndex(read);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                var read = Refresh();
                foreach (var chat in _index)
                {
                    _waiting.Add(chat.Id);
                    Remove(ChatPrefix + chat.Id);
                }
                _index.Clear();
                _chats.Clear();
                WriteIndex(read);
            }
        }

        /// <summary>A one-line title from the first thing the user said.</summary>
        public static string TitleFrom(string text)
        {
            var line = Regex.Replace(text ?? "", @"\s+", " ").Trim();
            return line.Length == 0 ? "Chat" : Cut(line, MaxTitleChars);
        }

        /// <summary>What was said in a conversation: the user's and Jameel Copilot's text, without tool calls and their results.</summary>
        public static List<SavedMessage> FromHistory(IReadOnlyList<JObject> history)
        {
            var messages = new List<SavedMessage>();
            foreach (var item in history ?? new JObject[0])
            {
                if (item == null || (string)item["type"] != "message") continue;
                var role = item["role"]?.Type == JTokenType.String ? (string)item["role"] : null;
                if (role != "user" && role != "assistant") continue;
                var text = TextOf(item["content"]);
                if (!string.IsNullOrWhiteSpace(text)) messages.Add(new SavedMessage { Role = role, Text = text });
            }
            return messages;
        }

        /// <summary>A saved chat as conversation history, so the next question has its context.</summary>
        public static List<JObject> ToHistory(SavedChat chat)
        {
            return (chat?.Messages ?? new List<SavedMessage>())
                .Select(m => new JObject(
                    new JProperty("type", "message"), new JProperty("role", m.Role), new JProperty("content", m.Text ?? "")))
                .ToList();
        }

        // -- the store ----------------------------------------------------------

        /// <summary>
        /// Brings the index up to date with what is stored now, keeping this
        /// session's own unwritten changes. False when the store could not be
        /// read: what is stored is then unknown and must not be written over.
        /// </summary>
        private bool Refresh()
        {
            var stored = ReadIndex();
            if (stored == null) return false;

            var mine = _index.Where(c => _waiting.Contains(c.Id)).ToList();
            var merged = stored.Where(c => !_waiting.Contains(c.Id)).Concat(mine)
                .OrderByDescending(c => c.Updated)   // a stable sort: the stored order holds among equals
                .ToList();
            _index.Clear();
            _index.AddRange(merged);
            foreach (var gone in _chats.Keys.Where(id => merged.All(c => c.Id != id)).ToList()) _chats.Remove(gone);
            return true;
        }

        private void WriteIndex(bool storedIndexWasRead)
        {
            if (storedIndexWasRead && Write(IndexName, IndexJson(_index))) _waiting.Clear();
        }

        /// <summary>The stored index; empty when nothing is stored or it is damaged; null when it could not be read.</summary>
        private List<ChatSummary> ReadIndex()
        {
            var json = Read(IndexName);
            if (json == null) return null;

            var index = new List<ChatSummary>();
            try
            {
                var chats = (json.Trim().Length == 0 ? null : StoredJson.Parse(json))?["chats"] as JArray;
                foreach (var entry in (chats ?? new JArray()).OfType<JObject>())
                {
                    var id = Text(entry["id"]);
                    if (string.IsNullOrEmpty(id) || index.Any(c => c.Id == id)) continue;
                    index.Add(new ChatSummary
                    {
                        Id = id,
                        Title = TitleFrom(Text(entry["title"])),
                        Scope = Text(entry["scope"]) ?? MailboxScope,
                        Updated = Date(entry["updated"]),
                        Count = entry["count"]?.Type == JTokenType.Integer ? (int)entry["count"] : 0,
                    });
                }
            }
            catch (JsonException)
            {
                // An index that cannot be understood is an empty one; the next write replaces it.
            }
            return index;
        }

        private static SavedChat ParseChat(string json)
        {
            try
            {
                var o = StoredJson.Parse(json);
                var id = Text(o?["id"]);
                if (string.IsNullOrEmpty(id)) return null;
                var chat = new SavedChat
                {
                    Id = id, Title = TitleFrom(Text(o["title"])), Scope = Text(o["scope"]) ?? MailboxScope, Updated = Date(o["updated"]),
                };
                foreach (var m in (o["messages"] as JArray ?? new JArray()).OfType<JObject>())
                {
                    var role = Text(m["r"]);
                    var text = Text(m["t"]);
                    if ((role == "user" || role == "assistant") && !string.IsNullOrWhiteSpace(text))
                    {
                        chat.Messages.Add(new SavedMessage { Role = role, Text = text });
                    }
                }
                return chat;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static string ChatJson(SavedChat chat)
        {
            return new JObject(
                new JProperty("v", 1),
                new JProperty("id", chat.Id),
                new JProperty("title", chat.Title),
                new JProperty("scope", chat.Scope),
                new JProperty("updated", chat.Updated.ToString("o")),
                new JProperty("messages", new JArray(chat.Messages.Select(m =>
                    new JObject(new JProperty("r", m.Role), new JProperty("t", m.Text)))))).ToString(Formatting.None);
        }

        private static string IndexJson(IEnumerable<ChatSummary> index)
        {
            return new JObject(
                new JProperty("v", 1),
                new JProperty("chats", new JArray(index.Select(c => new JObject(
                    new JProperty("id", c.Id),
                    new JProperty("title", c.Title),
                    new JProperty("scope", c.Scope),
                    new JProperty("updated", c.Updated.ToString("o")),
                    new JProperty("count", c.Count)))))).ToString(Formatting.None);
        }

        /// <summary>The stored text ("" for nothing stored), or null when there is no store or it could not be read.</summary>
        private string Read(string name)
        {
            try { return _store?.Load(name); }
            catch (Exception) { return null; }
        }

        private bool Write(string name, string json)
        {
            try { return _store != null && _store.Save(name, json); }
            catch (Exception) { return false; }   // the chat stays for this session
        }

        private void Remove(string name)
        {
            try { _store?.Delete(name); }
            catch (Exception) { /* an item left behind is never read: the index no longer names it */ }
        }

        private static string Text(JToken token)
        {
            return token != null && token.Type == JTokenType.String ? (string)token : null;
        }

        private static DateTimeOffset Date(JToken token)
        {
            return DateTimeOffset.TryParse(Text(token), System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out var date) ? date : DateTimeOffset.MinValue;
        }

        private static string TextOf(JToken content)
        {
            if (content == null) return "";
            if (content.Type == JTokenType.String) return (string)content;
            if (content.Type != JTokenType.Array) return "";
            return string.Concat(((JArray)content).OfType<JObject>().Select(part => Text(part["text"]) ?? ""));
        }

        private static string Cut(string text, int max)
        {
            const char ellipsis = (char)0x2026;
            return text.Length <= max ? text : text.Substring(0, max - 1) + ellipsis;
        }
    }
}
