using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Prefs;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// Connects a pane to the user's saved chats and memory: keeps the pane's
    /// conversation after each turn, reopens it when the pane opens again,
    /// and serves the page's Chats and Memory views. One per pane; the
    /// archive and the memory behind it are shared by every pane.
    /// </summary>
    internal sealed class ChatsBridge
    {
        /// <summary>The scope of a new email's chat: kept, never reopened by itself.</summary>
        public const string ComposeScope = "compose";

        private readonly ChatArchive _archive;
        private readonly UserMemoryService _memory;
        private readonly ConversationStore _store;
        private readonly Func<string, Task> _runScript;
        private readonly string _scope;
        private readonly string _title;
        private readonly Func<bool> _busy;
        private readonly Func<DateTimeOffset> _now;
        private readonly UserMemoryService _prompts;
        private readonly Action _promptsChanged;

        // The chat on screen. The id is null until a new chat is first saved.
        private string _chatId;
        private string _chatScope;
        private string _chatTitle;

        /// <param name="scope">What this pane's own chats belong to: the mailbox, or one email thread.</param>
        /// <param name="title">The title of this pane's new chats (a thread's subject); null takes the first thing the user says.</param>
        /// <param name="busy">True while a turn runs: the conversation is not swapped under it.</param>
        /// <param name="prompts">The user's own prompts, shown as chips.</param>
        /// <param name="promptsChanged">Called after the user adds, changes or deletes one: the pane shows its chips again.</param>
        public ChatsBridge(ChatArchive archive, UserMemoryService memory, ConversationStore store, Func<string, Task> runScript,
            string scope = null, string title = null, Func<bool> busy = null, Func<DateTimeOffset> now = null,
            UserMemoryService prompts = null, Action promptsChanged = null)
        {
            _archive = archive ?? new ChatArchive(null);
            _memory = memory ?? new UserMemoryService(null);
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _runScript = runScript ?? throw new ArgumentNullException(nameof(runScript));
            _scope = string.IsNullOrWhiteSpace(scope) ? ChatArchive.MailboxScope : scope;
            _title = title;
            _busy = busy ?? (() => false);
            _now = now ?? (() => DateTimeOffset.Now);
            _prompts = prompts ?? new UserMemoryService(null, "Prompts");
            _promptsChanged = promptsChanged;
            _chatScope = _scope;
            _chatTitle = _title;
        }

        /// <summary>What the user asked Jameel Copilot to remember, as the prompt carries it.</summary>
        public string Memory => _memory.PromptBlock();

        /// <summary>
        /// The user's own prompts as a chip group: each one is sent as it is, or
        /// waits in the input when it has a [placeholder]; the last chip opens
        /// the view where they are added, changed and deleted.
        /// </summary>
        public JObject PromptsGroup()
        {
            // The chips are shown again on every selection change: the mailbox is read for them
            // once, and again when the prompts are shown or changed.
            return (JObject)(_promptsGroup ?? (_promptsGroup = BuildPromptsGroup(_prompts.Notes))).DeepClone();
        }

        private JObject _promptsGroup;

        private static JObject BuildPromptsGroup(IReadOnlyList<MemoryNote> prompts)
        {
            var chips = new JArray(prompts.Select(n => new JObject(
                new JProperty("label", n.Text.Length <= 28 ? n.Text : n.Text.Substring(0, 26).TrimEnd() + "..."),
                new JProperty("prompt", n.Text),
                new JProperty("mode", n.Text.Contains("[") ? "fill" : "send"))));
            chips.Add(new JObject(new JProperty("label", chips.Count == 0 ? "+ Add your own prompt" : "+ Add"), new JProperty("mode", "prompts")));
            return new JObject(new JProperty("title", "My prompts"), new JProperty("chips", chips));
        }

        /// <summary>
        /// When the page is ready: an email thread gets its chat back; the
        /// mailbox gets its last chat if it was used today. Anything older is
        /// under Chats.
        /// </summary>
        public Task OpenAsync()
        {
            var latest = _archive.Latest(_scope);
            if (latest == null) return Task.CompletedTask;
            var today = latest.Updated.ToOffset(_now().Offset).Date == _now().Date;
            return _scope.StartsWith("thread:", StringComparison.Ordinal) || (_scope == ChatArchive.MailboxScope && today)
                ? ShowAsync(latest.Id)
                : Task.CompletedTask;
        }

        /// <summary>Keeps the conversation as it stands. Call after every turn and every draft.</summary>
        public void Save()
        {
            var chat = new SavedChat
            {
                Id = _chatId, Scope = _chatScope, Title = _chatTitle, Messages = ChatArchive.FromHistory(_store.Snapshot()),
            };
            if (chat.Messages.Count == 0) return;
            _archive.Save(chat);
            _chatId = chat.Id;
            _chatTitle = chat.Title;
        }

        /// <summary>True when the message was one of the Chats or Memory views'.</summary>
        public bool Handle(string type, JObject payload)
        {
            switch (type)
            {
                case "new_chat":
                    if (!_busy()) NewChat();
                    return true;
                case "show_chats":
                    PushChats();
                    return true;
                case "open_chat":
                    if (!_busy()) _ = ShowAsync(Text(payload, "id"));
                    return true;
                case "delete_chat":
                    if (!_busy()) Delete(Text(payload, "id"));
                    return true;
                case "clear_chats":
                    if (_busy()) return true;
                    _archive.Clear();
                    NewChat();
                    PushChats();
                    return true;
                case "show_memory":
                    PushMemory();
                    return true;
                case "memory_add":
                    _memory.Add(Text(payload, "text"));
                    PushMemory();
                    return true;
                case "memory_update":
                    _memory.Update(Text(payload, "id"), Text(payload, "text"));
                    PushMemory();
                    return true;
                case "memory_remove":
                    _memory.Remove(Text(payload, "id"));
                    PushMemory();
                    return true;
                case "show_prompts":
                    PushPrompts();
                    return true;
                case "prompt_add":
                    _prompts.Add(Text(payload, "text"));
                    PromptsChanged();
                    return true;
                case "prompt_update":
                    _prompts.Update(Text(payload, "id"), Text(payload, "text"));
                    PromptsChanged();
                    return true;
                case "prompt_remove":
                    _prompts.Remove(Text(payload, "id"));
                    PromptsChanged();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>"Today 09:18", "Yesterday 17:40", "12 Nov 16:20", in the user's time zone.</summary>
        internal static string When(DateTimeOffset updated, DateTimeOffset now)
        {
            var at = updated.ToOffset(now.Offset);
            var time = at.ToString("HH:mm", CultureInfo.InvariantCulture);
            var days = (now.Date - at.Date).Days;
            if (days == 0) return "Today " + time;
            if (days == 1) return "Yesterday " + time;
            return at.ToString(at.Year == now.Year ? "d MMM" : "d MMM yyyy", CultureInfo.InvariantCulture) + " " + time;
        }

        private void NewChat()
        {
            _chatId = null;
            _chatScope = _scope;
            _chatTitle = _title;
            _store.Clear();
            _ = _runScript("jameelcopilot.clear();");
        }

        private Task ShowAsync(string id)
        {
            var chat = _archive.Load(id);
            if (chat == null)
            {
                PushChats();   // the list on the page is out of date
                return Task.CompletedTask;
            }
            _chatId = chat.Id;
            _chatScope = chat.Scope;
            _chatTitle = chat.Title;
            _store.Clear();
            _store.AppendRange(ChatArchive.ToHistory(chat));
            var messages = new JArray(chat.Messages.Select(m =>
                new JObject(new JProperty("role", m.Role), new JProperty("text", m.Text))));
            return _runScript("jameelcopilot.loadConversation(" + messages.ToString(Formatting.None) + ");");
        }

        private void Delete(string id)
        {
            _archive.Delete(id);
            if (id != null && id == _chatId) NewChat();
            PushChats();
        }

        private void PushChats()
        {
            var now = _now();
            var list = new JArray(_archive.List().Select(c => new JObject(
                new JProperty("id", c.Id),
                new JProperty("title", c.Title),
                new JProperty("thread", c.Scope != null && c.Scope.StartsWith("thread:", StringComparison.Ordinal)),
                new JProperty("when", When(c.Updated, now)))));
            _ = _runScript("jameelcopilot.setChats(" + list.ToString(Formatting.None) + ", " + JsonConvert.SerializeObject(_chatId ?? "") + ");");
        }

        private void PushMemory()
        {
            var notes = new JArray(_memory.Notes.Select(n =>
                new JObject(new JProperty("id", n.Id), new JProperty("text", n.Text))));
            _ = _runScript("jameelcopilot.setMemory(" + notes.ToString(Formatting.None) + ");");
        }

        private void PushPrompts()
        {
            var stored = _prompts.Notes;
            _promptsGroup = BuildPromptsGroup(stored);
            var notes = new JArray(stored.Select(n =>
                new JObject(new JProperty("id", n.Id), new JProperty("text", n.Text))));
            _ = _runScript("jameelcopilot.setPrompts(" + notes.ToString(Formatting.None) + ");");
        }

        private void PromptsChanged()
        {
            PushPrompts();
            _promptsChanged?.Invoke();
        }

        private static string Text(JObject payload, string name)
        {
            var token = payload?[name];
            return token != null && token.Type == JTokenType.String ? (string)token : null;
        }
    }
}
