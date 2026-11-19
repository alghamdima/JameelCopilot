using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Prefs
{
    public sealed class MemoryNote
    {
        public string Id { get; set; }
        public string Text { get; set; }
    }

    /// <summary>
    /// What the user asked Jameel Copilot to remember: short notes they can read, change
    /// and delete, kept in their own mailbox and given to the model with
    /// every request. Only the user adds a note. A store that is missing or
    /// fails keeps the notes for the session; nothing here throws.
    /// <para>The mailbox can be open in two Outlook sessions at once, and a
    /// read can fail. So the notes are read again before every change, the
    /// change is applied to what is stored now, and after a read that failed
    /// nothing is written: this session's changes wait for the next write
    /// that follows a good read.</para>
    /// </summary>
    public sealed class UserMemoryService
    {
        public const int MaxNotes = 40;
        public const int MaxNoteChars = 300;
        public const string PromptHeader =
            "User preferences (written by the user for you to remember; follow them unless this conversation says otherwise):";

        private readonly string _itemName;

        private readonly IUserStore _store;
        private readonly object _lock = new object();
        private readonly List<MemoryNote> _notes = new List<MemoryNote>();

        // Notes added, changed or removed here whose change is not stored yet.
        private readonly HashSet<string> _waiting = new HashSet<string>(StringComparer.Ordinal);
        private bool _read;   // the stored notes were read at least once

        /// <param name="store">Null keeps the notes for this session only.</param>
        /// <param name="itemName">The stored item. The user's own prompts are kept the same way, in an item of their own.</param>
        public UserMemoryService(IUserStore store, string itemName = "Memory")
        {
            _store = store;
            _itemName = itemName;
        }

        /// <summary>The notes as they are stored now.</summary>
        public IReadOnlyList<MemoryNote> Notes
        {
            get
            {
                lock (_lock)
                {
                    Refresh();
                    return _notes.Select(n => new MemoryNote { Id = n.Id, Text = n.Text }).ToArray();
                }
            }
        }

        /// <summary>The new note, or null when the text is empty, already there, or the memory is full.</summary>
        public MemoryNote Add(string text)
        {
            var clean = Clean(text);
            lock (_lock)
            {
                var read = Refresh();
                if (clean.Length == 0 || _notes.Count >= MaxNotes || _notes.Any(n => Same(n.Text, clean))) return null;
                var note = new MemoryNote { Id = Guid.NewGuid().ToString("N").Substring(0, 12), Text = clean };
                _notes.Add(note);
                Changed(note.Id, read);
                return new MemoryNote { Id = note.Id, Text = note.Text };
            }
        }

        /// <summary>False when the note is not there, empty, or duplicates another note.</summary>
        public bool Update(string id, string text)
        {
            var clean = Clean(text);
            lock (_lock)
            {
                var read = Refresh();
                var note = _notes.FirstOrDefault(n => n.Id == id);
                if (note == null || clean.Length == 0 || _notes.Any(n => n.Id != id && Same(n.Text, clean))) return false;
                note.Text = clean;
                Changed(note.Id, read);
                return true;
            }
        }

        public bool Remove(string id)
        {
            lock (_lock)
            {
                var read = Refresh();
                if (_notes.RemoveAll(n => n.Id == id) == 0) return false;
                Changed(id, read);
                return true;
            }
        }

        /// <summary>The notes as the prompt carries them, or "" when there are none.</summary>
        public string PromptBlock()
        {
            lock (_lock)
            {
                // Every request asks for this: the mailbox is read until it answers once,
                // then again only when the notes are shown or changed.
                if (!_read) Refresh();
                if (_notes.Count == 0) return "";
                var block = new StringBuilder(PromptHeader);
                foreach (var note in _notes) block.Append("\n- ").Append(note.Text);
                return block.ToString();
            }
        }

        // A note is one line, so it cannot open a section of its own in the prompt.
        internal static string Clean(string text)
        {
            var line = new StringBuilder();
            var space = false;
            foreach (var c in text ?? "")
            {
                // The line and paragraph separators count as white space too.
                if (char.IsWhiteSpace(c) || char.IsControl(c))
                {
                    space = line.Length > 0;
                    continue;
                }
                if (space) line.Append(' ');
                space = false;
                line.Append(c);
            }
            return line.Length <= MaxNoteChars ? line.ToString() : line.ToString(0, MaxNoteChars).TrimEnd();
        }

        private static bool Same(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Brings the notes up to date with what is stored now, keeping this
        /// session's own unwritten changes. False when the store could not be
        /// read: what is stored is then unknown and must not be written over.
        /// </summary>
        private bool Refresh()
        {
            var stored = ReadStored();
            if (stored == null) return false;

            var mine = _notes.Where(n => _waiting.Contains(n.Id)).ToDictionary(n => n.Id, StringComparer.Ordinal);
            var merged = new List<MemoryNote>();
            foreach (var note in stored)
            {
                if (!_waiting.Contains(note.Id)) merged.Add(note);                               // someone else's, or unchanged
                else if (mine.TryGetValue(note.Id, out var changed)) merged.Add(changed);       // changed here
                // else: removed here
            }
            merged.AddRange(_notes.Where(n => _waiting.Contains(n.Id) && stored.All(s => s.Id != n.Id)));   // added here
            _notes.Clear();
            _notes.AddRange(merged.Take(MaxNotes));
            _read = true;
            return true;
        }

        private void Changed(string id, bool storedNotesWereRead)
        {
            _waiting.Add(id);
            if (storedNotesWereRead && Write()) _waiting.Clear();
        }

        /// <summary>The stored notes; empty when nothing is stored or it is damaged; null when it could not be read.</summary>
        private List<MemoryNote> ReadStored()
        {
            string json;
            try { json = _store?.Load(_itemName); }
            catch (Exception) { return null; }
            if (json == null) return null;

            var notes = new List<MemoryNote>();
            try
            {
                var stored = (json.Trim().Length == 0 ? null : StoredJson.Parse(json))?["notes"] as JArray;
                foreach (var entry in (stored ?? new JArray()).OfType<JObject>())
                {
                    var id = entry["id"]?.Type == JTokenType.String ? (string)entry["id"] : null;
                    var text = Clean(entry["text"]?.Type == JTokenType.String ? (string)entry["text"] : null);
                    if (string.IsNullOrEmpty(id) || text.Length == 0 || notes.Count >= MaxNotes || notes.Any(n => n.Id == id)) continue;
                    notes.Add(new MemoryNote { Id = id, Text = text });
                }
            }
            catch (JsonException)
            {
                // Notes that cannot be understood are no notes; the next write replaces them.
            }
            return notes;
        }

        private bool Write()
        {
            try
            {
                return _store != null && _store.Save(_itemName, new JObject(
                    new JProperty("v", 1),
                    new JProperty("notes", new JArray(_notes.Select(n =>
                        new JObject(new JProperty("id", n.Id), new JProperty("text", n.Text)))))).ToString(Formatting.None));
            }
            catch (Exception)
            {
                return false;   // the notes hold for this session
            }
        }
    }
}
