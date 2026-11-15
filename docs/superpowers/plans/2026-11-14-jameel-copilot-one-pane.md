# Jameel Copilot 1.2 "One Jameel Copilot" Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** One Jameel Copilot pane with a product-level look, saved chats, a memory the user can open and edit, and a reply flow that ends in a draft card (Insert / Regenerate / Discard).

**Why now:** On 2026-11-14 the owner tested 1.1.2 and found none of what the agent design promised on screen: two near-identical Inbox panes, a WinForms Actions tab, chats lost when a pane closes, no memory. This plan pulls the visible parts of agent-design phases 3, 6 and 8 forward.

**Architecture:** The two existing controllers stay (`ChatController` for a compose window, `InboxCopilotController` for the Inbox); the Reports pane and the Actions tab are removed and their functions move into the page. Everything new that both panes need lives once: storage in `Services/Prefs` (over `IUserStore`, hidden mailbox items), the page protocol in `TaskPane/Chat/ChatsBridge.cs`, the look in `WebUI/`.

**Spec:** `docs/superpowers/specs/2026-11-13-jameel-copilot-agent-design.md` (5.2 memory, 5.9 Synamail ideas, 5.11 look). Mockup shown to the owner: three screens (Inbox, Reply, Chats).

## Global Constraints

- Nothing new leaves the PC except to the Custom API. Chats and memory are kept in the user's own mailbox (hidden items `Jameel Copilot.Chats`, `Jameel Copilot.Chat.<id>`, `Jameel Copilot.Memory`), never on disk.
- A store never throws; without it the feature works for the session only.
- Mail text and memory notes are data: a note is one line, a block line is prefixed, nothing sender-written can start a line that reads as a label.
- No send, delete or move. Jameel Copilot puts text into an email only through `ComposeDocument`, and only on a button the user presses.
- Model output stays untrusted in the page: `markdown.js` escapes; new page code builds DOM with `textContent`.
- Old-style csproj: every new `.cs` gets a `<Compile Include>`.
- Version 1.2.0.

## Review Focus

1. A saved chat reopened after Outlook restarts shows its messages and the next question still has their context.
2. Two panes open at once (Inbox and a reply) do not overwrite each other's chat.
3. A memory note with a line break or a "label:" cannot change the prompt's structure.
4. Insert from a draft card on a reply whose window was closed fails with a message, not an exception.
5. The mailbox store unavailable (cached mode off, shared mailbox): chats and memory still work until Outlook closes.

---

### Task 1: Storage for chats and memory, and the memory block in the prompt

**Files:** create `Services/Prefs/ChatArchive.cs`, `Services/Prefs/UserMemoryService.cs`; modify `IUserStore.cs`, `MailboxUserStore.cs` (`bool Delete(string name)`), `SystemPromptBuilder.cs`, `ConversationContext.cs` (`string Memory`), `CustomApiChatService.cs`; tests `Services/Prefs/ChatArchiveTests.cs`, `UserMemoryServiceTests.cs`, `SystemPromptBuilderTests.cs`.

**Produces:**
- `ChatArchive(IUserStore store, Func<DateTimeOffset> now = null)`: `IReadOnlyList<ChatSummary> List()` (newest first), `SavedChat Load(string id)`, `void Save(SavedChat chat)`, `void Delete(string id)`, `void Clear()`, `ChatSummary Latest(string scope)`; statics `TitleFrom(string)`, `FromHistory(IReadOnlyList<JObject>)`, `ToHistory(SavedChat)`. Limits: 40 chats, 80 messages a chat, 12,000 characters a message. Only user and assistant text is kept; tool calls and results are not.
- `UserMemoryService(IUserStore store)`: `Notes`, `Add(text)`, `Update(id, text)`, `Remove(id)`, `PromptBlock()`. Limits: 40 notes, 300 characters, one line.
- `SystemPromptBuilder.Build(useTools, pane, context, now = null, memory = null)`: the memory block comes after the date and before the mailbox context, headed `User preferences (written by the user for you to remember; follow them unless this conversation says otherwise):`.

### Task 2: The page: look, header, views, grouped chips, draft cards

**Files:** `WebUI/index.html`, `chat.js`, `styles.css`; tests `WebUiSourceTests.cs`, `TaskPane/Chat/WebView2LockDownTests.cs` (DOM checks in a real WebView2).

**Produces (page API):** `setQuickActions(groups)` with `{title, chips:[{label, prompt, mode}]}` where mode is `send`, `fill` or `action` (posts `chip_action {id}`); `loadConversation(messages)`; `setChats(list, currentId)`; `setMemory(notes)`; `beginDraft(id, title)`, `finalizeDraft(id, buttons, opts)`; `setComposeMode(on)`.
**Page messages:** `new_chat`, `show_chats`, `open_chat {id}`, `delete_chat {id}`, `clear_chats`, `show_memory`, `memory_add {text}`, `memory_update {id,text}`, `memory_remove {id}`, `chip_action {id}`, `draft {text}`, `draft_action {id, action}`.

### Task 3: ChatsBridge, wired into both panes

**Files:** create `TaskPane/Chat/ChatsBridge.cs`; modify `ChatController.cs`, `InboxCopilotController.cs`, `ThisAddIn.cs` (`ChatArchive`, `UserMemory`), the two panes; tests `TaskPane/Chat/ChatsBridgeTests.cs`.

**Produces:** `ChatsBridge(ChatArchive, UserMemoryService, ConversationStore, Func<string,Task> runScript, string scope, string title)`: `Task OpenAsync()`, `bool Handle(type, payload)`, `void Save()`, `string Memory`. Scope `mailbox` for the Inbox pane, `thread:<ConversationID>` for a compose window. A thread reopens its chat; the Inbox reopens its last chat if it was used today.

### Task 4: Compose window: actions in the page, the Actions tab removed

**Files:** `ChatController.cs`, `AITaskPane.cs`, `AITaskPane.Designer.cs`, `Services/ComposePrompt.cs`; tests `ComposePromptTests.cs`, `CustomApiChatServiceTests.cs`.

Chips: "Write my reply" (Formal, Friendly, Short, Detailed) in a reply; "Edit what I wrote" (Proofread, Improve, Shorten, Lengthen, Formal, Friendly). The input's second button, Draft, writes from the typed instructions. A result is a draft card: Insert, Replace my text (when there is text), Regenerate, Discard.

### Task 5: Inbox: one pane, Reports merged, Draft a reply opens a reply

**Files:** `InboxCopilotController.cs`, `QuickActionChip.cs`, delete `TaskPane/InboxReports/*`, `ThisAddIn.cs`, `Ribbon.xml`, `Ribbon.cs`, `Prompts/pane-inbox.txt`; tests `QuickActionChipTests.cs`, `RibbonTests.cs`, delete `ReportQuickActionChipTests.cs`.

Chip groups: This email, My mailbox, Reports (prompts with a `[placeholder]` fill the input instead of sending). "Draft a reply" gives a draft card with Open reply, Open reply all, Regenerate, Discard; Open reply opens Outlook's reply window with the text above the signature.

### Task 6: "Remember this" from chat

A tool `jc_remember` that saves nothing: its result makes the page show a card "Save to memory?" and only the user's Save writes the note.

### Task 7: Guide, README, CLAUDE.md, version 1.2.0, MSI

`Deploy/README.txt` (per-mailbox items, acceptance checks), `README.md`, `CLAUDE.md`; `Installer\Build-Msi.ps1 -Version 1.2.0`.
