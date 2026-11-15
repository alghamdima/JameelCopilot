# Jameel Copilot for Outlook — Developer Guide

## Project Overview
Jameel Copilot is a VSTO add-in for classic Outlook on Windows: one assistant pane, the
same page in the Inbox and in a compose window (chat, reply and edit drafts,
reports, saved chats, memory), and Excel/PDF exports. It runs
in-process and talks to one service, the company's Custom API (a Chat
Completions endpoint). The address, key and model are set by an administrator
through the MSI; there is no sign-in, no updater and no other network
destination.

Jameel Copilot derives from an open-source MIT project. The C# namespace, the assembly and
the add-in registration are `JameelCopilot`; what a user sees is `Jameel Copilot`.

Stack: C# 7.3 on .NET Framework 4.7.2 (VSTO, WinForms), WebView2 with HTML/JS
for the chat UI, xUnit tests, WiX 5 for the MSI, Windows PowerShell 5.1 scripts.

## Quick Start
```powershell
# Once per machine: a certificate to sign the add-in's manifests.
$thumb = (New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Jameel Copilot dev" -CertStoreLocation Cert:\CurrentUser\My).Thumbprint
# Later sessions: (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object Subject -eq 'CN=Jameel Copilot dev').Thumbprint

# Build the add-in and the tests; run everything, or one class
.\tools\Build.ps1 -CertThumbprint $thumb -Test
.\tools\Build.ps1 -CertThumbprint $thumb -Test -Filter "FullyQualifiedName~ConfigTests"

# Build the MSI into out\
.\Installer\Build-Msi.ps1 -Version X.Y.Z -CertThumbprint $thumb
```

Build needs MSBuild with the "Visual Studio Tools for Office (VSTO) build
tools" component, and the .NET SDK. Office is not needed to build or test.

## Architecture
- `VSTO2/JameelCopilot/` — the add-in.
  - `ThisAddIn.cs`, `Ribbon.cs`, `Ribbon.xml`; `TaskPane/` hosts the two
    controllers of the one page (`Chat/ChatController.cs` for a compose
    window, `InboxCopilot/InboxCopilotController.cs` for the Inbox) and the
    WebView2 bridge. What both need lives once: `Chat/ChatsBridge.cs` (saved
    chats and memory), `Chat/ComposeChips.cs` (chips and draft-card scripts),
    `Chat/PrefsBridge.cs`, `Chat/ExportBridge.cs`.
  - `Services/CustomApiChatService.cs` — the only code that makes network
    requests: streaming, multi-round tool dispatch, user-facing errors.
    `ChatCompletionsWire.cs` (history to request) and `ChatStreamReader.cs`
    (streamed answer to text and tool calls) are pure and tested alone.
  - `Prompts/*.txt` + `Services/SystemPromptBuilder.cs` — the system prompt:
    a fixed core (language rules first), the pane's text, then mailbox
    context under a "data" header.
  - `Services/ComposeBody.cs`, `ComposePrompt.cs` + `Services/Tools/ComposeDocument.cs`
    — the compose window split into the user's own text, their signature and
    the quoted message; what the model is told about it; and the only code
    that writes into an email.
  - `Services/MachineSettings.cs` + `Config.cs` — settings from
    `HKLM\SOFTWARE\JameelCopilot`; `Services/ApiKeyProtector.cs` — the stored
    key's format (DPAPI, machine scope, base64).
  - `Services/ResponseStyles.cs` — the user's Precise/Balanced/Creative
    choice as a temperature within the administrator's limits;
    `Services/Prefs/` — what is kept in hidden items in the user's own
    mailbox: preferences (`UserPrefsService`), saved chats (`ChatArchive`)
    and memory notes (`UserMemoryService`).
  - `Services/LogoImages.cs` + `Resources/` — the Jameel Copilot mark for the ribbon and
    the About dialog (`WebUI/jc-mark.svg` is the source).
  - `Services/OutlookToolHost.cs` + `Services/Tools/` — the 23 model-callable
    tools: 20 over Outlook COM (mail, calendar, address book; four of them
    writes an administrator enables) and three that only propose
    (`jc_remember`, `jc_propose_meeting`, `jc_propose_rule`);
    `Services/Export/` — ClosedXML Excel, WebView2 PDF.
  - `SettingsForm.cs` — the read-only About dialog.
  - `WebUI/` — `chat.js`, `markdown.js`, styles and the print template.
- `VSTO2/JameelCopilot.Tests/` — xUnit (SDK-style project): unit tests, STA
  WinForms tests, and source guards that read the repo's files.
- `Installer/` — `Package.wxs`, the custom-action project, `Build-Msi.ps1`,
  `Set-JameelCopilotConfig.ps1`.
- `tools/` — `Build.ps1`; `Probe-CustomApi.ps1` and `Probe-Language.ps1`
  (run by the owner on the company network); `Render-Logo.ps1`.
- `Deploy/README.txt` — the deployment guide for IT.
- `docs/superpowers/` — the designs and the implementation plans (1.0, and
  the agent design with one plan per phase).

## Documentation Lookup Table
| Topic | Documentation | Key files |
|---|---|---|
| Install, settings, key rotation, uninstall | `Deploy/README.txt` | `Installer/Package.wxs` |
| Design and security decisions | `docs/superpowers/specs/2026-11-12-jameel-copilot-custom-api-msi-design.md` | — |
| The agent: actions, memory, prompts, token budget, phases | `docs/superpowers/specs/2026-11-13-jameel-copilot-agent-design.md` | — |
| System prompt and language | agent design 5.5 | `Prompts/`, `Services/SystemPromptBuilder.cs`, `tools/Probe-Language.ps1` |
| Model settings and response style | agent design 5.10 | `Services/MachineSettings.cs`, `Services/ResponseStyles.cs`, `Services/Prefs/` |
| Custom API client | spec section 4 | `Services/CustomApiChatService.cs`, `ChatCompletionsWire.cs`, `ChatStreamReader.cs` |
| Settings and the key | spec section 5 | `Services/MachineSettings.cs`, `Config.cs`, `Services/ApiKeyProtector.cs` |
| WebView2 hardening | spec section 6 | `TaskPane/Chat/WebView2Bootstrap.cs`, `WebUI/` |
| Tools | `README.md` | `Services/OutlookToolHost.cs`, `Services/Tools/` |
| Exports | `README.md` | `Services/Export/`, `TaskPane/Chat/ExportBridge.cs` |

## Key Patterns

### Add-in (C#)
- C# 7.3 and .NET Framework 4.7.2. The add-in is an old-style csproj: add
  each new `.cs` file as a `<Compile Include>` entry.
- Settings come only from `HKLM\SOFTWARE\JameelCopilot` (64-bit view), read by
  `MachineSettings` and exposed through `Config`. Nothing in the add-in writes
  them. Reading never throws: a missing or invalid value leaves Jameel Copilot "not
  configured" with `Config.Problem` naming it.
- All network requests go through `CustomApiChatService`. HTTPS only, no
  redirects, certificate validation left to Windows. Errors shown to the user
  come from `ApiErrors`; the service's own error text is never shown or logged.
- The chat service builds every system prompt with `SystemPromptBuilder`:
  the fixed core (`Prompts/core-agent.txt`, or `core-rewrite.txt` for a
  one-shot rewrite), then the pane's text (`ConversationContext.SystemInstructions`,
  from `SystemPromptBuilder.Pane(...)`), then what the pane read from the
  mailbox (`ConversationContext.Context`) under the data header. Callers never
  add rules and never put mail text in `SystemInstructions`. Prompt text is
  edited in the `.txt` files; tests pin the sentences that matter. The service
  adds today's date and the time zone before the context: without it the
  model takes its training date for "today". Dates for the model are written
  with `CultureInfo.InvariantCulture` (the PCs' calendar is Um Al-Qura).
- An answer the service stops at the length limit ends with
  `ApiErrors.CutAtLimit` (`ChatStreamReader`, on `finish_reason: length`).
- Temperature comes from the user's style (`PrefsBridge.Temperature`) within
  the administrator's `Temperature`/`TemperatureMax`; `off` sends none. The
  connection test never sends one. Numbers are parsed and written with
  `CultureInfo.InvariantCulture`: the PCs use Arabic regional settings.
- A store of user data (`IUserStore`) never throws: unavailable means null or
  false, and the data then holds for the session only. `Load` answers "" for
  "nothing stored" and null for "could not be read". The mailbox can be open
  in two Outlook sessions, so `ChatArchive` and `UserMemoryService` read
  again before every change, apply the change to what is stored now, and
  write nothing after a read that failed.
- Writing and editing an email (the chips, the Draft button, "Draft a reply"
  in the Inbox) are one-shot actions on `CustomApiChatService.NewActionContext`
  (`UseTools = false`): no tool is offered and none is run, because the quoted
  thread that travels with the request is not the user's own. The result is
  a draft card. The host keeps the card's text; what a card's button puts
  into an email is that text, not what the page sends back. (An answer in the
  chat that the user inserts with "Insert into email" is sent back by the
  page as text; it goes in as plain text through the editor.)
- Chats are saved after every turn (`ChatsBridge.Save`): the user's text and
  Jameel Copilot's answers, never tool calls or their results. A compose window's chat
  belongs to its thread by the store's ConversationID, never by the subject,
  which the sender chooses. Nothing is swapped while a turn runs.
- Memory notes are the user's. The model can only propose one
  (`jc_remember`); the page's Save button writes it. A note is one line
  (`UserMemoryService.Clean`) and goes into every prompt after the date.
- Jameel Copilot opens Outlook's reply window (`LiveOutlookSurface.OpenReplyWithText`)
  and meeting window (`OpenMeetingRequest`) only from a card's button. It
  never sends, saves or deletes.
- Answers Jameel Copilot shows as data (Prioritize, Review, Tidy) use the structured core
  (`core-structured.txt`, `ConversationContext.Structured`, no tools):
  `StructuredAnswers` puts emails in the prompt (sender text on one line,
  snippets behind "> ") and parses the JSON; the host keeps only ids it gave
  and keeps what a button acts on.
- A chip whose answer is a card (Start my day, Action items, Waiting for a
  reply) runs through `InboxCopilotController.RunCardAsync`: the gate, the
  working line, and what a failure shows. The user line such a chip leaves
  in the chat is Jameel Copilot's own English (`ChipLine`); it is marked so that
  `StructuredAnswers.ReplyLanguage` does not take it for the language the
  user writes in.
- An Outlook task ("Add to tasks") and an Outlook rule ("Create rule") are
  made on the user's click, once, from what the host kept (`_todos`,
  `RulesBridge`), never from what the page sends. A rule only moves mail to
  a folder Jameel Copilot listed (not Deleted Items or Junk) or sets a category:
  `LiveOutlookSurface.CreateRule` has no other action, and a test holds it.
- People for a meeting are resolved first (`outlook_resolve_people`, which
  is Outlook's own name check): a bare first name reaches nobody and no
  calendar is read for it. Rooms come from the address book's room list
  (`outlook_find_rooms`).
- Source guards cut `LiveOutlookSurface.cs` from one method's name to the
  next one's (`OpenMeetingRequest`, `OpenMessage`, `OpenReplyWithText`). Put
  a new public method at the end of the public ones, before `Run<T>`.
- Tidy's Apply changes mail (mark read, flag) on the user's click, whatever
  `EnabledWriteTools` says: that setting governs what the model may do by
  itself. Only unread, unflagged mail is proposed, so Undo restores it.
- Text written with the Write tool here turns `\uXXXX` escapes into the
  characters themselves. Put such an escape into a source file with a script.
- A compose window holds three texts: what the user typed, their signature and
  the message being replied to. Only the first is "the email".
  `ComposeStateResult.BodyPlaintext` is the user's text alone and
  `QuotedPlaintext` the quoted message; `ComposePrompt` names the user as the
  writer and puts the quoted message under Context as someone else's. With
  nothing typed, an action in a reply writes the user's reply in the action's
  manner (`ComposePrompt.Instructions`): the owner presses "Formal" there to
  get a formal reply, and a refusal read as "the buttons do not work". Only a
  new email with nothing typed is refused. The signature is never sent and
  the model is told to write none.
- Text goes into an email only through `ComposeDocument` (the Word editor, on
  the UI thread): above everything, or in place of the user's own text. Never
  assign `MailItem.Body` on an open email: it rebuilds the email as plain
  text and the signature and thread lose their formatting.
- A tool round goes into the history as its calls, then their results.
  `ChatCompletionsWire.ToMessages` sends each round as its own assistant
  message and pairs results with calls inside the round.
- Every WebView2 is created through `WebView2Bootstrap` (`CreateEnvironmentAsync`,
  `LockDown`). Web messages are handled only when
  `WebView2Bootstrap.IsTrustedSource(e.Source)`.
- Model output is untrusted: `markdown.js` escapes everything and renders
  links as text. Do not add anchors, images or inline HTML.
- Open the About dialog only through `SettingsForm.Open`: it re-reads the
  registry first.
- `Ribbon.xml` serves both the Inbox (`TabMail`) and compose
  (`TabNewMailMessage`) ribbons. Every callback needs a public handler in
  `Ribbon.cs` and control ids must be unique, or Office drops the add-in's
  whole ribbon (`RibbonTests`). Built-in icons are `imageMso` names. The
  owner wants Jameel Copilot's two buttons as icons alone: Office always shows a large
  button's label and the simplified (single-line) ribbon ignores
  `showLabel`, so they are `size="normal"`, `showLabel="false"` and
  `getLabel="GetNoLabel"` (empty text); the name is the `screentip`.
- Outlook COM access goes through `OutlookThreadMarshaller` onto the UI
  thread.
- WinForms buttons need explicit `ForeColor`/`BackColor` and
  `UseVisualStyleBackColor = false`; Windows Server 2025 renders them blank
  otherwise.
- The trace log takes names, sizes, codes and timings. Never mail content,
  prompts, tool arguments, tool results, response bodies or the key.
- Test seams are `internal` (e.g. `Config.Apply`, `SettingsForm.DescribeSettings`),
  visible to the tests through `InternalsVisibleTo`.

### Project file
- Every NuGet `HintPath` starts with `..\packages\`.
- A `<Reference Include="Name, Version=...">` must carry the version the DLL
  really has. If they differ MSBuild drops the reference with no warning and
  the DLL is missing from the installed add-in, while the tests still pass on
  their own copy (`ProjectReferenceTests`).
- `UseOfficeInterop` in `DefineConstants` selects the Office interop
  assemblies that ship with the VSTO build tools, so no Office is needed to
  build.
- The WebView2 package's build rules are imported: they copy
  `WebView2Loader.dll` to `runtimes\`. Without it WebView2 cannot start.
- The version is one number, `JcVersion` (default 1.0.0): the build writes the
  assembly and file version attributes from it (`JcStampVersion`), and the
  manifests and About follow. `AssemblyInfo.cs` carries no version.
- `tools\Build.ps1` signs the add-in DLL right after it is compiled
  (`JcSignAssembly`), before the manifests record its hash.

### Installer
- `Package.wxs` is WiX 5 (WiX 6 and later carry a maintenance fee).
- Public properties are remembered across upgrades through `*_SAVED`
  properties; a search straight into the public property would overwrite the
  command line.
- `API_KEY` is hidden and is turned into `API_KEY_PROTECTED` by the immediate
  custom action before anything is written. The custom-action project links
  the add-in's `ApiKeyProtector.cs`: one implementation of the key format.
- The MSI does not check for the VSTO runtime, change machine trust settings,
  or close Outlook. `PackageSourceTests` hold these.
- No settings property is `Secure`: a standard user's command line must never
  reach the elevated side.
- Running the installed MSI again is a maintenance run, which writes nothing
  unless a reinstall is asked for: the `REINSTALL`/`REINSTALLMODE` pair of
  `SetProperty` actions does that.
- The settings' launch conditions begin `REMOVE~="ALL" OR`, not
  `Installed OR`: a run that reaches the elevated side without settings stops
  instead of writing empty ones.
- A major upgrade removes the old version first (`afterInstallInitialize`),
  so no file is kept because its version number did not change.
- `Build-Msi.ps1` refuses to package a DLL with the wrong version or without
  the certificate's signature.
- PowerShell scripts are ASCII-only: Windows PowerShell 5.1 reads a script
  without a byte-order mark as ANSI.

### Testing
- Tests that touch `Config` use `[Collection("Config")]` and
  `new ConfigStateScope(TestSettings.Configured(...))`; WinForms tests run on
  an STA thread (`Sta.Run`).
- The pages have no JavaScript test runner; `WebUiSourceTests` read the files.
  `WebView2LockDownTests` starts a real WebView2 the way the panes do and
  checks the lock-down in it (skipped where the runtime is not installed).
- Guards that must stay green: `EgressGuardTests` (no address literal in the
  add-in), `NameGuardTests` (no old product name shown), `ProjectReferenceTests`,
  `PackageSourceTests`.
- For a behaviour change, check that the new test fails without the fix.

## Release Process
1. `.\Installer\Build-Msi.ps1 -Version X.Y.Z -CertThumbprint <company certificate>`.
   It builds from clean, runs the tests and refuses to package an add-in that
   lacks a file it needs.
2. Run the acceptance checklist in `Deploy/README.txt` on a PC with Outlook on
   the company network.
3. Hand the MSI and its SHA256 to IT.

## Code Quality Standards

### Code Review Checkpoints
After completing any feature or significant change, perform a self-review:
- [ ] All tests pass
- [ ] No hardcoded secrets, credentials, addresses or PII
- [ ] Admin-only settings stay in HKLM; nothing in the add-in can change them
- [ ] No new network destination; no address literal in the add-in
- [ ] Error handling covers the unhappy paths (service unreachable, refused,
      busy, answer cut short, settings missing)
- [ ] No unused usings or dead code introduced
- [ ] Documentation updated if behaviour changed (`README.md`,
      `Deploy/README.txt`)

### Security
- Never commit secrets, the real API address, keys or signing keys (`*.pfx`
  and `*.local.json` are gitignored).
- Validate input at the boundaries: registry values in `MachineSettings`;
  web messages by their source; file actions through `IExportPathPolicy`.
- No send, delete or move tools: model writes are limited to the four
  safe writes, off unless an administrator enables them.

## Data Sanitization
- Tests and docs use fake data only (generic names, `example.com` and
  `llm.example.internal` addresses, 555-xxxx phone numbers).
- Never commit real mailbox content, trace logs or local credentials.

## Important Gotchas
- **Custom API:** `POST {ApiBaseUrl}/chat/completions`, `stream: true`. Gateways
  differ: a tool-call fragment may omit `index`, parallel calls may all carry
  `index` 0 (a different id starts a new call), a call may come without an id
  (the reader gives it one), the last chunk may have no `choices`, a gateway
  may answer a streamed request with one JSON body, and an error object can
  arrive inside a 200 answer. `ChatStreamReaderTests` and
  `CustomApiChatServiceTests` cover these.
- **Build tools:** in Build Tools the VSTO component is *recommended*, not
  required, in the Office/SharePoint workload; add
  `Microsoft.VisualStudio.Component.TeamOffice.BuildTools` explicitly. Build
  Tools' MSBuild cannot build the SDK-style test project and the .NET SDK
  cannot load the VSTO targets, hence `tools\Build.ps1`.
- **Click-to-Run Office** does not register its VSTO runtime in the real
  registry, so an installer cannot detect it.
- **Windows Installer** reports `VersionNT` 603 on Windows 10 and 11.
- **Antivirus** can quarantine the add-in DLL under Program Files and deletes
  scratch `.ps1` files that contain delete loops; drive such experiments from
  Python.
- **WebView2:** the PDF renderer needs its own user-data folder
  (`%LOCALAPPDATA%\JameelCopilot\WebView2PdfData`); sharing the chat's fails with
  HRESULT 0x8007139F. The runtime's own processes contact Microsoft
  (configuration, updates, diagnostic data); an application can switch off
  only crash-dump upload and SmartScreen, and Jameel Copilot does. Microsoft's guidance is
  not to ship browser flags in production, so Jameel Copilot passes none.
- **Windows Installer:** in a run with a user interface the registry search
  runs once, on the client, and a standard user's properties are not passed to
  the elevated side.
- **New Outlook** does not load VSTO add-ins. Jameel Copilot needs classic Outlook.
- **Edge headless** cannot take a transparent screenshot (the switch is
  ignored), which is why `Render-Logo.ps1` captures on white and on black.
- **Ribbon pictures:** `getImage` returns a `System.Drawing.Bitmap`; the
  callback must never throw, or Office drops the add-in's ribbon.
- **Tests and Office types:** the test project cannot name Office interop
  types (they are embedded in the add-in), so a ribbon callback is called
  through reflection.
- **Modal dialogs and DoEvents:** when a WinForms modal dialog or
  `Application.DoEvents()` ends, WinForms leaves a plain
  `SynchronizationContext` on the UI thread, which posts to the thread pool.
  `SettingsForm.Open` puts the context back and `OutlookThreadMarshaller`
  never adopts a plain one. Do the same around any new `ShowDialog`.
- **Start-up:** Outlook disables an add-in whose start takes over a second,
  and the .NET/VSTO runtimes' load alone can cross it. `ThisAddIn_Startup`
  builds nothing (`EnsureServices` does, on the first click; `StartupTests`)
  and names Jameel Copilot in the user's `Resiliency\DoNotDisableAddinList`
  (`AddinResiliency`), as "Always enable this add-in" would.
- An answer lists the emails its tools found or read, built in the page from
  the tool results; a click opens one (`LiveOutlookSurface.OpenMessage`,
  which resolves only ids Jameel Copilot gave out).
- **Compose bookmarks:** Outlook marks the automatic signature (`_MailAutoSig`)
  and the start of the quoted message (`_MailOriginal`) with hidden Word
  bookmarks. They are not documented and no test here can open Outlook: a
  change to `ComposeDocument` needs the compose checks in `Deploy/README.txt`
  run by hand.
