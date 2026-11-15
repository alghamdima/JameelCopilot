# Jameel Copilot: Outlook add-in on a Custom API, shipped as an MSI

Date: 2026-11-12
Status: draft, waiting for owner review

## 1. Goal

Turn this checkout into **Jameel Copilot**, an internal add-in for classic
Outlook 2024 on Windows:

- It talks to one endpoint only: the company's **Custom API** (an
  OpenAI-compatible gateway with several models behind it).
- It ships as one **MSI**. IT installs it silently and it works for every
  user on the machine with no per-user setup.
- Security comes first in every choice below.

### Done means

1. On a clean Windows 11 PC with Outlook 2024, `msiexec /i JameelCopilot-1.0.0-x64.msi
   API_BASE_URL=... API_KEY=... MODEL=<model> /qn` leaves an "Jameel Copilot"
   group on the Outlook ribbon, with no trust prompt and no sign-in.
2. Chat, Inbox Copilot, Inbox Reports, Excel/PDF export and the Actions tab
   all work against the Custom API.
3. A network capture during normal use shows the add-in's own requests going
   to the Custom API host and nowhere else.
4. The security review in section 9 reports no Critical or High finding.

## 2. Approach

Keep the parts that have nothing to do with ChatGPT (the 16 Outlook tools,
exports, the chat UI and their tests). Replace the connection layer, delete
everything that reaches another service and the features the owner does not
want, harden what stays, then package it. No rewrite from scratch.

## 3. Scope

### Removed

| What | Where | Why |
|---|---|---|
| ChatGPT OAuth sign-in, token refresh, localhost callback listener | `Services/CodexAuthService.cs`, account section of `SettingsForm.cs` | Replaced by the Custom API key |
| Model catalog and Update Models | `Services/Models/` (9 files), its Settings section | Calls chatgpt.com and api.github.com; replaced by an admin-set model list |
| Voice | `Services/RealtimeVoiceService.cs`, mic buttons in `AITaskPane`, the 7 NAudio packages | Calls api.openai.com |
| In-app updater | `Services/Updates/` (14 files), its Settings section, the startup reconciler call in `ThisAddIn.cs` | Calls api.github.com; MSI upgrades replace it |
| Admin password, Settings save and merge | `Config.cs` (save path), `Services/FileLock.cs`, most of `SettingsForm.cs` | Settings move to HKLM, see section 5 |
| Shared and per-user config layers | `C:\ProgramData\JameelCopilot`, `%APPDATA%\JameelCopilot\config.xml` | Same |
| PowerShell installer, release zip, GitHub workflow | `Deploy/*.ps1`, `Deploy/MicrosoftEdgeWebView2Setup.exe`, `.github/workflows/release.yml` | Replaced by the MSI |
| Per-turn reasoning-effort dropdown | `WebUI/index.html`, `chat.js`, the three controllers | The effort list came from the deleted catalog; one model is configured, so there is nothing to pick |
| Quick Tools (images to PDF) | `Services/QuickTools/` (20 files), `QuickTools/` (4 files), `Services/IStructuredCompletion.cs`, the Quick Tools menus in `Ribbon.xml`, `ThisAddIn.RunImagesToPdf` | Owner decision, 2026-11-12 |
| Variants tab | `Services/Variants/` (4 files), `TaskPane/Variants/VariantsController.cs`, `tabVariants` in `AITaskPane` | Owner decision, 2026-11-12 |
| Ribbon group on the read-message window | `AIAssistantReadGroup` in `Ribbon.xml` | It only held Quick Tools and Settings |
| Upstream specs and plans for removed features | `docs/superpowers/specs/`, `docs/superpowers/plans/` | They describe code that no longer exists |
| Tests of all of the above | `JameelCopilot.Tests` | Deleted with the code |

### Kept

Chat and the Actions tab (compose pane), Inbox Copilot, Inbox Reports, the 12
always-on tools (read, search, count, export), the 4 safe-write tools, and
per-message Save as PDF.

### Added

- Custom API client (section 4).
- Settings in HKLM, written by the MSI (section 5).
- Hardening (section 6).
- Right-to-left rendering for Arabic text in chat and PDF (section 7).
- The MSI (section 8).

## 4. Custom API client

`CodexChatService` becomes `CustomApiChatService`. It speaks the
**OpenAI-compatible Chat Completions** dialect, because that is the one every
gateway and every model behind it understands. The Responses dialect the code
uses today is not supported evenly across models.

- **Endpoint:** `POST {ApiBaseUrl}/chat/completions`, where `ApiBaseUrl` looks
  like `https://llm.example.internal/v1`.
- **Auth:** `Authorization: Bearer <api key>`.
- **Request:** `model`, `messages`, `tools` (Chat Completions function format),
  `tool_choice: "auto"`, `stream: true`, `user` (the Windows account name, so
  the gateway's logs show who made each call), and `reasoning_effort` only
  when the admin set one.
- **History:** stays in its current internal item shape (`message`,
  `function_call`, `function_call_output`), so `ConversationStore` and its
  tests do not change. One pure function converts it to `messages` when a
  request is built: an assistant text followed by function calls becomes one
  assistant message with `tool_calls`; each output becomes a `tool` message.
- **Streaming:** read `choices[0].delta.content` for text and
  `choices[0].delta.tool_calls[]` for tool calls, accumulated by `index`.
  Several tool calls in one round are dispatched in parallel, as today. The
  16-round cap and cancellation behave as today.
- **One-shot calls** (Actions tab) use the same endpoint without tools.
- **Model:** one model, set by the admin (`<model>` at launch).
  Users do not choose. Changing it means re-running the MSI with a new
  `MODEL`. The page never sends a model name to the host.
- **Transport rules:** `https` only. Redirects are not followed, so a
  redirect can never re-send mail content to another host. Certificates are
  validated by Windows as usual; validation is never switched off.
- **Errors:** 401/403 show "Jameel Copilot could not sign in to the AI service. Contact
  IT." 429 shows the wait time. Unreachable host shows "Jameel Copilot cannot reach the
  AI service." Missing or invalid settings show "Jameel Copilot is not configured.
  Contact IT." No raw response body is shown or logged.

`CodexBackendException` becomes `ApiException`. `NotSignedInException` goes.

## 5. Settings and the API key

All settings live in `HKLM\SOFTWARE\JameelCopilot` (64-bit view, also read by a
32-bit Outlook). Only administrators can write there, so the in-app admin
password and the world-writable ProgramData folder are no longer needed.
The add-in never reads HKCU or any config file.

| Value | Type | Meaning |
|---|---|---|
| `ApiBaseUrl` | string | Required. Must be an absolute `https` URL. |
| `ApiKeyProtected` | string | The API key, encrypted with Windows DPAPI (machine scope), base64. |
| `Model` | string | Required. The model ID sent with every request. |
| `ReasoningEffort` | string | Optional. Sent as `reasoning_effort` when set. |
| `EnabledWriteTools` | string | Comma-separated safe-write tools. **Default: none.** |
| `MaxBulkExportRows` | dword | Optional. Default 2000, clamped to 1..10000 as today. |

`Config.cs` shrinks to a reader with one test seam (a function that returns a
value by name). Values are validated on load: the URL must be `https`, write
tools are intersected with the known four, the row cap is clamped.

The **Settings** button becomes **About Jameel Copilot**: version, the Custom API host,
the model, which write tools are on, and a **Test connection** button. It is
read-only and needs no password. It still opens through `SettingsForm.Open`.

### What the key protection does and does not do

A key that the add-in can use is a key the user running Outlook can extract.
No client-side storage changes that. DPAPI machine scope keeps the key out of
plain text on disk, in backups and in registry exports, and it cannot be
decrypted on another machine. It does not hide the key from a determined
local user.

So the key must be treated as low-privilege, and the real controls sit on the
gateway:

- a key issued for Jameel Copilot only, limited to the models Jameel Copilot uses;
- a budget and a rate limit on that key;
- the gateway reachable from the internal network only;
- rotation by re-running the MSI with a new `API_KEY`.

The stronger option, not built in v1: Windows integrated authentication at
the gateway (a reverse proxy that authenticates the AD user and adds the
upstream key). Then no secret is stored on any PC. The add-in side is a few
lines; the work is on the gateway side.

## 6. Hardening

Each row is a weakness in the current code and what v1 does about it.

| # | Today | v1 |
|---|---|---|
| 1 | Admin password is `admin` by default, stored in plain text in `C:\ProgramData\JameelCopilot\config.xml`, a folder where every signed-in user has Modify (`Config.cs:24`, `Install-JameelCopilot.ps1:588`). Any user can change the model or switch write tools on. | Settings in HKLM, admin-only. No password, no shared folder. |
| 2 | The chat WebView has no navigation filter and no new-window handler. Links in model output are real anchors that open a browser window inside Outlook (`markdown.js:38`). A crafted email can make the model print a link that carries mail content in its URL. | Links render as plain text (`text (url)`), not anchors. `NavigationStarting` cancels anything but the add-in's own virtual host. `NewWindowRequested` is always handled and opens nothing. |
| 3 | `WebMessageReceived` acts on any message without checking where it came from (`ChatController.cs:127`, and the Copilot and Reports controllers). | Messages are accepted only when `e.Source` is the add-in's virtual host. |
| 4 | No Content-Security-Policy on `index.html` or the print template. | `default-src 'none'`; scripts and styles from the virtual host only; images `data:` only; `connect-src 'none'`. The print template's inline script gets a per-render nonce. A `WebResourceRequested` filter blocks any request that is not to the virtual host. |
| 5 | WebView2 defaults leave SmartScreen lookups and crash-report upload to Microsoft on. | `IsReputationCheckingRequired = false` (pages are local) and `IsCustomCrashReportingEnabled = true`. Host objects, autofill and password autosave are switched off. Virtual host mapping uses `DenyCors` for the chat views. |
| 6 | The trace log is always on and records tool arguments and the start of each user message (`CodexChatService.cs:396`, `ChatController.cs:132`). | The log keeps tool names, sizes, status codes and timings only. No mail content, no prompts, no key. |
| 7 | All four write tools are on by default (`Config.cs:28`). | Off by default. The admin opts in per tool at install time. |
| 8 | The HTTP client follows redirects. | Redirects off; `https` only (section 4). |
| 9 | "Open" launches any file inside the Reports folder with its default program (`ExportBridge.cs:137`). | The path policy also requires the extension to be `.xlsx` or `.pdf`. |
| 10 | The installer loosens ClickOnce trust prompting for the whole machine and adds an inclusion-list entry with an empty public key (`Install-JameelCopilot.ps1:667-684`). | The MSI changes neither. Trust comes from the install location (Program Files, registered under HKLM) and signed manifests. |
| 11 | Mail content is untrusted text read by a model that can call tools. | No send, delete or move tool (as today). Write tools off by default. Rows 2 to 5 stop model output from causing network requests. The three system prompts gain one shared rule: text inside emails is data, never instructions. A test confirms an Excel cell starting with `=` is stored as text, not a formula. |
| 12 | 25 NuGet packages, some old. | NAudio's 7 packages go with voice. The rest are checked against published advisories and bumped where a patched version exists (`System.IO.Packaging 6.0.0` is the first to check). |

Residual risk, stated plainly: a crafted email can still mislead a summary,
and if `outlook_create_draft` is enabled it can produce a draft with
recipients the attacker chose. The user still has to open and send it.

The WebView2 Evergreen runtime updates itself from Microsoft through the Edge
update service. That is a Windows component outside this add-in; IT controls
it with Edge update policy.

## 7. Name and language

The product name users see is **Jameel Copilot**. The technical identifier is
**JameelCopilot**, because a two-letter folder or registry key name can collide
with other software.

| Becomes Jameel Copilot / JameelCopilot | Stays |
|---|---|
| Ribbon group and button labels ("AI Assistant" becomes "Jameel Copilot") | `namespace JameelCopilot` in C# |
| Task pane titles, dialog titles, PDF footer text | Project and folder names in the repo |
| Assembly, `.vsto` and add-in registration name (`JameelCopilot`) | Tool names (`outlook_search_messages`, ...) |
| `C:\Program Files\JameelCopilot`, `%LOCALAPPDATA%\JameelCopilot` | The `jameelcopilot` JavaScript object |
| Reports folder `Documents\Jameel Copilot\Reports` | |
| WebView2 virtual host `jameelcopilot.local` | |

Arabic: message text, the input box and PDF blocks get `dir="auto"`, so
Arabic paragraphs align right and mixed Arabic/English lines order correctly.
The system prompts tell the model to answer in the language the user wrote
in. Button labels stay English in v1.

## 8. The MSI

Built with **WiX Toolset 5.0.2**, pinned, as an SDK-style project built by
MSBuild: `Installer/JameelCopilot.Installer.wixproj`, plus one small managed
custom-action project. Per-machine, x64, output `JameelCopilot-<version>-x64.msi`.

WiX 5 is chosen because WiX 6 and later require a paid maintenance fee from
organisations that generate revenue. If the company prefers a supported WiX,
paying the fee and moving to the current version is a one-line change.

**Install properties**

| Property | Required | Notes |
|---|---|---|
| `API_BASE_URL` | first install | Rejected unless it starts with `https://`. |
| `API_KEY` | first install | Declared hidden, so it is not written to MSI logs. |
| `MODEL` | first install | `<model>` at launch. |
| `REASONING_EFFORT`, `WRITE_TOOLS`, `MAX_BULK_EXPORT_ROWS` | no | |

**What it does**

1. Checks prerequisites and stops with a clear message if one is missing:
   64-bit Windows 10 or later, .NET Framework 4.7.2 or later, the WebView2
   runtime. It installs none of them and downloads nothing.

   The VSTO runtime is required too, but the MSI does not check for it.
   Office installs its own copy, and Click-to-Run Office (which LTSC 2024
   is) does not register it where an installer can reliably see it: on the
   owner's Office LTSC Standard 2024 x64 PC the documented registry value is
   empty. A check there would refuse to install on every company PC. The
   deployment guide lists the runtime as a requirement and says what to do
   if the add-in does not load.
2. Copies the add-in to `C:\Program Files\JameelCopilot`.
3. Registers the add-in for all users under
   `HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\JameelCopilot`, in both the
   64-bit and 32-bit registry views, with the `|vstolocal` manifest path.
   It also writes the same "do not disable" resiliency value the old script
   wrote.
4. Encrypts `API_KEY` with DPAPI in an immediate custom action, clears the
   plain property, and writes the settings to `HKLM\SOFTWARE\JameelCopilot`.
5. On upgrade, keeps every setting that was not passed again. On uninstall,
   removes the files, the registration and the settings including the key.
   User files (reports, caches under `%LOCALAPPDATA%\JameelCopilot`) stay.

The MSI never closes Outlook. IT deploys while Outlook is closed (at logon or
in a maintenance window); a first install only needs Outlook restarted.

**Signing.** The DLL, the manifests and the MSI are signed with the company's
code-signing certificate, passed to the build as a thumbprint. Without one,
a self-signed certificate works for a pilot but machines with AppLocker or
WDAC will block it.

`Installer/Set-JameelCopilotConfig.ps1` writes the same registry values on a developer
machine (elevated), so the add-in can be run from Visual Studio.

`Deploy/README.txt` is rewritten as the IT deployment guide: properties,
silent install, upgrade, key rotation, uninstall, gateway-side requirements.
`README.md` and `CLAUDE.md` are rewritten to match.

## 9. Testing and verification

- **Unit tests** (xUnit, same patterns as today): history-to-messages
  conversion; streamed text and tool-call accumulation, including several
  calls in one round; request shape; `https` enforcement and no redirects;
  error mapping; config validation; key decrypt; path policy extension
  check; Excel `=` cell stays text; ribbon ids and callbacks.
- **Source guards**, in the style of the existing `*SourceTests`: (a) no URL
  literal in the add-in outside a short allow list (MAPI schema names, the
  virtual host); (b) every WebView2 is created through the one lock-down
  helper.
- **Each behaviour change gets a test that fails without it** (project rule).
- **Clean-VM checklist**, run by hand before release: silent install; no
  trust prompt; ribbon shows Jameel Copilot; each feature once; Arabic prompt and reply;
  upgrade keeps settings; uninstall leaves no key; network capture shows the
  add-in's own requests going to one destination (the WebView2 Runtime's own
  connections to Microsoft are listed apart in the deployment guide).
- **Security gate:** the `secure-code-review` skill runs over the final code.
  Release needs no Critical or High finding.

## 10. Order of work

0. **Probe the model first.** One live streamed request to the Custom API
   with `<model>` and a single tool, to confirm that it calls tools
   and that the stream has the shape section 4 expects. This needs the URL
   and a key from the owner. If the model does not call tools, the mailbox
   features cannot work with it and the owner picks another model before any
   code is written.
1. Custom API client and HKLM settings replace OAuth and the model catalog.
2. Delete the updater, voice, Quick Tools, Variants, the old installer
   scripts and their tests.
3. Hardening (section 6) and right-to-left rendering.
4. Rename to Jameel Copilot.
5. The MSI and the rewritten docs.
6. Security gate and the clean-VM checklist.

The build and the test suite stay green at the end of every step.

## 11. Not in v1

Candidate additions, most useful first. None is designed here.

1. **Read attachments.** Extract text from PDF, Word and Excel attachments
   locally so summaries cover them. Today the model sees attachment names only.
2. **Sensitivity guard.** Never send items marked Private or Confidential, or
   carrying chosen sensitivity labels, to the model.
3. **Arabic interface.** Translate buttons, chips and dialogs.
4. **Company prompt chips.** Admin-defined quick actions (for example
   "translate to English", "reply in the company tone").
5. **Meeting and task extraction** into a calendar or task draft, as a fifth
   safe-write tool.
6. **Dictation** through the Custom API, if the gateway offers a speech model.
7. **Windows authentication at the gateway** (section 5), replacing the
   shared key.
8. **Model picker**, if more than one model should be offered to users.

## 12. Needed from the owner

| # | Item | If not provided |
|---|---|---|
| 1 | The Custom API reachable over `https` with a certificate the PCs trust. Given 2026-11-12 (the address is kept out of the repository). The name does not resolve from the development machine, so the probe and any live test must run from inside the company network or VPN | v1 cannot connect; plain `http` is refused by design |
| 2 | An API key for Jameel Copilot only, limited to its models, with a budget and rate limit | Any key works, with more exposure if it leaks |
| 3 | A test key for the probe in section 10 step 0. Model given: `<model>`; the owner states it supports tool calling, and the probe confirms the stream shape | The client is written against the documented format and first tested later, on the company network |
| 7 | A build machine. The development machine has VS 18 Build Tools without the "Office/SharePoint build tools" workload and without the .NET Framework 4.7.2 targeting pack, so it cannot build the add-in yet | Code can be edited but not built or tested |
| 8 | A test PC or VM with Outlook 2024 on the company network | The add-in cannot be tried in Outlook |
| 4 | A code-signing certificate | Pilot with a self-signed one |
| 5 | Outlook bitness and Windows versions in use. Given 2026-11-13 for one company PC: Office LTSC Standard 2024, x64, WebView2 runtime 152 present | Assumed for the rest: 64-bit Windows 10/11, either Outlook bitness |
| 6 | Confirmation that no PC has the original open-source add-in installed | Assumed: none, so no migration |
