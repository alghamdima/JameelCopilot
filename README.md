# Jameel Copilot for Outlook

Jameel Copilot is a VSTO add-in for classic Outlook on Windows. It adds an assistant that
reads, searches and summarises the signed-in user's mail, helps write replies,
and exports results to Excel or PDF.

It talks to one service: the company's **Custom API**, a Chat Completions
endpoint whose address, key and model an administrator sets at install time.
Jameel Copilot's own code has no other network destination, no sign-in, no telemetry and
no update check. (The Microsoft WebView2 Runtime that draws the chat makes its
own connections to Microsoft; `Deploy/README.txt` says what and how to control
it.)

Jameel Copilot derives from an open-source project (MIT, see [`LICENSE`](LICENSE)).
That project's sign-in to a public AI service, model catalog, voice, in-app
updater, Quick Tools and Variants were removed.

Deploying it? Read [`Deploy/README.txt`](Deploy/README.txt).

## What users get

One Jameel Copilot pane, the same page in the Inbox and in a compose window.

- **In the Inbox**: chat over the mailbox and the selected messages. Chips in
  groups: the selected email (Draft a reply, Summarize, Action items,
  Translate), the mailbox, reports (digest, to-dos, statistics, mail with a
  person), and the user's own prompts.
  "Draft a reply" shows the reply as a card; Open reply puts it in Outlook's
  own reply window, above the signature.
- **In a compose window**: chips that write the reply (Formal, Friendly,
  Short, Detailed, In Arabic, In English) or edit what the user wrote
  (including To Arabic and To English), and a Draft button that
  writes from typed instructions. Every result is a draft card: Insert,
  Replace my text, Regenerate, Discard. Jameel Copilot writes as the user, answers the
  message being replied to instead of rewriting it, adds no signature, and
  never touches the signature or the quoted thread.
- **Prioritize my inbox**: recent unread mail ranked "Needs you", "Can wait",
  "For information", with a reason each; Reply on a row writes a reply card.
- **Tidy my inbox**: proposed "mark as read" and "flag" for unread mail; the
  user ticks and applies, and Undo puts it back.
- **Start my day**: one card with what the recent mail asks of the user,
  topics worth knowing, and today's meetings from the calendar.
- **Action items** of the selected email, each with "Add to tasks" (an
  Outlook task, made when the user presses it).
- **Waiting for a reply**: emails the user sent that nobody answered and
  that asked for something; Follow up writes a follow-up card.
- **Calendar**: Today's agenda, Prepare for my next meeting, and Schedule a
  meeting (common free times; a click opens Outlook's meeting window). Names
  the user types are looked up in the address book, and meeting rooms are
  listed with whether they are free.
- **Rules**: asked to always file or label a kind of mail, Jameel Copilot proposes an
  Outlook rule as a card; the user creates it. Such a rule can only move
  mail to one of the user's folders or give it a category.
- **My prompts**: requests the user saves as chips of their own.
- **Review** in a compose window: points about the draft before sending; the
  user picks which to apply. An edit chip with text selected works on that
  text only.
- **Emails used**: under an answer, the emails Jameel Copilot found or read; a click
  opens one in Outlook.
- **Chats** are saved in the user's own mailbox: a list to reopen, delete or
  clear; a reply window gets its thread's chat back.
- **Memory**: notes the user adds, changes and deletes ("Reply to HR in
  Arabic"). They go with every request. In the chat, Jameel Copilot can propose a note;
  only the user saves it.
- **Excel and PDF export**, saved to `Documents\Jameel Copilot\Reports`.
- **About Jameel Copilot** (the gear next to the Jameel Copilot mark on the ribbon; both buttons show
  no text): version, AI service host, model, and a connection test.

The model works through 23 tools: 16 that read, search, count and export
mail and read the calendar and the address book, three that only propose (a
memory note, a meeting, a rule) for the user to act on, and 4 "safe writes"
(create a draft, mark read, flag, set category) that are off unless an
administrator enables them. There is no tool that sends, deletes or moves
mail.

## Layout

| Path | What it is |
|---|---|
| `VSTO2/JameelCopilot/` | The add-in (C# 7.3, .NET Framework 4.7.2). Assembly name and namespace `JameelCopilot`. |
| `VSTO2/JameelCopilot.Tests/` | xUnit tests |
| `Installer/` | The MSI (WiX 5), its custom action, `Build-Msi.ps1`, `Set-JameelCopilotConfig.ps1` |
| `tools/` | `Build.ps1`, `Probe-CustomApi.ps1`, `Probe-Language.ps1` |
| `Deploy/README.txt` | Deployment guide for IT |
| `docs/superpowers/` | Design and implementation plan |

Inside the add-in:

- `Services/CustomApiChatService.cs`: the only code that makes network
  requests. `ChatCompletionsWire.cs` and `ChatStreamReader.cs` hold the wire
  format.
- `Prompts/*.txt`, `Services/SystemPromptBuilder.cs`: the system prompt. A
  fixed core with the language rules first, the pane's own text, then what
  the pane read from the mailbox, marked as data.
- `Services/ResponseStyles.cs`, `Services/Prefs/`: the user's response style
  (Precise, Balanced, Creative), kept in a hidden item in their mailbox.
- `Services/MachineSettings.cs`, `Config.cs`: settings, read from
  `HKLM\SOFTWARE\JameelCopilot` only. `Services/ApiKeyProtector.cs`: the stored
  key's format.
- `Services/OutlookToolHost.cs`, `Services/Tools/`: the Outlook tools.
- `Services/Export/`: Excel (ClosedXML) and PDF (off-screen WebView2).
- `TaskPane/`: the panes. `TaskPane/Chat/WebView2Bootstrap.cs` locks down every
  WebView2.
- `WebUI/`: the chat page and the print page.

## Security model in short

- **One destination.** Requests go to `{ApiBaseUrl}/chat/completions` over
  HTTPS; redirects are not followed. A test fails the build if another address
  literal appears in the add-in (`EgressGuardTests`).
- **Settings are an administrator's.** They live in HKLM. The add-in reads
  them and cannot write them; there is no settings editor and no password.
- **The key** is stored encrypted with DPAPI for the machine. That keeps it
  out of plain text; it does not hide it from users of that PC. The real
  limits belong on the Custom API (see the deployment guide).
- **Model output is untrusted.** A crafted email can influence what the model
  says. The pages that show it carry a Content-Security-Policy, cannot navigate
  or open windows, and may load only the add-in's own files; links are shown as
  text, not anchors. Web messages are accepted only from the add-in's own page.
- **Limited actions.** No send, delete or move; write tools off by default;
  "Open" works only on `.xlsx` and `.pdf` files in the Reports folder.
- **Logs** hold timings, tool names and error codes, never mail content,
  prompts or the key.

## Build and test

Needs Visual Studio or Build Tools with the **Visual Studio Tools for Office
(VSTO) build tools** component (in Build Tools it is a *recommended* part of
the "Office/SharePoint build tools" workload, not installed by default), and
the .NET SDK. Office does not have to be installed.

```powershell
# Once: a certificate to sign the add-in's manifests on this machine.
$thumb = (New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Jameel Copilot dev" `
          -CertStoreLocation Cert:\CurrentUser\My).Thumbprint

# Build the add-in and the tests, and run the tests.
.\tools\Build.ps1 -CertThumbprint $thumb -Test

# One test class.
.\tools\Build.ps1 -CertThumbprint $thumb -Test -Filter "FullyQualifiedName~RibbonTests"
```

`tools\Build.ps1` builds the add-in with MSBuild and the tests with the .NET
SDK; its header explains why two tools are needed.

## Build the MSI

```powershell
.\Installer\Build-Msi.ps1 -Version 1.0.0 -CertThumbprint $thumb
```

It rebuilds the add-in in Release from clean, runs the tests, builds
`out\JameelCopilot-1.0.0-x64.msi` and prints its SHA256. The version given is the MSI's,
the DLL's and the one About shows; a higher version upgrades an installed
lower one. The add-in DLL, its manifests, the installer's custom action and
the MSI are all signed with the certificate. For a release, use the
company's code-signing certificate: machines with AppLocker or WDAC will block
a self-signed build.

## Run a developer build in Outlook

On a PC with Outlook, from an elevated PowerShell:

```powershell
.\Installer\Set-JameelCopilotConfig.ps1 -ApiBaseUrl https://llm.example.internal/v1 -Model <model id>
```

It writes the same registry values the MSI writes and asks for the key
(typed hidden), or reads it from a git-ignored `jameelcopilot.local.json`
(`{ "apiKey": "..." }`) in the repository root. Then install an MSI built
from your branch without `API_KEY_FILE`: it keeps the settings already there.

`tools\Probe-CustomApi.ps1` asks the model for two tool calls, sends the results
back, and prints what the Custom API answers. Use it to check that a model
calls tools the way Jameel Copilot expects.

`tools\Probe-Language.ps1` sends the cases in `tools\language-cases.json` under
Jameel Copilot's own system prompt (`Prompts\core-agent.txt`) and measures whether an
Arabic question gets a wholly Arabic answer and an English one an English
answer. It also reports whether the model accepts `temperature` and
`max_tokens`. Run both on the company network before a rollout; each ends
with `RESULT: OK` or names what failed.

## Rules for changes

- Every behaviour change comes with a test that fails without it.
- Tests and docs use fake data only (`example.com`, `llm.example.internal`).
  The real API address and key never enter the repository.
- The add-in project is old-style: add each new `.cs` file to
  `JameelCopilot.csproj`.
- A new NuGet assembly must be referenced with the version its DLL really
  has, or MSBuild drops it from the output without a warning
  (`ProjectReferenceTests`).
