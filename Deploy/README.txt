Jameel Copilot FOR OUTLOOK - DEPLOYMENT GUIDE
=================================

For the IT team that installs and maintains Jameel Copilot on company PCs.


WHAT Jameel Copilot IS
----------
Jameel Copilot is an add-in for classic Outlook on Windows. It adds an assistant that can
read, search and summarise the user's mail, help write replies, and export
results to Excel or PDF.

Jameel Copilot sends mail content to one place only: the company's Custom API, at the
address you give the installer. Jameel Copilot itself contacts no other service (see WHAT
Jameel Copilot SENDS AND WHERE for the Microsoft component it uses). It cannot send,
delete or move mail.


REQUIREMENTS ON EACH PC
-----------------------
  - 64-bit Windows 10 or Windows 11
  - Classic Outlook 2016 or later, 32-bit or 64-bit (tested target: Office
    LTSC 2024). Jameel Copilot does not run in "new Outlook" or Outlook on the web.
  - .NET Framework 4.7.2 or later (part of current Windows)
  - Microsoft Edge WebView2 Runtime
  - Visual Studio 2010 Tools for Office Runtime (installed with Office)
  - Network access from the PC to the Custom API over HTTPS, with a
    certificate the PC trusts

The installer checks Windows, .NET and WebView2 and stops with a message if
one is missing. It installs none of them and downloads nothing.

It does not check the Tools for Office Runtime: Office keeps its own copy
where an installer cannot see it. See TROUBLESHOOTING if Jameel Copilot does not appear.


INSTALL
-------
One file: JameelCopilot-<version>-x64.msi. Install it per machine, as administrator or
through your deployment tool (Intune, Configuration Manager, Group Policy).

  msiexec /i JameelCopilot-1.0.0-x64.msi /qn ^
      API_BASE_URL=https://llm.example.internal/v1 ^
      API_KEY_FILE=jc-key.txt ^
      MODEL=<model id>

Required on the first install:

  API_BASE_URL   The Custom API base address. Must begin with https://
  API_KEY_FILE   The name of a text file next to the MSI that holds the key
                 Jameel Copilot uses, and nothing else (spaces and line
                 breaks around it are ignored). A full path works too.
                 Stored encrypted; see THE API KEY. If the file is missing
                 or empty, the install stops (exit code 1603).
  MODEL          The model ID sent with every request. The model must
                 support tool calling, or the mailbox features will not work.

Optional:

  WRITE_TOOLS    Comma-separated list of the actions the assistant may take
                 in the mailbox. Default: none (read-only). Possible values:
                   outlook_create_draft    create a draft (never sends it)
                   outlook_mark_as_read    mark a message read or unread
                   outlook_flag_message    set or clear a follow-up flag
                   outlook_set_category    set a message's category
  REASONING_EFFORT
                 Passed to the model as "reasoning_effort" (for example low).
                 Leave it out unless your model uses it.
  MAX_BULK_EXPORT_ROWS
                 Largest number of rows in a "complete list" Excel export.
                 Default 2000, at most 10000.
  TEMPERATURE    How varied the answers are, from 0 to 2. Default 0.5. This
                 is the value of the user's "Balanced" response style. Write
                 it with a dot (0.3). The word  off  sends no temperature at
                 all, for a model that rejects it; users then get no style
                 choice.
  TEMPERATURE_MAX
                 The value of the user's "Creative" style, and the highest
                 temperature Jameel Copilot ever sends. Default 0.9.
  MAX_OUTPUT_TOKENS
                 Longest answer, in tokens. Default: no limit is sent. An
                 answer that reaches the limit stops there and ends with a
                 line that says so. To remove a limit that was set, install
                 again with MAX_OUTPUT_TOKENS=0: a setting that is left out
                 keeps its last value.

Each user picks a response style in the Jameel Copilot pane: Precise (0.2), Balanced or
Creative. The choice stays within the two values above.

Deploy while Outlook is closed (at logon, or in a maintenance window). The
installer never closes Outlook. After a first install, Jameel Copilot appears the next
time Outlook starts: an "Jameel Copilot" group on the Home tab and in new-message windows.

To keep a log, add:  /l*v install.log
The key is masked in that log.

Never put the key itself on the command line (API_KEY=<key> is still
accepted, for a test PC). Whatever records command lines records it in plain
text: your deployment tool's package definition and its logs, and Windows
process-creation auditing (event 4688) or an EDR agent, which send it on to
your security monitoring. With API_KEY_FILE they record only the file's
name. Keep the file in the package's source folder, readable only by the
deployment team and the deployment tool's accounts, and rotate the key if it
was ever on a command line.


WHERE THINGS ARE
----------------
  Program files   C:\Program Files\JameelCopilot
  Settings        HKLM\SOFTWARE\JameelCopilot  (64-bit registry view)
  Per user        %LOCALAPPDATA%\JameelCopilot      page cache and trace.log
                  Documents\Jameel Copilot\Reports           exported Excel and PDF files
                  (%LOCALAPPDATA%\JameelCopilot\Reports when Documents is on a
                  network share)
  Per mailbox     hidden items in the user's own Inbox, which follow the
                  user to another PC and stay in your Exchange:
                    Jameel Copilot.Preferences   the response style
                    Jameel Copilot.Chats, Jameel Copilot.Chat.<id>   the user's saved Jameel Copilot chats: what
                                     they asked and what Jameel Copilot answered (up to
                                     40 chats; the user can delete them)
                    Jameel Copilot.Memory        the notes the user asked Jameel Copilot to remember

Only administrators can change the settings. Users see them, read-only, under
About Jameel Copilot (the gear button next to the Jameel Copilot mark on the ribbon), which also has
a "Test connection" button.


UPGRADE
-------
Install the newer MSI the same way. Settings you do not pass again are kept:

  msiexec /i JameelCopilot-1.1.0-x64.msi /qn

The installer removes the old version and installs the new one in one step.
Jameel Copilot > About Jameel Copilot then shows the new version number.

If Outlook is open during an upgrade, Windows finishes replacing the files at
the next restart. Upgrade while Outlook is closed.


CHANGE A SETTING OR ROTATE THE KEY
----------------------------------
Run the MSI that is already installed again, with the setting to change.
Everything you do not pass stays as it is.

  New key:    msiexec /i JameelCopilot-1.0.0-x64.msi /qn API_KEY_FILE=jc-key.txt
  New model:  msiexec /i JameelCopilot-1.0.0-x64.msi /qn MODEL=<model id>

Run it as administrator or through your deployment tool, silently (/qn) as
shown. Users pick up the change the next time they start Outlook: Jameel Copilot > About
Jameel Copilot shows the model and the address in use, and Test connection checks the key.

A standard user cannot change the settings this way: Windows Installer does
not accept these values from a user who is not an administrator.


DEPLOYING WITH CONFIGURATION MANAGER (SCCM)
-------------------------------------------
Jameel Copilot has no management console of its own: Configuration Manager is the console.

Install
  Create an Application from the MSI. Installation program: the silent
  command from INSTALL, with your address and model, and the key file in
  the Application's content folder next to the MSI. Install for system,
  whether or not a user is logged on.

Detection
  Do not detect by MSI product code: every build of Jameel Copilot has a new one. Use
  both of these:
    - registry value HKLM\SOFTWARE\JameelCopilot\Model exists (64-bit registry);
    - file C:\Program Files\JameelCopilot\JameelCopilot.dll has the version of the
      build, for example 1.1.0.0.

Upgrade
  Add the new MSI as a new Application that supersedes the old one, with
  "Uninstall" left unticked: the new MSI removes the old version itself and
  keeps the settings. Give it the full command, settings included, as for a
  first install. A PC that does not have Jameel Copilot yet gets this Application too,
  and without the settings the install stops there.

Change a setting for everyone
  Configuration Manager runs an Application's command only on a PC where
  detection says Jameel Copilot is not installed. Editing the command changes nothing on
  the PCs that already have Jameel Copilot. Use one of these instead:
    - a Package whose Program runs the installed version's MSI again with
      the one property, for example
          msiexec /i JameelCopilot-<version>-x64.msi TEMPERATURE=0.3 /qn
      deployed as Required with the rerun behavior "Always rerun program";
    - a configuration baseline that sets the value directly. The values are
      strings under HKLM\SOFTWARE\JameelCopilot (64-bit registry):
          Model, ReasoningEffort, EnabledWriteTools, MaxBulkExportRows,
          Temperature, TemperatureMax, MaxOutputTokens
      Leave ApiBaseUrl to the installer, which checks it, and never write
      ApiKeyProtected by hand.
  Jameel Copilot reads the values when Outlook starts. Then put the new value in the
  Application's command as well, for the PCs that get Jameel Copilot later.

Rotate the key
  Only through the MSI, and for the same reason not by editing the
  Application. Deploy a Package that holds the MSI and a file with the new
  key, whose Program is
      msiexec /i JameelCopilot-<version>-x64.msi API_KEY_FILE=jc-key.txt /qn
  to the PCs that have Jameel Copilot (Required, "Always rerun program"), then put the
  new key in the Application's key file for the PCs that get Jameel Copilot later. The
  key is encrypted separately on each PC, so a value copied from one PC does
  not work on another.

Keep Jameel Copilot enabled
  Outlook disables an add-in that makes it start slowly, and a user who
  does not notice the yellow bar is left without Jameel Copilot. To exempt Jameel Copilot for
  everyone, set the Outlook policy "List of managed add-ins" (User
  Configuration > Administrative Templates > Microsoft Outlook 2016 >
  Miscellaneous): name JameelCopilot, value 1 (always enabled). The name is
  case sensitive. As a registry setting for the user, deployed by a
  configuration baseline or a logon script:
      HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList
      JameelCopilot = "1"   (string)
  Microsoft describes it under "Support for keeping add-ins enabled". The
  MSI does not set it: it is a per-user policy and yours to decide.

The key in SCCM
  The client logs under C:\Windows\CCM\Logs on every PC record the command
  line (AppEnforce.log for an Application, execmgr.log for a Package), so
  the key goes in the key file, never on the command line. The file sits in
  the content source, on the distribution points and in each PC's client
  cache: limit who can read the source folder, and make the key a restricted
  one (see THE API KEY): for Jameel Copilot only, with a budget and a rate limit,
  accepted only from the company network.


UNINSTALL
---------
  msiexec /x JameelCopilot-1.0.0-x64.msi /qn

or Settings > Apps > Jameel Copilot. This removes the program files, the add-in
registration and the settings, including the stored key. The hidden items in
each mailbox (preferences, saved chats, memory) stay there: they are the
user's and the uninstaller does not open mailboxes. Each user's reports
and the folder under %LOCALAPPDATA% stay.


THE API KEY
-----------
The key is stored in the registry encrypted with Windows data protection for
that machine. The stored value is not plain text on disk, in backups or in
registry exports, and it cannot be decrypted on another PC. (The command line
that installed it is a separate matter: see INSTALL.)

This does not hide the key from the people who use that PC: any program a
user runs there can ask Windows to decrypt it, as Jameel Copilot itself does. A key that
Jameel Copilot can use is a key its users can obtain. Treat it accordingly, and put the
real controls on the Custom API:

  - issue a key for Jameel Copilot only, not a shared or administrative key;
  - limit that key to the model(s) Jameel Copilot uses;
  - give it a budget and a rate limit;
  - accept it only from the company network;
  - rotate it on your normal schedule (see above).

Jameel Copilot sends the Windows account name with each request (the "user" field), so
the Custom API's logs show who made each call.

If you need each user to be identified without a shared key, put a gateway
that uses Windows authentication in front of the Custom API. Jameel Copilot would need a
small change to use it; ask the developers.


WHAT Jameel Copilot SENDS AND WHERE
-----------------------
To the Custom API: the user's question, the mail content the assistant reads
to answer it, the Windows account name, and the PC's date and time zone. In
a compose window that content is the name and address the email is written
from, its recipients and subject, what the user has typed and the message
being replied to. The user's signature is not sent. The notes under Memory
go with every question, and so do the earlier questions and answers of the
chat that is open. For calendar questions and meeting times, the user's
meetings (subject, time, place, attendees) and the free/busy times of the
people a meeting is for, as Outlook shows them, are sent too. So are the
name and address of a person the user names for a meeting, as the address
book has them, and the names of meeting rooms with whether they are free.
Prioritize, Tidy and Start my day send the sender, subject, date and first
lines of the recent unread emails. Waiting for a reply sends the same of the
user's own sent emails of the last 14 days that nobody answered. Action
items sends the selected email.

Changes to mail: Jameel Copilot marks emails read or flags them only when the user
presses Apply on "Tidy my inbox", and Undo reverses it. That is the user's
own action and does not depend on WRITE_TOOLS, which governs what the model
may do by itself. Jameel Copilot never sends, deletes or moves mail; meeting requests
and replies open in Outlook's own window for the user to send.

Tasks and rules: Jameel Copilot adds a task to the user's Outlook tasks only when the
user presses "Add to tasks" on a to-do. Jameel Copilot creates an Outlook rule only
when the user presses "Create rule" on a rule card that shows the whole
rule. A rule made this way can only move incoming mail to one of the
user's own folders (never Deleted Items or Junk) or give it a category. It
cannot forward, redirect, copy or delete mail. The user sees and removes it
under Rules in Outlook like any other rule.

Nowhere else. Jameel Copilot has no update check, no telemetry and no sign-in service.
The chat is drawn by Jameel Copilot's own local pages, which are blocked from loading
anything from, or sending anything to, any network address.

Not Jameel Copilot's traffic, but you will see it in a capture of that PC:

  - The WebView2 Runtime (msedgewebview2.exe) is a Microsoft component that
    Jameel Copilot uses to draw the chat. While a Jameel Copilot pane is open, the runtime itself
    contacts Microsoft for its own configuration, updates and diagnostic
    data, as it does inside every application that uses it. In our test it
    opened a few HTTPS connections to Microsoft addresses in the first
    minute. Jameel Copilot's pages cannot add mail content to them. Jameel Copilot switches off the
    two parts an application is allowed to switch off (crash-dump upload and
    SmartScreen lookups). The rest is yours to control: the Windows
    "Diagnostic data" policy, your Edge update policies, or a firewall rule
    for msedgewebview2.exe. Jameel Copilot's pages are local files and need none of it.
  - Windows may check the Custom API's certificate (revocation and chain
    downloads), and Jameel Copilot's requests use the PC's system proxy if one is set.
  - Outlook's own connections (Exchange, Office services) come from the
    same OUTLOOK.EXE process that Jameel Copilot runs in.


TROUBLESHOOTING
---------------
"Outlook disabled the add-in 'Jameel Copilot'. It caused Outlook to start slowly."
  Outlook measures how long each add-in takes while it starts and disables
  one that takes over a second (the median of five starts). Jameel Copilot does no work
  of its own at start since 1.2.1, but on a PC where the .NET runtime and
  the other add-ins load slowly it can still cross the line. Since 1.2.3 Jameel Copilot
  names itself in the user's DoNotDisableAddinList when it starts, as the
  user's "Always enable this add-in" does; a policy that blocks Jameel Copilot still
  wins. A Jameel Copilot that was disabled before it could do that has to be enabled
  once by hand:
  For one user: press "View More Details..." on the yellow bar (or File >
  Slow and Disabled COM Add-ins) and choose "Always enable this add-in".
  For everyone: see "Keep Jameel Copilot enabled" under DEPLOYING WITH CONFIGURATION
  MANAGER. To see the time Outlook measured: Event Viewer > Windows Logs >
  Application, source Outlook, event 45 lists each add-in's "Boot Time" in
  milliseconds.

Jameel Copilot is not on the ribbon
  Outlook > File > Options > Add-ins. If Jameel Copilot is under "Disabled", enable it
  under Manage: Disabled Items. If it is under "Inactive" with a load error,
  the Tools for Office Runtime is missing or damaged: install "Visual Studio
  2010 Tools for Office Runtime" from Microsoft, or repair Office.
  Check that the user is in classic Outlook, not "new Outlook".

"Jameel Copilot is not configured. Contact IT."
  A required setting is missing or invalid. Jameel Copilot > About Jameel Copilot names it. Install
  the MSI again with the right value.

"The AI service returned an error (HTTP 400)" on every question after 1.1
  The model may reject the "temperature" (or "max_tokens") setting. Run
  tools\Probe-Language.ps1: its last lines say which one. Then set
  TEMPERATURE=off or MAX_OUTPUT_TOKENS=0 (leaving the property out keeps
  its last value), and restart Outlook.

The assistant mixes Arabic and English in one answer
  Run tools\Probe-Language.ps1 with -ShowReplies and send the output to the
  developers: it shows which kind of question the model answers badly.

The response style goes back to Balanced after Outlook restarts
  The mailbox could not keep Jameel Copilot's hidden item (some shared or file-only
  profiles). The choice still works until Outlook closes.

A setting did not change after running the MSI again
  Check that the command ran as administrator and ended with exit code 0,
  and that Outlook was restarted. Exit code 1603 with "API_BASE_URL is
  required" means the run was started by a standard user with a user
  interface: run it silently as administrator. If it still does not change,
  add  REINSTALL=ALL REINSTALLMODE=omus  to the command and tell the
  developers.

"Jameel Copilot cannot reach the AI service."
  The PC cannot connect to the Custom API: name resolution, firewall, proxy,
  or a certificate the PC does not trust. Open the address in a browser on
  that PC to see which.

"Jameel Copilot could not sign in to the AI service. Contact IT."
  The Custom API refused the key (expired, revoked, or not allowed to use the
  model). Rotate the key.

"The AI service is busy."
  The key's rate limit or budget was reached.

The chat area says Jameel Copilot needs the WebView2 runtime
  Install the Microsoft Edge WebView2 Runtime (Evergreen) on that PC.

The assistant answers but never searches or reads mail
  The model does not support tool calling. Set MODEL to one that does.

Log for the developers: %LOCALAPPDATA%\JameelCopilot\trace.log. It holds timings,
tool names and error codes. It does not hold mail content, questions or the
key. It is rewritten each time Outlook starts.


ACCEPTANCE CHECKLIST
--------------------
Run once on a clean PC with Outlook, on the company network, before a wide
rollout. Each line should be true.

  [ ] The silent install command finishes with exit code 0.
  [ ] Installing with API_BASE_URL=http://... is refused (exit code 1603) and
      nothing is installed.
  [ ] An install log (/l*v) does not contain the key.
  [ ] Installing with API_KEY_FILE naming a file that does not exist is
      refused (exit code 1603).
  [ ] Outlook starts with no trust or security prompt about Jameel Copilot.
  [ ] After five starts of Outlook there is no yellow bar saying Jameel Copilot was
      disabled, and event 45 (Application log, source Outlook) shows Jameel Copilot's
      boot time. If it is near 1000 ms, set the policy under "Keep Jameel Copilot
      enabled" before a wide rollout.
  [ ] The Home tab has a Jameel Copilot group with two buttons and no text: the Jameel Copilot
      mark (opens the pane) and a gear (About Jameel Copilot; its name shows on
      hover). There is one Jameel Copilot pane, whose header shows the mark alone.
  [ ] About Jameel Copilot shows the right address host and model; Test connection says
      "Connection OK."
  [ ] A question in the Jameel Copilot pane gets an answer.
  [ ] "How many unread emails do I have?" makes the assistant search the
      mailbox (a short "Counting messages" line appears) and answer.
  [ ] tools\Probe-Language.ps1, run on this PC, ends with "RESULT: OK".
  [ ] A question in Arabic gets an Arabic answer, aligned to the right, with
      no English sentences in it; names and addresses stay as written.
  [ ] Reply to an email, type nothing, open Jameel Copilot and press Formal: a draft
      card shows your answer to the sender (not a new version of their
      message), with no name or signature at the end. Insert puts it above
      your signature; the signature (picture and formatting) and the quoted
      message below it are unchanged, and the card says "in the email".
  [ ] In a new email with nothing typed, Proofread asks you to write your
      email first.
  [ ] In a reply, type a short instruction and press Draft: the card holds
      your answer written from that instruction.
  [ ] Type two lines in a reply, press Formal, then Replace my text: only
      your two lines change.
  [ ] In the Inbox, select an email and press Draft a reply, then Open
      reply: Outlook's reply window opens with the text above your
      signature. Nothing is sent.
  [ ] Ask Jameel Copilot something, close Outlook and open it again the same day: the
      chat is there. Chats lists it; Delete removes it.
  [ ] Reply to an email, ask Jameel Copilot about it, close the reply window, press
      Reply on the same email again and open Jameel Copilot: that chat is back.
  [ ] Memory: add "Reply in one short paragraph." and ask a question: the
      answer follows it. Edit and Delete work. After restarting Outlook the
      note is still there.
  [ ] Prioritize my inbox lists recent unread mail under "Needs you", "Can
      wait" and "For information"; a row opens the email.
  [ ] Tidy my inbox: untick one row, Apply: the others change (read or
      flagged) and Undo puts them back.
  [ ] In a reply with text typed, Review lists points; Apply selected makes
      a draft card with those changes. Select one sentence and press
      Shorten: the card says "selected text" and Replace selection changes
      only that sentence.
  [ ] Today's agenda lists today's meetings with the right times (check one
      repeating meeting). Schedule a meeting on an email shows times; a
      time opens Outlook's meeting window, filled in, and nothing is sent.
  [ ] Ask "Set up a meeting with <two colleagues' first names> tomorrow":
      the card shows their addresses, and the meeting window opens with the
      people resolved (underlined), not as typed. Ask with a name that two
      people share: Jameel Copilot asks which one.
  [ ] Ask "Which meeting rooms are free tomorrow at 10?": rooms are listed
      with free or busy. Ask for the meeting in one of them: the meeting
      window has the room as its location. (If no room is listed, tell the
      vendor the name of the room list in the address book.)
  [ ] Start my day shows To do, Catch up and Today's meetings; the meetings
      match the calendar. A to-do's line opens its email.
  [ ] Action items on an email that asks something: "Add to tasks" on a row
      adds one task to Outlook Tasks, with the due date when the email
      names one. Pressing it again adds no second task.
  [ ] Waiting for a reply lists emails you sent in the last 14 days that
      got no answer (not ones that were answered). Follow up writes a
      card; Open follow-up opens a reply to the people the email went to,
      and nothing is sent.
  [ ] Translate on an Arabic email gives English, and on an English email
      gives Arabic. In a reply, In Arabic writes the reply in Arabic.
  [ ] My prompts > + Add: save "Summarize what [name] sent me this week."
      It becomes a chip that waits in the input. After restarting Outlook
      it is still there.
  [ ] Ask "Always move mail from <a sender> to <one of your folders>": a
      rule card appears and nothing exists under Rules until Create rule.
      After it, the rule is under Rules and Alerts and only moves mail.
  [ ] In the chat, "Remember that I am in the Finance team" shows a card
      "Remember this?"; the note appears under Memory only after Save.
  [ ] Reports > With a person puts its text in the input and waits for you
      to replace [name or email]; Week digest runs at once.
  [ ] The Jameel Copilot button on the ribbon shows the teal Abdul Latif Jameel emblem with clear
      corners (no white or black box around it).
  [ ] Choosing Creative in the pane, closing Outlook and opening it again
      still shows Creative.
  [ ] About Jameel Copilot shows the mark, the temperature line and the right version.
  [ ] "Export my last 10 emails to Excel" produces a file that opens.
  [ ] With WRITE_TOOLS left out, asking the assistant to create a draft does
      not create one.
  [ ] Installing the same MSI again with only MODEL=<other> changes the
      model in About Jameel Copilot and keeps the address and the key (Test connection
      still says "Connection OK.").
  [ ] Installing the same MSI again with only API_KEY_FILE=<a file with
      another valid key> keeps the address and the model, and Test
      connection still passes.
  [ ] The same command run by a standard user (not an administrator) does
      not change the model.
  [ ] Installing a higher version over it (build one with -Version 1.0.1)
      keeps every setting, Jameel Copilot still loads, and About Jameel Copilot shows 1.0.1.
  [ ] After uninstall, C:\Program Files\JameelCopilot and HKLM\SOFTWARE\JameelCopilot
      are gone.
  [ ] In a network capture during a chat, the only destination that is not
      Microsoft's or your mail server's is the Custom API host, reached from
      OUTLOOK.EXE. (msedgewebview2.exe may contact Microsoft addresses: see
      WHAT Jameel Copilot SENDS AND WHERE.)
