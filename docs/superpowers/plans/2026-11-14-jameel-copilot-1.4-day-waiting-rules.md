# Jameel Copilot 1.4: people and rooms, Start my day, waiting for a reply, tasks, translate, my prompts, rules

**Goal:** fix what the owner saw when scheduling (names left as typed, the rooms question unanswered) and add the six features he picked from the 1.4 ideas list.

**Spec:** `docs/superpowers/specs/2026-11-13-jameel-copilot-agent-design.md` (rules that bind every part: Jameel Copilot never sends; a card's button acts on what the host kept, never on text from the page; structured answers keep only ids Jameel Copilot gave; the model can only propose).

## Global constraints

- One pane, one page. Every new screen is a card in the chat or a view of the page.
- No new network destination, no new dependency.
- The model changes nothing by itself. A task and a rule are created by the user's click on a card that shows exactly what will be created.
- A rule can only move mail to a folder that exists (never Deleted Items or Junk) or set a category. Never forward, redirect or delete.
- Sender-written text stays on one line in prompts; bodies go behind "> ".

## Parts

1. **People and rooms (the fix).** Tools `outlook_resolve_people` (names to address-book people, by Outlook's own name check) and `outlook_find_rooms` (the address book's room list with free/busy for the time). `jc_propose_meeting` takes a `room`; the meeting window adds it as a resource. The Inbox prompt tells the model to resolve names first, to try the other script, to look in the mail, then to ask.
2. **Start my day.** Chip, structured answer over the recent unread mail: to-dos and topics from the model, today's meetings from the calendar (no model). One card.
3. **Add to my tasks.** "Action items" becomes a structured card. Each to-do row (there and in the day card) has "Add to tasks": Jameel Copilot creates an Outlook task from the row it kept.
4. **Waiting for a reply.** Sent mail of the last 14 days with no reply in its conversation (`Conversation.GetChildren`); the model keeps the ones that ask for something. Each row has "Follow up": a draft card that opens Outlook's reply-all window.
5. **Translate.** Inbox chip "Translate" (Arabic to English, anything else to Arabic). Compose chips: reply "In Arabic" / "In English", and "To Arabic" / "To English" for the user's own text.
6. **My prompts.** The user's own chips, kept in the mailbox like memory notes (a second `UserMemoryService` on another item). A view to add, change and delete them.
7. **Rule proposals.** Tool `jc_propose_rule`; card "Create this rule?"; the button creates the Outlook rule.

## Review focus

- A name Outlook cannot resolve must be said to the user, not silently kept.
- A to-do or a rule is created once per click, from the host's copy.
- An unknown id in a structured answer is dropped.
- No room list in the address book: the model is told so and asks for the room's name.
- A store without conversations: "Waiting for a reply" says it found nothing rather than listing every sent email.

## Not verifiable here (needs Outlook on the company network)

Name resolution against the real address book, the room list's name, free/busy of rooms, `GetChildren` on the Sent folder, `ReplyAll` on a sent item, saving a rule. They go on the acceptance list in `Deploy/README.txt`.
