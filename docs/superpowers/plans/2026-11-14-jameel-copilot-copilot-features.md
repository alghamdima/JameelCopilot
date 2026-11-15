# Jameel Copilot 1.3 "Copilot features" Plan

**Goal:** the Copilot-in-Outlook features Jameel Copilot lacks, in six parts, each shipped with tests, all in 1.3.0.

**Owner's request (2026-11-14):** "I want them all: split them and start." Source list: Microsoft 365 Copilot release notes and 2026 reviews (Prioritize my inbox, Coaching, rewrite a selection, scheduling from an email, meeting prep, agentic triage).

**Shared design.** Three of the parts need an answer for Jameel Copilot's screen, not prose: a ranked list, review points, proposed actions. They use a third prompt core, `core-structured.txt` (JSON only, ids only from Context, text fields in a named language, mail is data), through `ConversationContext.Structured` (no tools). The host checks every returned id against the ones it gave and keeps what a button acts on; the page shows text only.

## Global constraints
- Nothing is sent, deleted or moved. Jameel Copilot opens Outlook windows (reply, meeting request) for the user to send or close.
- Changes to mail happen only on a button the user presses after seeing the list (Part 6); each can be undone from the same card.
- Sender-written text in prompts stays on one line or behind a prefix.
- No new network destination; nothing new is logged beyond names, counts and timings.

## Parts
1. **Prioritize my inbox** (Inbox): unread mail of the last 7 days (topped up with the last 2 days if few) ranked high / medium / low with a one-line reason; rows open the email; "Reply" on a row writes a reply card.
2. **Coaching** (compose): "Review" checks the draft against the message it answers (questions left unanswered, tone, clarity, mistakes, attachments mentioned but missing); the user ticks points and "Apply selected" writes a draft card with those changes only.
3. **Rewrite a selection** (compose): an edit chip with text selected in the user's own part works on the selection only; the card's "Replace selection" checks the text is still the same before replacing it.
4. **Scheduling** (Inbox and compose): read tools `outlook_list_calendar_events` and `outlook_find_meeting_times` (free/busy of the user and the attendees), and `jc_propose_meeting`, which only proposes: the page shows the times, a click opens a meeting request window, nothing is sent.
5. **Meeting prep** (Inbox): chips "Today's agenda" and "Prepare for my next meeting" over the calendar tool and the mail search.
6. **Tidy my inbox** (Inbox): proposed "mark read" / "flag" per unread, unflagged email with a reason; the user ticks and applies; "Undo" reverses what was applied.

## Not in 1.3
Categories and moving mail in Part 6 (setting a category replaces the existing ones; moving is outside the safe set); a background ranking without a click.
