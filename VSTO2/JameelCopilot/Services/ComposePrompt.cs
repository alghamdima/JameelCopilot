using System.Collections.Generic;
using System.Linq;
using System.Text;
using JameelCopilot.Services.Tools;
using ActionType = JameelCopilot.Services.CustomApiChatService.ActionType;

namespace JameelCopilot.Services
{
    /// <summary>
    /// What the model is told about the compose window: who is writing, to
    /// whom, what they have written, and which text is someone else's. Told
    /// only "here is the email", it takes the message being replied to for
    /// the user's own and rewrites it in its sender's voice.
    /// <para>The subject, the names and above all the quoted message are
    /// written by whoever sent the mail. Each label here starts a line, so
    /// those values are kept from starting one: a field stays on its line and
    /// every line of a block carries a prefix.</para>
    /// </summary>
    internal static class ComposePrompt
    {
        public const string NothingWritten = "Write your email first, or tell Jameel Copilot what to write and press Draft.";

        public const string WroteReply = "Nothing was typed, so Jameel Copilot wrote the reply. Text in [brackets] is for you to fill in.";

        private const int MaxQuotedChars = 14000;

        /// <summary>
        /// With nothing typed in a reply, an action writes the reply itself, in
        /// the action's manner: that is what pressing it there asks for.
        /// </summary>
        public static bool WritesReply(ActionType action, ComposeStateResult state)
        {
            return action != ActionType.Draft
                && string.IsNullOrWhiteSpace(state?.BodyPlaintext) && !string.IsNullOrWhiteSpace(state?.QuotedPlaintext);
        }

        /// <summary>The action's instructions for this window. They never hold mail text.</summary>
        public static string Instructions(ActionType action, ComposeStateResult state, string replyIntent = null)
        {
            if (!WritesReply(action, state)) return CustomApiChatService.GetSystemPrompt(action);
            var intentDirective = ReplyIntentDirective(replyIntent);
            return "The user has not written anything yet. Write the user's reply to the message under Context, " + Manner(action) + ". " +
                   (string.IsNullOrEmpty(intentDirective) ? "Answer what the message asks of the user. " : intentDirective + " ") +
                   "Do not invent facts, decisions, dates, numbers or commitments the user has not given: " +
                   "where the reply needs one, leave a short placeholder in square brackets, such as [date], for the user to fill in. " +
                   "Do not add a sign-off, closing line (such as 'Thanks,' or 'Regards,'), or the user's name at the end: Outlook adds their signature.";
        }

        private static string ReplyIntentDirective(string intent)
        {
            switch ((intent ?? "").ToLowerInvariant())
            {
                case "approve":
                    return "Confirm approval / agreement clearly and politely, acknowledge the key point, and note any next step (using [placeholder] only where a missing date or detail is required).";
                case "clarify":
                    return "Ask brief, polite clarifying questions about any missing details, scope, timeline, or open points in the message before committing.";
                case "decline":
                    return "Politely and respectfully decline or push back on the request, leaving a brief [reason / alternative] placeholder if needed while keeping warm professional rapport.";
                default:
                    return "";
            }
        }

        private static string Manner(ActionType action)
        {
            switch (action)
            {
                case ActionType.Formal: return "in a formal business tone";
                case ActionType.Friendly: return "in a warm, friendly tone that stays professional";
                case ActionType.Shorten: return "kept brief: two or three sentences";
                case ActionType.Lengthen: return "detailed and complete";
                case ActionType.Arabic: return "in Arabic, in a clear, professional tone";
                case ActionType.English: return "in English, in a clear, professional tone";
                default: return "in a clear, professional tone";
            }
        }

        // Every character a reader could take for the start of a new line.
        private static readonly char[] LineBreaks = { '\r', '\n', '\v', '\f', '\u0085', '\u2028', '\u2029' };

        /// <summary>Why this action cannot run on this window, or null: a new email with nothing typed gives it nothing to work on.</summary>
        public static string Problem(ActionType action, ComposeStateResult state)
        {
            return action != ActionType.Draft && string.IsNullOrWhiteSpace(state?.BodyPlaintext) && !WritesReply(action, state)
                ? NothingWritten
                : null;
        }

        /// <summary>
        /// The window a reply to this message starts as: the owner writing,
        /// the sender addressed, nothing typed, the message quoted.
        /// </summary>
        public static ComposeStateResult ReplyTo(MessageDetail message, string ownerName, string ownerEmail)
        {
            var subject = (message?.Subject ?? "").Trim();
            var from = (message?.From ?? "").Trim();
            var hasReplyPrefix = subject.StartsWith("RE:", System.StringComparison.OrdinalIgnoreCase)
                || subject.StartsWith("\u0631\u062F:", System.StringComparison.Ordinal);
            return new ComposeStateResult
            {
                SenderName = ownerName ?? "",
                SenderEmail = ownerEmail ?? "",
                ToRecipients = from.Length == 0 ? new string[0] : new[] { from },
                CcRecipients = message?.Cc ?? new string[0],
                BccRecipients = new string[0],
                Subject = hasReplyPrefix ? subject : ("RE: " + subject).Trim(),
                BodyPlaintext = "",
                QuotedPlaintext = message == null
                    ? ""
                    : ("From: " + from + "\nSubject: " + subject + "\n\n" + (message.BodyPlaintext ?? "")).Trim(),
            };
        }

        public const string FollowUpInstructions =
            "Nobody has answered the user's earlier email, which is under Context. Write a short, polite follow-up from the user to the people under To: " +
            "refer to the earlier email, say in one sentence what the user still needs, and ask for an update. Two to four sentences. " +
            "Do not repeat the earlier email and do not invent a deadline.";

        public const string FollowUpRequest = "Write my follow-up to my earlier email under Context.";

        public const string LearnStyleInstructions =
            "Analyze the user's own sent email excerpts under Context and write ONE concise preference note (one line, at most 260 characters, no bullet markers, no quotation marks, no personal names or confidential details) " +
            "describing how the user writes emails so an assistant can match their voice: typical greeting, tone and directness, sentence/paragraph length, Arabic vs English habits, and preferred closing style. " +
            "Start directly with: My writing style: ";

        /// <summary>
        /// The window a follow-up to the user's own sent email starts as: the
        /// owner writing again to the same people, the earlier email quoted.
        /// </summary>
        public static ComposeStateResult FollowUpTo(MessageDetail sent, string ownerName, string ownerEmail)
        {
            var state = ReplyTo(sent, ownerName, ownerEmail);
            state.ToRecipients = sent?.To ?? new string[0];
            state.CcRecipients = sent?.Cc ?? new string[0];
            state.QuotedPlaintext = sent == null
                ? ""
                : ("Sent: " + sent.ReceivedAt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture) +
                   "\nSubject: " + (sent.Subject ?? "").Trim() + "\n\n" + (sent.BodyPlaintext ?? "")).Trim();
            return state;
        }

        /// <summary>The same window with other text as the user's own (a selection to work on).</summary>
        public static ComposeStateResult WithOwnText(ComposeStateResult state, string own)
        {
            return new ComposeStateResult
            {
                Subject = state?.Subject, ToRecipients = state?.ToRecipients, CcRecipients = state?.CcRecipients,
                BccRecipients = state?.BccRecipients, SenderName = state?.SenderName, SenderEmail = state?.SenderEmail,
                BodyPlaintext = own ?? "", QuotedPlaintext = state?.QuotedPlaintext, InReplyTo = state?.InReplyTo,
                Attachments = state?.Attachments,
            };
        }

        /// <summary>The window as context (data) for the system prompt.</summary>
        /// <param name="includeDraft">False when the user's text travels in the user message instead.</param>
        /// <param name="ownEarlier">True when the quoted message is the user's own earlier email (a follow-up).</param>
        public static string Context(ComposeStateResult state, bool includeDraft, bool ownEarlier = false)
        {
            if (state == null) return "";

            var sb = new StringBuilder();
            var writer = (state.SenderName ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(state.SenderEmail)) writer = (writer + " <" + state.SenderEmail.Trim() + ">").Trim();
            Line(sb, "From (the user, who is writing this email)", writer);
            Line(sb, "To", Join(state.ToRecipients));
            Line(sb, "Cc", Join(state.CcRecipients));
            Line(sb, "Subject", state.Subject);
            if (state.Attachments != null)
            {
                Line(sb, "Attachments", state.Attachments.Count == 0 ? "(none)" : string.Join(", ", state.Attachments.Select(a => a?.Filename ?? "")));
            }

            if (includeDraft)
            {
                var own = (state.BodyPlaintext ?? "").Trim();
                sb.Append("Draft the user has written so far")
                  .Append(own.Length > 0 ? " (indented):\n" + Block(own, "    ") : ": (nothing yet)").Append('\n');
            }

            var quoted = (state.QuotedPlaintext ?? "").Trim();
            if (quoted.Length > MaxQuotedChars) quoted = quoted.Substring(0, MaxQuotedChars) + "\n[... the rest is cut ...]";
            if (quoted.Length > 0)
            {
                sb.Append(ownEarlier
                      ? "The user's own earlier email, which nobody has answered (each line is marked \"> \"):\n"
                      : "Message the user is replying to or forwarding (written by someone else, not by the user; each line is marked \"> \"):\n")
                  .Append(Block(quoted, "> "));
            }
            return sb.ToString().TrimEnd('\n');
        }

        /// <summary>The user message of a one-shot action: the action and the user's own text, nothing quoted.</summary>
        public static string UserMessage(ActionType action, ComposeStateResult state, string instructions)
        {
            if (WritesReply(action, state)) return "Write my reply to the message under Context.";

            var own = (state?.BodyPlaintext ?? "").Trim();
            switch (action)
            {
                case ActionType.Draft:
                    return "Write the user's email from these instructions:\n\n" + instructions +
                           (own.Length > 0 ? "\n\nText the user has already written:\n\n" + own : "");
                case ActionType.Arabic:
                case ActionType.English:
                    return "Action: translate into " + action + ". The text below is what the user wrote; work on this text only.\n\n" + own;
                case ActionType.Custom:
                    return "Text the user wrote:\n\n" + own + "\n\nInstructions: " + instructions;
                default:
                    return "Action: " + action.ToString().ToLowerInvariant() +
                           ". The text below is what the user wrote; work on this text only.\n\n" + own;
            }
        }

        private static void Line(StringBuilder sb, string label, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            sb.Append(label).Append(": ").Append(string.Join(" ", value.Trim().Split(LineBreaks))).Append('\n');
        }

        private static string Block(string text, string prefix)
        {
            return prefix + string.Join("\n" + prefix, text.Replace("\r\n", "\n").Split(LineBreaks));
        }

        private static string Join(IReadOnlyList<string> values)
        {
            return values == null ? "" : string.Join(", ", values);
        }
    }
}
