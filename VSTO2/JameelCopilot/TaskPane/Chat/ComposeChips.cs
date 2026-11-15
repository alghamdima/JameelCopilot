using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ActionType = JameelCopilot.Services.CustomApiChatService.ActionType;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// The chips of a compose window. Each one runs an action whose result is
    /// a draft card: the user decides whether it goes into the email.
    /// </summary>
    internal static class ComposeChips
    {
        /// <summary>The chip that reviews the user's draft (coaching).</summary>
        public const string ReviewChipId = "review";

        private const string ReplyPrefix = "reply:";
        private const string EditPrefix = "edit:";

        /// <summary>The chip groups, as the page takes them. A new email has no reply to write.</summary>
        public static JArray Groups(bool isReply)
        {
            var groups = new JArray();
            if (isReply)
            {
                groups.Add(Group("Write my reply",
                    Chip("Formal", ReplyPrefix, ActionType.Formal, primary: true),
                    Chip("Friendly", ReplyPrefix, ActionType.Friendly),
                    Chip("Short", ReplyPrefix, ActionType.Shorten),
                    Chip("Detailed", ReplyPrefix, ActionType.Lengthen),
                    Chip("In Arabic", ReplyPrefix, ActionType.Arabic),
                    Chip("In English", ReplyPrefix, ActionType.English)));
            }
            var review = new JObject(new JProperty("label", "Review"), new JProperty("mode", "action"), new JProperty("id", ReviewChipId));
            if (!isReply) review.Add("primary", true);
            groups.Add(Group("Edit what I wrote", review,
                Chip("Proofread", EditPrefix, ActionType.Proofread),
                Chip("Improve", EditPrefix, ActionType.Revise),
                Chip("Shorten", EditPrefix, ActionType.Shorten),
                Chip("Lengthen", EditPrefix, ActionType.Lengthen),
                Chip("Formal", EditPrefix, ActionType.Formal),
                Chip("Friendly", EditPrefix, ActionType.Friendly),
                Chip("To Arabic", EditPrefix, ActionType.Arabic),
                Chip("To English", EditPrefix, ActionType.English)));
            return groups;
        }

        /// <summary>The action a chip stands for; false for an id that is not one of ours.</summary>
        public static bool TryParse(string chipId, out ActionType action)
        {
            action = ActionType.Custom;
            var name = chipId == null ? null
                : chipId.StartsWith(ReplyPrefix, StringComparison.Ordinal) ? chipId.Substring(ReplyPrefix.Length)
                : chipId.StartsWith(EditPrefix, StringComparison.Ordinal) ? chipId.Substring(EditPrefix.Length)
                : null;
            return name != null && Enum.TryParse(name, false, out action) && Enum.IsDefined(typeof(ActionType), action)
                && action != ActionType.Draft && action != ActionType.Custom && name == action.ToString();
        }

        /// <summary>The heading of the draft card.</summary>
        public static string DraftTitle(ActionType action, bool writesReply)
        {
            if (action == ActionType.Draft) return "Draft";
            if (action == ActionType.Custom) return "Draft with the review's changes";
            if (!writesReply) return Label(action);
            return action == ActionType.Proofread || action == ActionType.Revise ? "Reply draft" : "Reply draft - " + ReplyLabel(action);
        }

        /// <summary>
        /// The buttons of a finished draft card. With text of the user's in the
        /// email, the draft most likely stands in for it; with none, it goes in.
        /// </summary>
        public static JArray DraftButtons(bool hasOwnText)
        {
            var buttons = new JArray();
            if (hasOwnText)
            {
                buttons.Add(Button("Replace my text", "replace", primary: true));
                buttons.Add(Button("Insert above", "insert"));
            }
            else
            {
                buttons.Add(Button("Insert", "insert", primary: true));
            }
            buttons.Add(Button("Regenerate", "regenerate"));
            buttons.Add(Button("Discard", "discard"));
            return buttons;
        }

        /// <summary>The buttons of a reply written in the Inbox: it goes into Outlook's own reply window.</summary>
        public static JArray ReplyWindowButtons()
        {
            return new JArray(
                Button("Open reply", "reply", primary: true), Button("Reply all", "reply_all"),
                Button("Regenerate", "regenerate"), Button("Discard", "discard"));
        }

        /// <summary>The buttons of a follow-up to the user's own sent email: it goes to the people that email went to.</summary>
        public static JArray FollowUpButtons()
        {
            return new JArray(Button("Open follow-up", "reply_all", primary: true), Button("Regenerate", "regenerate"), Button("Discard", "discard"));
        }

        /// <summary>The page call that ends a draft card with its buttons.</summary>
        public static string FinalizeScript(string id, JArray buttons, string status, string note, bool error)
        {
            var opts = new JObject(new JProperty("status", status ?? ""), new JProperty("error", error));
            if (note != null) opts.Add("note", note);
            return "jameelcopilot.finalizeDraft(" + JsonConvert.SerializeObject(id ?? "") + ", " +
                   buttons.ToString(Formatting.None) + ", " + opts.ToString(Formatting.None) + ");";
        }

        /// <summary>The page call that says where a draft went, or that it could not go there.</summary>
        public static string MarkScript(string id, bool done, string where, string failed)
        {
            return "jameelcopilot.markDraft(" + JsonConvert.SerializeObject(id ?? "") + ", " +
                   JsonConvert.SerializeObject(done ? where : failed) + ", " + (done ? "true" : "false") + ");";
        }

        /// <summary>The buttons of a draft written for the selected text.</summary>
        public static JArray SelectionButtons()
        {
            return new JArray(Button("Replace selection", "replace_selection", primary: true),
                Button("Regenerate", "regenerate"), Button("Discard", "discard"));
        }

        /// <summary>The buttons of a draft that did not finish.</summary>
        public static JArray DiscardOnly()
        {
            return new JArray(Button("Discard", "discard"));
        }

        private static JObject Button(string label, string action, bool primary = false)
        {
            var button = new JObject(new JProperty("label", label), new JProperty("action", action));
            if (primary) button.Add("primary", true);
            return button;
        }

        /// <summary>What the chat keeps as the user's side of a draft, so a later "make it shorter" has its context.</summary>
        public static string Asked(ActionType action, bool writesReply, string instructions)
        {
            if (action == ActionType.Draft) return "Write the email: " + (instructions ?? "").Trim();
            if (action == ActionType.Custom) return "Apply the review's changes.";
            if (writesReply) return "Write my reply (" + ReplyLabel(action).ToLowerInvariant() + ").";
            return action == ActionType.Arabic || action == ActionType.English ? "Translate what I wrote to " + action + "." : Label(action) + " what I wrote.";
        }

        private static string Label(ActionType action)
        {
            return action == ActionType.Revise ? "Improve"
                : action == ActionType.Arabic || action == ActionType.English ? "Translate to " + action
                : action.ToString();
        }

        private static string ReplyLabel(ActionType action)
        {
            switch (action)
            {
                case ActionType.Shorten: return "Short";
                case ActionType.Lengthen: return "Detailed";
                case ActionType.Formal: return "Formal";
                case ActionType.Friendly: return "Friendly";
                case ActionType.Arabic: return "In Arabic";
                case ActionType.English: return "In English";
                default: return "Standard";
            }
        }

        private static JObject Group(string title, params JObject[] chips)
        {
            return new JObject(new JProperty("title", title), new JProperty("chips", new JArray(chips)));
        }

        private static JObject Chip(string label, string prefix, ActionType action, bool primary = false)
        {
            var chip = new JObject(
                new JProperty("label", label), new JProperty("mode", "action"), new JProperty("id", prefix + action));
            if (primary) chip.Add("primary", true);
            return chip;
        }
    }
}
