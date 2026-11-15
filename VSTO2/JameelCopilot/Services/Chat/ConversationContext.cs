using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Chat
{
    /// <summary>
    /// Per-turn context for <c>CustomApiChatService.RunTurnAsync</c>. Carries the
    /// accumulated conversation history, the system instructions to seed the
    /// turn, and which tools the chat service should expose for this turn.
    /// </summary>
    public sealed class ConversationContext
    {
        public string SystemInstructions { get; set; } = "";
        public List<JObject> History { get; set; } = new List<JObject>();

        /// <summary>What the pane read from the mailbox for this turn (draft, selection). Sent as data.</summary>
        public string Context { get; set; } = "";
        public bool IncludeWriteTools { get; set; } = true;

        /// <summary>What the user asked Jameel Copilot to remember, as the prompt carries it; empty for none.</summary>
        public string Memory { get; set; } = "";

        /// <summary>Sent as <c>temperature</c> when set; null sends none.</summary>
        public double? Temperature { get; set; }

        /// <summary>False for a one-shot rewrite: no tool is offered and none is run.</summary>
        public bool UseTools { get; set; } = true;

        /// <summary>The answer is JSON for Jameel Copilot's screen (the structured core); set with UseTools false.</summary>
        public bool Structured { get; set; }
    }
}
