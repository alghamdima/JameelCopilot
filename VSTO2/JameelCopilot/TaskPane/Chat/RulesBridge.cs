using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Diagnostics;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// The rule cards of a pane. Jameel Copilot keeps each rule it proposed; the page's
    /// "Create rule" button names the card, and the rule created is the one
    /// kept here, once, never one the page sends.
    /// </summary>
    internal sealed class RulesBridge
    {
        private readonly IRuleSurface _surface;
        private readonly Func<string, Task> _runScript;
        private readonly Dictionary<string, RuleProposal> _proposed = new Dictionary<string, RuleProposal>(StringComparer.Ordinal);

        public RulesBridge(IRuleSurface surface, Func<string, Task> runScript)
        {
            _surface = surface;
            _runScript = runScript ?? throw new ArgumentNullException(nameof(runScript));
        }

        /// <summary>A tool finished: when it proposed a rule, keep it for its card.</summary>
        public void Offered(string callId, string toolName, bool ok, string resultJson)
        {
            if (!ok || callId == null || toolName != ProposeRuleTool.ToolName) return;
            var rule = RuleProposal.FromResult(resultJson);
            if (rule != null) _proposed[callId] = rule;
        }

        /// <summary>True when the message was a rule card's.</summary>
        public bool Handle(string type, JObject payload)
        {
            if (type != "rule_create") return false;
            var token = payload?["call"];
            var callId = token != null && token.Type == JTokenType.String ? (string)token : null;
            if (callId == null || !_proposed.TryGetValue(callId, out var rule)) return true;
            _proposed.Remove(callId);   // one click, one rule

            var created = false;
            try
            {
                created = _surface != null && _surface.CreateRule(rule);
            }
            catch (Exception ex)
            {
                TraceLog.Write("CreateRule failed: " + ex.GetType().Name, "Jameel Copilot");
            }
            _ = _runScript("jameelcopilot.markRule(" + JsonConvert.SerializeObject(callId) + ", " +
                JsonConvert.SerializeObject(created ? "created - see Rules in Outlook" : "could not be created; make it in Outlook's Rules instead") + ", " +
                (created ? "true" : "false") + ");");
            return true;
        }
    }
}
