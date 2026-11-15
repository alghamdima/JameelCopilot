using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Tools
{
    /// <summary>
    /// An Outlook rule Jameel Copilot proposes for incoming mail. It can only file mail in
    /// a folder that exists or give it a category: nothing that sends mail on
    /// or removes it.
    /// </summary>
    public sealed class RuleProposal
    {
        public string Name { get; set; }
        public IReadOnlyList<string> From { get; set; }
        public IReadOnlyList<string> SubjectContains { get; set; }
        public string FolderId { get; set; }
        public string FolderName { get; set; }
        public string Category { get; set; }

        /// <summary>A proposal from the tool's result, or null when it is not one.</summary>
        public static RuleProposal FromResult(string resultJson)
        {
            try
            {
                var o = Prefs.StoredJson.Parse(resultJson ?? "");
                return (string)o?["result_type"] == "rule_proposal" ? Read(o) : null;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        /// <summary>The rule, or null without a name, a condition and an action.</summary>
        internal static RuleProposal Read(JObject o)
        {
            var rule = new RuleProposal
            {
                Name = StructuredAnswers.Field(o, "name", 100),
                From = Strings(o?["from"], 200),
                SubjectContains = Strings(o?["subject_contains"], 100),
                FolderId = StructuredAnswers.Field(o, "folder_id", 100),
                FolderName = StructuredAnswers.Field(o, "folder", 200),
                Category = StructuredAnswers.Field(o, "category", 60),
            };
            var hasCondition = rule.From.Count + rule.SubjectContains.Count > 0;
            var hasAction = rule.FolderId.Length > 0 || rule.Category.Length > 0;
            return rule.Name.Length > 0 && hasCondition && hasAction ? rule : null;
        }

        private static List<string> Strings(JToken list, int maxChars)
        {
            return (list as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => StructuredAnswers.OneLine((string)t)).Where(s => s.Length > 0 && s.Length <= maxChars)
                .Distinct(System.StringComparer.OrdinalIgnoreCase).Take(10).ToList();
        }
    }

    /// <summary>Outlook's rules, for the surfaces that have them.</summary>
    public interface IRuleSurface
    {
        /// <summary>The name of a folder a rule may move mail to; null for an unknown id, and for Deleted Items and Junk.</summary>
        string RuleFolderName(string folderId);

        /// <summary>Creates the rule in Outlook. Not a model tool: only a button the user presses calls it.</summary>
        bool CreateRule(RuleProposal rule);
    }

    /// <summary>
    /// Tool: jc_propose_rule. Shows the user a rule to create. It creates
    /// nothing: the user reads the card and presses Create rule, or does not.
    /// </summary>
    public sealed class ProposeRuleTool : IOutlookTool
    {
        public const string ToolName = "jc_propose_rule";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var rule = RuleProposal.Read(JObject.Parse(argsJson ?? "{}"));
            if (rule == null)
            {
                return Task.FromResult(CalendarArgs.Error("a name, a condition (from or subject_contains) and an action (folder_id or category) are required"));
            }
            if (rule.FolderId.Length > 0)
            {
                // The folder's name on the card is Outlook's, never the model's.
                rule.FolderName = (surface as IRuleSurface)?.RuleFolderName(rule.FolderId);
                if (string.IsNullOrEmpty(rule.FolderName))
                {
                    return Task.FromResult(CalendarArgs.Error("folder_id must be the id of a mail folder from outlook_list_folders, and not Deleted Items or Junk"));
                }
            }
            return Task.FromResult(new JObject(
                new JProperty("result_type", "rule_proposal"),
                new JProperty("name", rule.Name),
                new JProperty("from", new JArray(rule.From.Cast<object>())),
                new JProperty("subject_contains", new JArray(rule.SubjectContains.Cast<object>())),
                new JProperty("folder_id", rule.FolderId),
                new JProperty("folder", rule.FolderName ?? ""),
                new JProperty("category", rule.Category),
                new JProperty("status", "Shown to the user, who creates the rule or not. Nothing is created yet: say that you proposed it."))
                .ToString(Formatting.None));
        }
    }
}
