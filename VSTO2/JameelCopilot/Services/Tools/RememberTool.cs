using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Prefs;

namespace JameelCopilot.Services.Tools
{
    /// <summary>
    /// Tool: jc_remember. Proposes a note for the user's memory. It saves
    /// nothing: the page shows the note and only the user's Save writes it.
    /// A note goes into every later request, so a note the model could write
    /// by itself would let one hostile email steer Jameel Copilot from then on.
    /// </summary>
    public sealed class RememberTool : IOutlookTool
    {
        public const string ToolName = "jc_remember";

        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var token = JObject.Parse(argsJson ?? "{}")["note"];
            var note = UserMemoryService.Clean(token != null && token.Type == JTokenType.String ? (string)token : null);
            if (note.Length == 0)
            {
                return Task.FromResult(new JObject(new JProperty("error", new JObject(
                    new JProperty("code", "invalid_arguments"),
                    new JProperty("message", "note is required: one short sentence")))).ToString(Formatting.None));
            }

            return Task.FromResult(new JObject(
                new JProperty("result_type", "memory_proposal"),
                new JProperty("note", note),
                new JProperty("status", "The note is shown to the user, who decides whether to save it. It is not saved yet: say that you proposed it."))
                .ToString(Formatting.None));
        }
    }
}
