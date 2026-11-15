using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services
{
    public sealed class CoachPoint
    {
        public string Id { get; set; }
        public string Issue { get; set; }
        public string Suggestion { get; set; }
    }

    /// <summary>
    /// "Review" (Copilot's coaching): points about the user's draft before it
    /// is sent. The user picks the ones to apply; the rewrite then makes those
    /// changes and nothing else.
    /// </summary>
    internal static class DraftCoaching
    {
        public const int MaxPoints = 6;

        public static string Instructions(string language) =>
            "Review the user's draft under Context as a careful colleague would before it is sent. Reply language: " + language + ".\n" +
            "Check: does it answer everything the message being replied to asks; is the tone right for the people it goes to; " +
            "is it clear and complete (dates, numbers, next steps); spelling and grammar; does it mention an attachment that is not attached.\n" +
            "Give at most " + MaxPoints + " points, the most important first. Each says what is wrong and exactly what to change. " +
            "No points when it is ready to send.\n" +
            "Format: {\"summary\":\"<one sentence on the draft as a whole>\",\"points\":[{\"id\":\"1\",\"issue\":\"<what is wrong>\",\"suggestion\":\"<what to change>\"}]}";

        public static string Summary(JObject answer) => StructuredAnswers.Field(answer, "summary", 300);

        /// <summary>The points, one line each, at most six; a point without a suggestion is dropped.</summary>
        public static List<CoachPoint> Parse(JObject answer)
        {
            var points = new List<CoachPoint>();
            foreach (var entry in (answer?["points"] as JArray ?? new JArray()).OfType<JObject>())
            {
                var suggestion = StructuredAnswers.Field(entry, "suggestion", 300);
                if (suggestion.Length == 0 || points.Count >= MaxPoints) continue;
                points.Add(new CoachPoint
                {
                    Id = (points.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),   // ours, not the model's
                    Issue = StructuredAnswers.Field(entry, "issue", 300),
                    Suggestion = suggestion,
                });
            }
            return points;
        }

        public static JArray ToJson(IEnumerable<CoachPoint> points) =>
            new JArray(points.Select(p => new JObject(
                new JProperty("id", p.Id), new JProperty("issue", p.Issue), new JProperty("suggestion", p.Suggestion))));

        /// <summary>The instructions of the rewrite: the chosen changes, and nothing else.</summary>
        public static string ApplyInstructions(IEnumerable<CoachPoint> chosen)
        {
            var sb = new StringBuilder("Make these changes to the text and nothing else. Keep everything else as it is, in the same language:");
            foreach (var p in chosen) sb.Append("\n- ").Append(p.Suggestion);
            return sb.ToString();
        }

        /// <summary>What the chat keeps, so a follow-up knows what the review said.</summary>
        public static string AsText(string summary, IReadOnlyList<CoachPoint> points)
        {
            var sb = new StringBuilder(summary);
            foreach (var p in points) sb.Append("\n- ").Append(p.Issue.Length > 0 ? p.Issue + " -> " : "").Append(p.Suggestion);
            return sb.ToString().Trim();
        }
    }
}
