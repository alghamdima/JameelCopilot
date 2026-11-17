using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Tools
{
    /// <summary>A name the user gave and the person the address book has for it.</summary>
    public sealed class PersonMatch
    {
        public string Query { get; set; }
        public string Name { get; set; }

        /// <summary>Null when the address book has nobody, or more than one person, under the name.</summary>
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public string Department { get; set; }
    }

    public sealed class OrgPerson
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public string Department { get; set; }
    }

    public sealed class OrgContextInfo
    {
        public OrgPerson User { get; set; }
        public OrgPerson Manager { get; set; }
        public IReadOnlyList<OrgPerson> DirectReports { get; set; }
        public IReadOnlyList<OrgPerson> Peers { get; set; }
        public string PreferredBuilding { get; set; }
    }

    public interface IOrgSurface
    {
        OrgContextInfo GetOrgContext();
    }

    public sealed class RoomInfo
    {
        public string Name { get; set; }
        public string Email { get; set; }

        /// <summary>Free for the whole time asked; null when the room's calendar cannot be read.</summary>
        public bool? Free { get; set; }
    }

    /// <summary>The address book, for the surfaces that have one (the tests' fakes need not).</summary>
    public interface IPeopleSurface
    {
        /// <summary>One entry for each name, in the order given.</summary>
        IReadOnlyList<PersonMatch> ResolvePeople(IReadOnlyList<string> names);

        /// <summary>
        /// The meeting rooms whose name contains the text (all of them for an
        /// empty text), each with whether it is free from start to end. Null
        /// when the address book has no list of rooms.
        /// </summary>
        IReadOnlyList<RoomInfo> FindRooms(DateTime start, DateTime end, string nameContains, int max);
    }

    /// <summary>
    /// Tool: outlook_resolve_people. Turns the names the user typed into the
    /// people the address book knows, the way "Check Names" does. A meeting
    /// set up with a bare first name reaches nobody and no calendar is read.
    /// </summary>
    public sealed class ResolvePeopleTool : IOutlookTool
    {
        public const string ToolName = "outlook_resolve_people";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var book = surface as IPeopleSurface;
            if (book == null) return Task.FromResult(CalendarArgs.Error("the address book is not available here"));
            var names = (JObject.Parse(argsJson ?? "{}")["names"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => StructuredAnswers.OneLine((string)t)).Where(s => s.Length > 0 && s.Length <= 200)
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(10).ToList();
            if (names.Count == 0) return Task.FromResult(CalendarArgs.Error("names is required: one to ten names or email addresses"));

            var matches = book.ResolvePeople(names) ?? new PersonMatch[0];
            return Task.FromResult(new JObject(
                new JProperty("people", new JArray(matches.Select(m =>
                {
                    var found = !string.IsNullOrEmpty(m.Email);
                    var person = new JObject(new JProperty("asked", m.Query ?? ""), new JProperty("found", found));
                    if (found)
                    {
                        person.Add("name", m.Name ?? "");
                        person.Add("email", m.Email);
                        if (!string.IsNullOrEmpty(m.JobTitle)) person.Add("job_title", m.JobTitle);
                        if (!string.IsNullOrEmpty(m.Department)) person.Add("department", m.Department);
                    }
                    return person;
                }))),
                new JProperty("note", "found:false means the address book has nobody, or more than one person, under that name."))
                .ToString(Formatting.None));
        }
    }

    /// <summary>Tool: outlook_find_rooms. The meeting rooms and which are free for a time.</summary>
    public sealed class FindRoomsTool : IOutlookTool
    {
        public const string ToolName = "outlook_find_rooms";
        public const int MaxRooms = 12;
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var book = surface as IPeopleSurface;
            if (book == null) return Task.FromResult(CalendarArgs.Error("the address book is not available here"));
            var args = JObject.Parse(argsJson ?? "{}");
            var start = CalendarArgs.Date(args, "start");
            var end = CalendarArgs.Date(args, "end");
            if (!start.HasValue || !end.HasValue || end <= start || (end.Value - start.Value).TotalHours > 12)
            {
                return Task.FromResult(CalendarArgs.Error("start and end are required: local time, end after start, at most 12 hours apart"));
            }

            var prefBuilding = (surface as IOrgSurface)?.GetOrgContext()?.PreferredBuilding ?? "";
            var rooms = book.FindRooms(start.Value, end.Value, StructuredAnswers.Field(args, "name_contains", 60), MaxRooms);
            var result = new JObject(
                new JProperty("rooms", new JArray((rooms ?? new RoomInfo[0])
                    .OrderBy(r => r.Free == true ? 0 : r.Free == null ? 1 : 2)
                    .ThenBy(r => !string.IsNullOrEmpty(prefBuilding) && (r.Name ?? "").IndexOf(prefBuilding, StringComparison.OrdinalIgnoreCase) >= 0 ? 0 : 1)
                    .Select(r => new JObject(
                        new JProperty("name", r.Name ?? ""),
                        new JProperty("email", r.Email ?? ""),
                        new JProperty("free", r.Free == true ? "yes" : r.Free == false ? "no" : "unknown"))))),
                new JProperty("start", CalendarArgs.Local(start.Value)),
                new JProperty("end", CalendarArgs.Local(end.Value)));
            if (!string.IsNullOrEmpty(prefBuilding)) result.Add("preferred_building", prefBuilding);
            if (rooms == null) result.Add("note", "The address book has no list of meeting rooms. Ask the user for the room's name.");
            else if (rooms.Count >= MaxRooms) result.Add("note", "Only the first " + MaxRooms + " rooms are listed. Use name_contains to narrow them.");
            return Task.FromResult(result.ToString(Formatting.None));
        }
    }

    /// <summary>
    /// Tool: outlook_get_org_context. Returns the current user's Exchange
    /// organizational profile: job title, department, direct manager, direct
    /// reports (their team), peers under the same manager, and preferred building.
    /// </summary>
    public sealed class GetOrgContextTool : IOutlookTool
    {
        public const string ToolName = "outlook_get_org_context";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var org = (surface as IOrgSurface)?.GetOrgContext();
            if (org == null || org.User == null)
            {
                return Task.FromResult(CalendarArgs.Error("organizational profile is not available here"));
            }

            JObject ToJson(OrgPerson p) => p == null ? null : new JObject(
                new JProperty("name", p.Name ?? ""),
                new JProperty("email", p.Email ?? ""),
                new JProperty("job_title", p.JobTitle ?? ""),
                new JProperty("department", p.Department ?? ""));

            var obj = new JObject(
                new JProperty("user", ToJson(org.User)),
                new JProperty("manager", org.Manager != null ? (JToken)ToJson(org.Manager) : JValue.CreateNull()),
                new JProperty("direct_reports", new JArray((org.DirectReports ?? new OrgPerson[0]).Select(ToJson))),
                new JProperty("peers", new JArray((org.Peers ?? new OrgPerson[0]).Select(ToJson))),
                new JProperty("preferred_building", org.PreferredBuilding ?? ""));
            return Task.FromResult(obj.ToString(Formatting.None));
        }
    }
}
