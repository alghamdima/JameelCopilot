using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Tools
{
    public sealed class CalendarEvent
    {
        public string Subject { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; }
        public string Organizer { get; set; }
        public string Attendees { get; set; }
        public bool AllDay { get; set; }
        /// <summary>"free", "tentative", "busy", "out_of_office" or "elsewhere".</summary>
        public string ShowAs { get; set; }
    }

    /// <summary>The calendar, for the surfaces that have one (the tests' fakes need not).</summary>
    public interface ICalendarSurface
    {
        IReadOnlyList<CalendarEvent> ListCalendarEvents(DateTime from, DateTime to, int max);

        /// <summary>
        /// Free/busy of each person from <paramref name="day"/> at midnight, one
        /// character per 30 minutes ('0' free); null for someone whose calendar
        /// cannot be read. The user is under the key "me".
        /// </summary>
        IDictionary<string, string> GetFreeBusy(IEnumerable<string> people, DateTime day);
    }

    /// <summary>Common free times in free/busy strings. Pure, so it is tested alone.</summary>
    internal static class MeetingTimes
    {
        public const int SlotMinutes = 30;

        public static List<Tuple<DateTime, DateTime>> Find(
            IEnumerable<string> freeBusy, DateTime gridStart, DateTime from, DateTime to, int durationMinutes,
            TimeSpan dayStart, TimeSpan dayEnd, ISet<DayOfWeek> workDays, int max = 8, int perDay = 3)
        {
            var grids = freeBusy.Where(f => f != null).ToList();
            var need = Math.Max(1, (durationMinutes + SlotMinutes - 1) / SlotMinutes);
            var found = new List<Tuple<DateTime, DateTime>>();
            for (var day = from.Date; day <= to.Date && found.Count < max; day = day.AddDays(1))
            {
                if (!workDays.Contains(day.DayOfWeek)) continue;
                var today = 0;
                for (var start = day + dayStart; start.AddMinutes(need * SlotMinutes) <= day + dayEnd && found.Count < max && today < perDay;
                     start = start.AddMinutes(SlotMinutes))
                {
                    if (start < from) continue;
                    var first = (int)((start - gridStart).TotalMinutes / SlotMinutes);
                    if (first < 0) continue;
                    var free = grids.All(g => first + need <= g.Length && g.Substring(first, need).All(c => c == '0'));
                    if (!free) continue;
                    found.Add(Tuple.Create(start, start.AddMinutes(durationMinutes)));
                    today++;
                    start = start.AddMinutes(need * SlotMinutes - SlotMinutes);   // the next proposal starts after this one ends
                }
            }
            return found;
        }

        /// <summary>Whether the whole time is free in a free/busy string; null when the string does not cover it.</summary>
        public static bool? IsFree(string freeBusy, DateTime gridStart, DateTime start, DateTime end)
        {
            if (string.IsNullOrEmpty(freeBusy) || end <= start) return null;
            var first = (int)Math.Floor((start - gridStart).TotalMinutes / SlotMinutes);
            var last = (int)Math.Ceiling((end - gridStart).TotalMinutes / SlotMinutes);   // the first slot after the time
            if (first < 0 || last > freeBusy.Length) return null;
            return freeBusy.Substring(first, last - first).All(c => c == '0');
        }

        /// <summary>"sun-thu" (the default) or "mon-fri".</summary>
        public static ISet<DayOfWeek> WorkDays(string spec)
        {
            return string.Equals((spec ?? "").Trim(), "mon-fri", StringComparison.OrdinalIgnoreCase)
                ? new HashSet<DayOfWeek> { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday }
                : new HashSet<DayOfWeek> { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday };
        }
    }

    internal static class CalendarArgs
    {
        public static DateTime? Date(JObject args, string name)
        {
            var token = args?[name];
            if (token == null) return null;
            if (token.Type == JTokenType.Date) return ((DateTime)token);
            return token.Type == JTokenType.String && DateTime.TryParse((string)token, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out var date) ? date : (DateTime?)null;
        }

        public static TimeSpan Time(JObject args, string name, TimeSpan fallback)
        {
            var token = args?[name];
            return token != null && token.Type == JTokenType.String && TimeSpan.TryParseExact((string)token, @"hh\:mm", CultureInfo.InvariantCulture, out var time)
                ? time : fallback;
        }

        public static string Error(string message) =>
            new JObject(new JProperty("error", new JObject(new JProperty("code", "invalid_arguments"), new JProperty("message", message))))
                .ToString(Formatting.None);

        public static string Local(DateTime time) => time.ToString("yyyy-MM-dd'T'HH:mm", CultureInfo.InvariantCulture);
    }

    /// <summary>Tool: outlook_list_calendar_events. The user's meetings and appointments in a period.</summary>
    public sealed class ListCalendarEventsTool : IOutlookTool
    {
        public const string ToolName = "outlook_list_calendar_events";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var calendar = surface as ICalendarSurface;
            if (calendar == null) return Task.FromResult(CalendarArgs.Error("the calendar is not available here"));
            var args = JObject.Parse(argsJson ?? "{}");
            var from = CalendarArgs.Date(args, "date_from") ?? DateTime.Today;
            var to = CalendarArgs.Date(args, "date_to") ?? from.Date.AddDays(1);
            if (to <= from) return Task.FromResult(CalendarArgs.Error("date_to must be after date_from"));
            if ((to - from).TotalDays > 31) to = from.AddDays(31);
            var max = Math.Max(1, Math.Min(50, args["max_results"]?.Type == JTokenType.Integer ? (int)args["max_results"] : 30));

            var events = calendar.ListCalendarEvents(from, to, max) ?? new CalendarEvent[0];
            return Task.FromResult(new JObject(
                new JProperty("events", new JArray(events.Select(e => new JObject(
                    new JProperty("subject", e.Subject ?? ""),
                    new JProperty("start", CalendarArgs.Local(e.Start)),
                    new JProperty("end", CalendarArgs.Local(e.End)),
                    new JProperty("all_day", e.AllDay),
                    new JProperty("location", e.Location ?? ""),
                    new JProperty("organizer", e.Organizer ?? ""),
                    new JProperty("attendees", e.Attendees ?? ""),
                    new JProperty("show_as", e.ShowAs ?? ""))))),
                new JProperty("time_zone", "local")).ToString(Formatting.None));
        }
    }

    /// <summary>Tool: outlook_find_meeting_times. Times the user and the people named are all free.</summary>
    public sealed class FindMeetingTimesTool : IOutlookTool
    {
        public const string ToolName = "outlook_find_meeting_times";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var calendar = surface as ICalendarSurface;
            if (calendar == null) return Task.FromResult(CalendarArgs.Error("the calendar is not available here"));
            var args = JObject.Parse(argsJson ?? "{}");
            var people = (args["attendees"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => ((string)t).Trim()).Where(s => s.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).Take(20).ToList();
            var duration = args["duration_minutes"]?.Type == JTokenType.Integer ? (int)args["duration_minutes"] : 30;
            if (duration < 15 || duration > 480) return Task.FromResult(CalendarArgs.Error("duration_minutes must be 15 to 480"));
            var now = DateTime.Now;
            var from = CalendarArgs.Date(args, "date_from") ?? now;
            if (from < now) from = now;
            var to = CalendarArgs.Date(args, "date_to") ?? from.Date.AddDays(7);
            if ((to - from).TotalDays > 21) to = from.AddDays(21);
            if (to <= from) return Task.FromResult(CalendarArgs.Error("date_to must be after date_from"));

            var freeBusy = calendar.GetFreeBusy(people, from.Date) ?? new Dictionary<string, string>();
            var unknown = people.Where(p => !freeBusy.TryGetValue(p, out var fb) || fb == null).ToList();
            var slots = MeetingTimes.Find(freeBusy.Values, from.Date, from, to, duration,
                CalendarArgs.Time(args, "day_start", new TimeSpan(8, 0, 0)), CalendarArgs.Time(args, "day_end", new TimeSpan(17, 0, 0)),
                MeetingTimes.WorkDays((string)args["working_days"]));

            return Task.FromResult(new JObject(
                new JProperty("slots", new JArray(slots.Select(s => new JObject(
                    new JProperty("start", CalendarArgs.Local(s.Item1)), new JProperty("end", CalendarArgs.Local(s.Item2)))))),
                new JProperty("calendars_not_read", new JArray(unknown.Cast<object>())),
                new JProperty("time_zone", "local")).ToString(Formatting.None));
        }
    }

    public sealed class MeetingProposal
    {
        public string Subject { get; set; }
        public IReadOnlyList<string> Attendees { get; set; }
        public IReadOnlyList<Tuple<DateTime, DateTime>> Slots { get; set; }
        public string Agenda { get; set; }

        /// <summary>The meeting room's name or address, from outlook_find_rooms; "" for none.</summary>
        public string Room { get; set; }

        /// <summary>A proposal from the tool's result, or null when it is not one.</summary>
        public static MeetingProposal FromResult(string resultJson)
        {
            try
            {
                var o = Prefs.StoredJson.Parse(resultJson ?? "");
                if ((string)o?["result_type"] != "meeting_proposal") return null;
                return Read(o);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        internal static MeetingProposal Read(JObject o)
        {
            var slots = new List<Tuple<DateTime, DateTime>>();
            foreach (var s in (o["slots"] as JArray ?? new JArray()).OfType<JObject>())
            {
                var start = CalendarArgs.Date(s, "start");
                var end = CalendarArgs.Date(s, "end");
                if (start.HasValue && end.HasValue && end > start && (end.Value - start.Value).TotalHours <= 8 && slots.Count < 5)
                {
                    slots.Add(Tuple.Create(start.Value, end.Value));
                }
            }
            var attendees = (o["attendees"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => StructuredAnswers.OneLine((string)t)).Where(s => s.Length > 0 && s.Length <= 200).Distinct().Take(30).ToList();
            var subject = StructuredAnswers.Field(o, "subject", 200);
            if (slots.Count == 0 || subject.Length == 0) return null;
            var agendaToken = o["agenda"];
            var agenda = agendaToken != null && agendaToken.Type == JTokenType.String ? (string)agendaToken : "";
            return new MeetingProposal
            {
                Subject = subject, Attendees = attendees, Slots = slots, Agenda = agenda.Length > 2000 ? agenda.Substring(0, 2000) : agenda,
                Room = StructuredAnswers.Field(o, "room", 200),
            };
        }
    }

    /// <summary>
    /// Tool: jc_propose_meeting. Shows the user a meeting to set up. It sends
    /// nothing: the user picks a time and Outlook's own meeting window opens.
    /// </summary>
    public sealed class ProposeMeetingTool : IOutlookTool
    {
        public const string ToolName = "jc_propose_meeting";
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var proposal = MeetingProposal.Read(JObject.Parse(argsJson ?? "{}"));
            if (proposal == null) return Task.FromResult(CalendarArgs.Error("subject and one to five slots (start, end) are required"));
            return Task.FromResult(new JObject(
                new JProperty("result_type", "meeting_proposal"),
                new JProperty("subject", proposal.Subject),
                new JProperty("attendees", new JArray(proposal.Attendees.Cast<object>())),
                new JProperty("slots", new JArray(proposal.Slots.Select(s => new JObject(
                    new JProperty("start", CalendarArgs.Local(s.Item1)), new JProperty("end", CalendarArgs.Local(s.Item2)))))),
                new JProperty("agenda", proposal.Agenda),
                new JProperty("room", proposal.Room),
                new JProperty("status", "Shown to the user, who picks a time to open the meeting request. Nothing is sent: say that you proposed it."))
                .ToString(Formatting.None));
        }
    }
}
