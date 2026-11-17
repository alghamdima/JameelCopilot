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
            TimeSpan dayStart, TimeSpan dayEnd, ISet<DayOfWeek> workDays, int max = 8, int perDay = 3,
            Func<DateTime, DateTime, bool> skipSlot = null)
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
                    var end = start.AddMinutes(durationMinutes);
                    if (skipSlot != null && skipSlot(start, end)) continue;
                    var first = (int)((start - gridStart).TotalMinutes / SlotMinutes);
                    if (first < 0) continue;
                    var free = grids.All(g => first + need <= g.Length && g.Substring(first, need).All(c => c == '0'));
                    if (!free) continue;
                    found.Add(Tuple.Create(start, end));
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
        public static readonly TimeSpan DefaultBreakStart = new TimeSpan(12, 30, 0);
        public static readonly TimeSpan DefaultBreakEnd = new TimeSpan(13, 30, 0);
        public string Name => ToolName;

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var calendar = surface as ICalendarSurface;
            if (calendar == null) return Task.FromResult(CalendarArgs.Error("the calendar is not available here"));
            var args = JObject.Parse(argsJson ?? "{}");
            var people = (args["attendees"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => ((string)t).Trim()).Where(s => s.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).Take(20).ToList();
            var optionalPeople = (args["optional_attendees"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => ((string)t).Trim()).Where(s => s.Length > 0 && !people.Contains(s, StringComparer.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(20).ToList();
            var checkOptional = args["check_optional_availability"]?.Type == JTokenType.Boolean && (bool)args["check_optional_availability"];
            var duration = args["duration_minutes"]?.Type == JTokenType.Integer ? (int)args["duration_minutes"] : 60;
            if (duration < 15 || duration > 480) return Task.FromResult(CalendarArgs.Error("duration_minutes must be 15 to 480"));
            var now = DateTime.Now;
            var from = CalendarArgs.Date(args, "date_from") ?? now;
            if (from < now) from = now;
            var to = CalendarArgs.Date(args, "date_to") ?? from.Date.AddDays(7);
            if ((to - from).TotalDays > 21) to = from.AddDays(21);
            if (to <= from) return Task.FromResult(CalendarArgs.Error("date_to must be after date_from"));

            var queriedPeople = checkOptional ? people.Concat(optionalPeople).ToList() : people;
            var freeBusy = calendar.GetFreeBusy(queriedPeople, from.Date) ?? new Dictionary<string, string>();
            var unknown = queriedPeople.Where(p => !freeBusy.TryGetValue(p, out var fb) || fb == null).ToList();
            var dayStart = CalendarArgs.Time(args, "day_start", new TimeSpan(9, 0, 0));
            var dayEnd = CalendarArgs.Time(args, "day_end", new TimeSpan(17, 0, 0));
            var workDays = MeetingTimes.WorkDays((string)args["working_days"]);

            // Check required + me first; if checkOptional is true and everyone fits, prefer slots where optional attendees are also free.
            var requiredGrids = new List<string>();
            if (freeBusy.TryGetValue("me", out var myGrid) && myGrid != null) requiredGrids.Add(myGrid);
            foreach (var p in people)
            {
                if (freeBusy.TryGetValue(p, out var g) && g != null) requiredGrids.Add(g);
            }
            // Fallback for unit-test fakes that key FreeBusy differently.
            if (requiredGrids.Count == 0) requiredGrids = freeBusy.Values.Where(v => v != null).ToList();

            var allGrids = checkOptional ? freeBusy.Values.Where(v => v != null).ToList() : requiredGrids;
            var rawAll = MeetingTimes.Find(allGrids, from.Date, from, to, duration, dayStart, dayEnd, workDays, max: 24, perDay: 8, skipSlot: OverlapsBreak);
            if (rawAll.Count == 0 && checkOptional && optionalPeople.Count > 0)
            {
                rawAll = MeetingTimes.Find(requiredGrids, from.Date, from, to, duration, dayStart, dayEnd, workDays, max: 24, perDay: 8, skipSlot: OverlapsBreak);
            }
            if (rawAll.Count == 0)
            {
                rawAll = MeetingTimes.Find(requiredGrids, from.Date, from, to, duration,
                    new TimeSpan(8, 0, 0), new TimeSpan(17, 0, 0), workDays, max: 18, perDay: 6, skipSlot: OverlapsBreak);
            }

            var slots = rawAll
                .GroupBy(s => s.Item1.Date)
                .SelectMany(g => g.OrderBy(s => HasAdjacentBusy(allGrids, from.Date, s.Item1, s.Item2) ? 1 : 0).ThenBy(s => s.Item1).Take(3))
                .Take(9)
                .ToList();

            // Partial-availability fallback when 100% overlap has 0 slots.
            var conflictsBySlot = new Dictionary<int, List<string>>();
            if (slots.Count == 0 && people.Count > 1 && myGrid != null)
            {
                var mySlots = MeetingTimes.Find(new[] { myGrid }, from.Date, from, to, duration, dayStart, dayEnd, workDays, max: 24, perDay: 8, skipSlot: OverlapsBreak);
                var bestPartial = mySlots
                    .Select(s => new
                    {
                        Slot = s,
                        Busy = people.Where(p => freeBusy.TryGetValue(p, out var pg) && pg != null && MeetingTimes.IsFree(pg, from.Date, s.Item1, s.Item2) == false).ToList()
                    })
                    .Where(x => x.Busy.Count < people.Count)
                    .OrderBy(x => x.Busy.Count)
                    .ThenBy(x => x.Slot.Item1)
                    .Take(5)
                    .ToList();
                for (int idx = 0; idx < bestPartial.Count; idx++)
                {
                    slots.Add(bestPartial[idx].Slot);
                    conflictsBySlot[idx] = bestPartial[idx].Busy;
                }
            }

            var slotArray = new JArray();
            var book = surface as IPeopleSurface;
            var prefBuilding = (surface as IOrgSurface)?.GetOrgContext()?.PreferredBuilding ?? "";
            for (int i = 0; i < slots.Count; i++)
            {
                var s = slots[i];
                var slotObj = new JObject(
                    new JProperty("start", CalendarArgs.Local(s.Item1)),
                    new JProperty("end", CalendarArgs.Local(s.Item2)));
                if (conflictsBySlot.TryGetValue(i, out var busyList) && busyList.Count > 0)
                {
                    slotObj.Add("busy_attendees", new JArray(busyList.Cast<object>()));
                }
                if (checkOptional && optionalPeople.Count > 0)
                {
                    var optFree = optionalPeople.Where(p => freeBusy.TryGetValue(p, out var og) && og != null && MeetingTimes.IsFree(og, from.Date, s.Item1, s.Item2) == true).ToList();
                    var optBusy = optionalPeople.Where(p => freeBusy.TryGetValue(p, out var og) && og != null && MeetingTimes.IsFree(og, from.Date, s.Item1, s.Item2) == false).ToList();
                    slotObj.Add("optional_free", new JArray(optFree.Cast<object>()));
                    if (optBusy.Count > 0) slotObj.Add("optional_busy", new JArray(optBusy.Cast<object>()));
                }
                if (book != null && i < 3)
                {
                    var rooms = book.FindRooms(s.Item1, s.Item2, "", FindRoomsTool.MaxRooms);
                    if (rooms != null)
                    {
                        var freeRooms = rooms
                            .Where(r => r.Free == true)
                            .OrderBy(r => !string.IsNullOrEmpty(prefBuilding) && (r.Name ?? "").IndexOf(prefBuilding, StringComparison.OrdinalIgnoreCase) >= 0 ? 0 : 1)
                            .Take(4)
                            .Select(r => new JObject(new JProperty("name", r.Name ?? ""), new JProperty("email", r.Email ?? "")))
                            .ToList();
                        slotObj.Add("free_rooms", new JArray(freeRooms));
                    }
                }
                slotArray.Add(slotObj);
            }

            var outObj = new JObject(
                new JProperty("slots", slotArray),
                new JProperty("duration_minutes", duration),
                new JProperty("excluded_break", "12:30-13:30"),
                new JProperty("calendars_not_read", new JArray(unknown.Cast<object>())),
                new JProperty("time_zone", "local"));
            if (!string.IsNullOrEmpty(prefBuilding)) outObj.Add("preferred_building", prefBuilding);
            return Task.FromResult(outObj.ToString(Formatting.None));
        }

        public static bool OverlapsBreak(DateTime start, DateTime end)
        {
            if (end.Date != start.Date) return true;
            return start.TimeOfDay < DefaultBreakEnd && end.TimeOfDay > DefaultBreakStart;
        }

        private static bool HasAdjacentBusy(IReadOnlyList<string> grids, DateTime gridStart, DateTime start, DateTime end)
        {
            if (grids == null || grids.Count == 0) return false;
            var beforeIdx = (int)Math.Floor((start - gridStart).TotalMinutes / MeetingTimes.SlotMinutes) - 1;
            var afterIdx = (int)Math.Ceiling((end - gridStart).TotalMinutes / MeetingTimes.SlotMinutes);
            foreach (var g in grids)
            {
                if (string.IsNullOrEmpty(g)) continue;
                if (beforeIdx >= 0 && beforeIdx < g.Length && g[beforeIdx] != '0') return true;
                if (afterIdx >= 0 && afterIdx < g.Length && g[afterIdx] != '0') return true;
            }
            return false;
        }
    }

    public sealed class MeetingProposal
    {
        public string Subject { get; set; }
        public IReadOnlyList<string> Attendees { get; set; }
        public IReadOnlyList<string> OptionalAttendees { get; set; }
        public IReadOnlyList<Tuple<DateTime, DateTime>> Slots { get; set; }
        public string Agenda { get; set; }

        /// <summary>The meeting room's name or address, from outlook_find_rooms; "" for none.</summary>
        public string Room { get; set; }
        public IReadOnlyList<RoomOption> RoomOptions { get; set; }

        public sealed class RoomOption
        {
            public string Label { get; set; }
            public string Value { get; set; }
        }

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
                .Select(t => StructuredAnswers.OneLine((string)t)).Where(s => s.Length > 0 && s.Length <= 200)
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(30).ToList();
            var optionalAttendees = (o["optional_attendees"] as JArray ?? new JArray()).Where(t => t.Type == JTokenType.String)
                .Select(t => StructuredAnswers.OneLine((string)t))
                .Where(s => s.Length > 0 && s.Length <= 200 && !attendees.Contains(s, StringComparer.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(30).ToList();
            var subject = StructuredAnswers.Field(o, "subject", 200);
            if (slots.Count == 0 || subject.Length == 0) return null;
            var agendaToken = o["agenda"];
            var agenda = agendaToken != null && agendaToken.Type == JTokenType.String ? (string)agendaToken : "";
            var options = (o["room_options"] as JArray ?? new JArray()).OfType<JObject>()
                .Select(r => new RoomOption { Label = StructuredAnswers.Field(r, "label", 160), Value = StructuredAnswers.Field(r, "value", 200) })
                .Where(r => r.Label.Length > 0 && r.Value.Length > 0)
                .Take(15)
                .ToList();
            return new MeetingProposal
            {
                Subject = subject, Attendees = attendees, OptionalAttendees = optionalAttendees,
                Slots = slots, Agenda = agenda.Length > 2000 ? agenda.Substring(0, 2000) : agenda,
                Room = StructuredAnswers.Field(o, "room", 200),
                RoomOptions = options,
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

            var roomDisplay = proposal.Room;
            var roomOptions = new JArray();
            var seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var book = surface as IPeopleSurface;
            var org = (surface as IOrgSurface)?.GetOrgContext();
            var prefBuilding = org?.PreferredBuilding ?? "";
            var selfEmail = org?.User?.Email ?? "";
            var reqAttendees = proposal.Attendees.Where(a => string.IsNullOrEmpty(selfEmail) || !string.Equals(a, selfEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            var optAttendees = (proposal.OptionalAttendees ?? new string[0]).Where(a => string.IsNullOrEmpty(selfEmail) || !string.Equals(a, selfEmail, StringComparison.OrdinalIgnoreCase)).ToList();

            if (book != null && proposal.Slots.Count > 0)
            {
                var perSlotSets = new List<IReadOnlyList<RoomInfo>>();
                foreach (var s in proposal.Slots)
                {
                    var rs = book.FindRooms(s.Item1, s.Item2, "", FindRoomsTool.MaxRooms);
                    if (rs != null) perSlotSets.Add(rs);
                }
                if (perSlotSets.Count > 0)
                {
                    var firstRooms = perSlotSets[0];
                    var matched = firstRooms.FirstOrDefault(r =>
                        string.Equals(r.Email, proposal.Room, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(r.Name, proposal.Room, StringComparison.OrdinalIgnoreCase));
                    if (matched != null && !string.IsNullOrEmpty(matched.Name)) roomDisplay = matched.Name;

                    foreach (var r in firstRooms.OrderBy(x => !string.IsNullOrEmpty(prefBuilding) && (x.Name ?? "").IndexOf(prefBuilding, StringComparison.OrdinalIgnoreCase) >= 0 ? 0 : 1))
                    {
                        var val = !string.IsNullOrEmpty(r.Email) ? r.Email : r.Name;
                        if (string.IsNullOrEmpty(val) || string.IsNullOrEmpty(r.Name)) continue;
                        var freeAll = perSlotSets.All(list => list.Any(x => string.Equals(!string.IsNullOrEmpty(x.Email) ? x.Email : x.Name, val, StringComparison.OrdinalIgnoreCase) && x.Free == true));
                        var freeAny = perSlotSets.Any(list => list.Any(x => string.Equals(!string.IsNullOrEmpty(x.Email) ? x.Email : x.Name, val, StringComparison.OrdinalIgnoreCase) && x.Free == true));
                        if (!freeAny) continue;
                        if (!seenValues.Add(val)) continue;
                        var label = freeAll || perSlotSets.Count == 1 ? r.Name : (r.Name + " (selected slots)");
                        roomOptions.Add(new JObject(new JProperty("label", label), new JProperty("value", val)));
                    }
                }
            }
            if (book != null)
            {
                if (seenValues.Add("Microsoft Teams Meeting"))
                {
                    roomOptions.Add(new JObject(new JProperty("label", "Microsoft Teams Meeting"), new JProperty("value", "Microsoft Teams Meeting")));
                }
            }

            var result = new JObject(
                new JProperty("result_type", "meeting_proposal"),
                new JProperty("subject", proposal.Subject),
                new JProperty("attendees", new JArray(reqAttendees.Cast<object>())),
                new JProperty("slots", new JArray(proposal.Slots.Select(s => new JObject(
                    new JProperty("start", CalendarArgs.Local(s.Item1)), new JProperty("end", CalendarArgs.Local(s.Item2)))))),
                new JProperty("agenda", proposal.Agenda),
                new JProperty("room", proposal.Room),
                new JProperty("status", "Shown to the user, who picks a time to open the meeting request. Nothing is sent: say that you proposed it."));
            if (optAttendees.Count > 0)
            {
                result.Add("optional_attendees", new JArray(optAttendees.Cast<object>()));
            }
            if (!string.IsNullOrEmpty(roomDisplay) && roomDisplay != proposal.Room)
            {
                result.Add("room_display", roomDisplay);
            }
            if (roomOptions.Count > 0)
            {
                result.Add("room_options", roomOptions);
            }
            return Task.FromResult(result.ToString(Formatting.None));
        }
    }
}
