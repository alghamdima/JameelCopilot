using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Export;
using JameelCopilot.Services.Tools;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services.Tools
{
    /// <summary>
    /// The calendar (Copilot's scheduling and meeting prep): reading it,
    /// finding common free times, and proposing a meeting the user sends.
    /// </summary>
    public class CalendarToolsTests
    {
        // Sunday 15 Nov 2026 is a working day in sun-thu; Friday 20 and Saturday 21 are not.
        private static readonly DateTime Sunday = new DateTime(2026, 11, 15);

        private static string Grid(params (int fromSlot, int toSlot)[] busy)
        {
            var chars = Enumerable.Repeat('0', 48 * 14).ToArray();
            foreach (var (a, b) in busy) for (var i = a; i < b; i++) chars[i] = '2';
            return new string(chars);
        }

        private static int Slot(int day, int hour, int minute = 0) => day * 48 + hour * 2 + minute / 30;

        [Fact]
        public void Find_GivesTimesEveryoneIsFree_InWorkingHours()
        {
            var me = Grid((Slot(0, 8), Slot(0, 10)));           // busy 08:00-10:00 Sunday
            var sara = Grid((Slot(0, 10), Slot(0, 11)));        // busy 10:00-11:00 Sunday

            var slots = MeetingTimes.Find(new[] { me, sara }, Sunday, Sunday, Sunday.AddDays(1).AddHours(-1), 60,
                new TimeSpan(8, 0, 0), new TimeSpan(17, 0, 0), MeetingTimes.WorkDays("sun-thu"));

            Assert.Equal(new DateTime(2026, 11, 15, 11, 0, 0), slots[0].Item1);
            Assert.Equal(new DateTime(2026, 11, 15, 12, 0, 0), slots[0].Item2);
            Assert.Equal(new DateTime(2026, 11, 15, 12, 0, 0), slots[1].Item1);   // the next one starts when this ends
            Assert.Equal(3, slots.Count);                                        // at most three a day
        }

        [Fact]
        public void Find_SkipsTheWeekend_ThePast_AndUnreadableCalendars()
        {
            var thursdayNoon = new DateTime(2026, 11, 19, 16, 0, 0);   // only 16:00-17:00 left that day
            var slots = MeetingTimes.Find(new[] { Grid(), null }, Sunday, thursdayNoon, new DateTime(2026, 11, 22, 23, 0, 0), 30,
                new TimeSpan(8, 0, 0), new TimeSpan(17, 0, 0), MeetingTimes.WorkDays(null), max: 8);

            Assert.Equal(new DateTime(2026, 11, 19, 16, 0, 0), slots[0].Item1);
            Assert.DoesNotContain(slots, s => s.Item1.DayOfWeek == DayOfWeek.Friday || s.Item1.DayOfWeek == DayOfWeek.Saturday);
            Assert.Contains(slots, s => s.Item1.Date == new DateTime(2026, 11, 22));   // the next Sunday
        }

        [Fact]
        public void WorkDays_SunToThu_ByDefault_OrMonToFri()
        {
            Assert.Contains(DayOfWeek.Sunday, MeetingTimes.WorkDays("sun-thu"));
            Assert.DoesNotContain(DayOfWeek.Friday, MeetingTimes.WorkDays(""));
            Assert.Contains(DayOfWeek.Friday, MeetingTimes.WorkDays("MON-FRI"));
            Assert.DoesNotContain(DayOfWeek.Sunday, MeetingTimes.WorkDays("mon-fri"));
        }

        private sealed class Calendar : ICalendarSurface, IOutlookSurface
        {
            public IReadOnlyList<CalendarEvent> Events = new CalendarEvent[0];
            public IDictionary<string, string> FreeBusy = new Dictionary<string, string>();
            public DateTime? AskedFrom, AskedTo;
            public IReadOnlyList<CalendarEvent> ListCalendarEvents(DateTime from, DateTime to, int max) { AskedFrom = from; AskedTo = to; return Events; }
            public IDictionary<string, string> GetFreeBusy(IEnumerable<string> people, DateTime day) => FreeBusy;

            public ComposeStateResult GetCurrentComposeState(bool includeFullBody) => null;
            public IReadOnlyList<FolderResult> ListFolders() => null;
            public SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public MessageDetail ReadMessage(string messageId, bool includeFullBody) => null;
            public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken)) => null;
            public int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken)) => 0;
            public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) => null;
            public FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public CreatedDraft CreateDraft(CreateDraftArgs args) => null;
            public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) => null;
            public void MarkAsRead(string messageId, bool read) { }
            public void FlagMessage(string messageId, string flag) { }
            public void SetCategory(string messageId, string category) { }
        }

        [Fact]
        public async Task ListEvents_ReturnsTheCalendar_AndCapsThePeriod()
        {
            var calendar = new Calendar
            {
                Events = new[] { new CalendarEvent { Subject = "Budget review", Start = Sunday.AddHours(10), End = Sunday.AddHours(11), ShowAs = "busy", Attendees = "Sara Example" } },
            };

            var json = JObject.Parse(await new ListCalendarEventsTool().ExecuteAsync(
                "{\"date_from\":\"2026-11-15T00:00\",\"date_to\":\"2027-03-01T00:00\"}", calendar, CancellationToken.None));

            Assert.Equal("Budget review", (string)json["events"][0]["subject"]);
            Assert.Equal("2026-11-15T10:00", (string)json["events"][0]["start"]);
            Assert.Equal(Sunday.AddDays(31), calendar.AskedTo);
        }

        [Fact]
        public async Task CalendarTools_OnASurfaceWithoutACalendar_SayWhy()
        {
            var json = JObject.Parse(await new ListCalendarEventsTool().ExecuteAsync("{}", null, CancellationToken.None));

            Assert.Equal("invalid_arguments", (string)json["error"]["code"]);
        }

        [Fact]
        public async Task FindTimes_NamesCalendarsItCouldNotRead()
        {
            var calendar = new Calendar { FreeBusy = new Dictionary<string, string> { ["me"] = Grid(), ["sara@example.com"] = Grid(), ["ghost@example.com"] = null } };

            var json = JObject.Parse(await new FindMeetingTimesTool().ExecuteAsync(
                "{\"attendees\":[\"sara@example.com\",\"ghost@example.com\"],\"duration_minutes\":30,\"date_from\":\"2030-01-06T08:00\",\"date_to\":\"2030-01-07T17:00\"}",
                calendar, CancellationToken.None));

            Assert.Equal(new[] { "ghost@example.com" }, json["calendars_not_read"].Select(t => (string)t));
            Assert.Equal("local", (string)json["time_zone"]);
            Assert.True(json["slots"] is JArray);
        }

        [Theory]
        [InlineData("{\"duration_minutes\":5}")]
        [InlineData("{\"duration_minutes\":600}")]
        public async Task FindTimes_RefusesAnOddDuration(string args)
        {
            var json = JObject.Parse(await new FindMeetingTimesTool().ExecuteAsync(args, new Calendar(), CancellationToken.None));

            Assert.Equal("invalid_arguments", (string)json["error"]["code"]);
        }

        // The model can only propose; the page shows the times; a click opens Outlook's window.
        [Fact]
        public async Task ProposeMeeting_IsAProposal_CheckedAndCut()
        {
            var args = "{\"subject\":\"Budget\\nreview\",\"attendees\":[\"sara@example.com\",\"\",\"sara@example.com\"],\"agenda\":\"1. Numbers\"," +
                       "\"slots\":[{\"start\":\"2026-11-16T10:00\",\"end\":\"2026-11-16T10:30\"},{\"start\":\"2026-11-16T12:00\",\"end\":\"2026-11-16T11:00\"}," +
                       "{\"start\":\"bad\",\"end\":\"2026-11-16T11:00\"}]}";

            var result = await new ProposeMeetingTool().ExecuteAsync(args, null, CancellationToken.None);
            var json = JObject.Parse(result);

            Assert.Equal("meeting_proposal", (string)json["result_type"]);
            Assert.Equal("Budget review", (string)json["subject"]);
            Assert.Equal(new[] { "sara@example.com" }, json["attendees"].Select(t => (string)t));
            Assert.Single(json["slots"]);
            Assert.Contains("Nothing is sent", (string)json["status"]);

            var kept = MeetingProposal.FromResult(result);
            Assert.Equal(new DateTime(2026, 11, 16, 10, 0, 0), kept.Slots[0].Item1);
            Assert.Null(MeetingProposal.FromResult("{\"result_type\":\"memory_proposal\"}"));
            Assert.Equal("invalid_arguments", (string)JObject.Parse(await new ProposeMeetingTool().ExecuteAsync("{\"subject\":\"x\"}", null, CancellationToken.None))["error"]["code"]);
        }

        [Fact]
        public async Task ProposeMeeting_BuildsRoomDisplay_AndDropdownOptions()
        {
            var surface = new RoomAwareCalendar
            {
                PreferredBuilding = "HO1",
                Rooms = new[]
                {
                    new RoomInfo { Name = "Room Lexus HO2 FF Wing4", Email = "RoomHO2Lexus@example.com", Free = true },
                    new RoomInfo { Name = "Room Crown HO1 FF Wing4", Email = "RoomHO1Crown@example.com", Free = true },
                }
            };
            var args = "{\"subject\":\"Daily Loan Sync\",\"attendees\":[\"sara@example.com\"],\"room\":\"RoomHO1Crown@example.com\"," +
                       "\"slots\":[{\"start\":\"2026-11-18T10:00\",\"end\":\"2026-11-18T10:30\"}]}";

            var result = await new ProposeMeetingTool().ExecuteAsync(args, surface, CancellationToken.None);
            var json = JObject.Parse(result);

            Assert.Equal("Room Crown HO1 FF Wing4", (string)json["room_display"]);
            var opts = (JArray)json["room_options"];
            Assert.NotNull(opts);
            Assert.Equal("Room Crown HO1 FF Wing4", (string)opts[0]["label"]); // HO1 preferred first
            Assert.Contains(opts, o => (string)o["value"] == "Microsoft Teams Meeting");

            var kept = MeetingProposal.FromResult(result);
            Assert.Equal("RoomHO1Crown@example.com", kept.Room);
            Assert.Equal("Room Crown HO1 FF Wing4", kept.RoomOptions[0].Label);
            Assert.True(kept.RoomOptions.Count >= 3);
        }

        [Fact]
        public async Task FindTimes_DefaultsToOneHour_ExcludesBreak_AndSupportsOptionalAttendees()
        {
            // Busy everywhere except 11:30-14:30 on Monday 2030-01-07 (day 0 relative to date_from).
            // With 1h default and 12:30-13:30 break excluded, 12:00-13:00, 12:30-13:30, and 13:00-14:00 MUST NOT appear;
            // only 11:30-12:30 and 13:30-14:30 are valid!
            var busyMorning = (Slot(0, 0), Slot(0, 11, 30));
            var busyAfternoon = (Slot(0, 14, 30), Slot(0, 24));
            var me = Grid(busyMorning, busyAfternoon);
            var req = Grid(busyMorning, busyAfternoon);
            var optBusyAt1130 = Grid(busyMorning, (Slot(0, 11, 30), Slot(0, 12, 30)), busyAfternoon);

            var calendar = new Calendar
            {
                FreeBusy = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["me"] = me,
                    ["req@example.com"] = req,
                    ["cc@example.com"] = optBusyAt1130,
                }
            };

            // 1) Without checking optional availability: both 11:30-12:30 and 13:30-14:30 are returned (60 min, break skipped).
            var resSkipOpt = JObject.Parse(await new FindMeetingTimesTool().ExecuteAsync(
                "{\"attendees\":[\"req@example.com\"],\"optional_attendees\":[\"cc@example.com\"],\"check_optional_availability\":false,\"date_from\":\"2030-01-07T09:00\",\"date_to\":\"2030-01-07T17:00\"}",
                calendar, CancellationToken.None));
            Assert.Equal(60, (int)resSkipOpt["duration_minutes"]);
            Assert.Equal("12:30-13:30", (string)resSkipOpt["excluded_break"]);
            var skipSlots = resSkipOpt["slots"].Select(s => (string)s["start"] + "->" + (string)s["end"]).ToArray();
            Assert.Equal(new[] { "2030-01-07T11:30->2030-01-07T12:30", "2030-01-07T13:30->2030-01-07T14:30" }, skipSlots);

            // 2) With check_optional_availability=true: 11:30-12:30 is busy for cc@example.com, so only 13:30-14:30 is returned!
            var resCheckOpt = JObject.Parse(await new FindMeetingTimesTool().ExecuteAsync(
                "{\"attendees\":[\"req@example.com\"],\"optional_attendees\":[\"cc@example.com\"],\"check_optional_availability\":true,\"date_from\":\"2030-01-07T09:00\",\"date_to\":\"2030-01-07T17:00\"}",
                calendar, CancellationToken.None));
            var checkSlots = resCheckOpt["slots"].Select(s => (string)s["start"] + "->" + (string)s["end"]).ToArray();
            Assert.Equal(new[] { "2030-01-07T13:30->2030-01-07T14:30" }, checkSlots);
        }

        [Fact]
        public async Task ProposeMeeting_KeepsOptionalAttendees_SeparateFromRequired()
        {
            var args = "{\"subject\":\"Project Sync\",\"attendees\":[\"req@example.com\"],\"optional_attendees\":[\"cc1@example.com\",\"req@example.com\",\"cc2@example.com\"]," +
                       "\"slots\":[{\"start\":\"2026-11-18T10:00\",\"end\":\"2026-11-18T11:00\"}]}";
            var result = await new ProposeMeetingTool().ExecuteAsync(args, null, CancellationToken.None);
            var kept = MeetingProposal.FromResult(result);

            Assert.Equal(new[] { "req@example.com" }, kept.Attendees);
            Assert.Equal(new[] { "cc1@example.com", "cc2@example.com" }, kept.OptionalAttendees);
        }

        private sealed class RoomAwareCalendar : IOutlookSurface, IPeopleSurface, IOrgSurface
        {
            public string PreferredBuilding;
            public IReadOnlyList<RoomInfo> Rooms = new RoomInfo[0];
            public IReadOnlyList<PersonMatch> ResolvePeople(IReadOnlyList<string> names) => new PersonMatch[0];
            public IReadOnlyList<RoomInfo> FindRooms(DateTime start, DateTime end, string nameContains, int max) => Rooms;
            public OrgContextInfo GetOrgContext() => new OrgContextInfo { PreferredBuilding = PreferredBuilding };
            public ComposeStateResult GetCurrentComposeState(bool includeFullBody) => null;
            public IReadOnlyList<FolderResult> ListFolders() => null;
            public SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public MessageDetail ReadMessage(string messageId, bool includeFullBody) => null;
            public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken)) => null;
            public int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken)) => 0;
            public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) => null;
            public FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken)) => null;
            public CreatedDraft CreateDraft(CreateDraftArgs args) => null;
            public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) => null;
            public void MarkAsRead(string messageId, bool read) { }
            public void FlagMessage(string messageId, string flag) { }
            public void SetCategory(string messageId, string category) { }
        }

        [Fact]
        public void TheMeetingWindow_OpensOnlyFromTheHostsProposal_AndSendsNothing()
        {
            var surface = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "Services", "Tools", "LiveOutlookSurface.cs"));
            var method = surface.Substring(surface.IndexOf("public bool OpenMeetingRequest(", StringComparison.Ordinal));
            method = method.Substring(0, method.IndexOf("/// <summary>The mailbox owner's name", StringComparison.Ordinal));

            Assert.Contains("meeting.Display(false);", method);
            foreach (var verb in new[] { ".Send(", ".Save(", ".Delete(" }) Assert.DoesNotContain(verb, method);
            foreach (var controller in new[] { "ChatController.cs", "InboxCopilotController.cs" })
            {
                var dir = controller.StartsWith("Chat") ? "Chat" : "InboxCopilot";
                var source = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "TaskPane", dir, controller));
                Assert.Contains("_meetings.TryGetValue(callId, out var proposal)", source);
                Assert.Contains("MeetingProposal.FromResult(resultJson)", source);
            }
        }
    }
}
