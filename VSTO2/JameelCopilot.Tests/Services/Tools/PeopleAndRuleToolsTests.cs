using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Export;
using JameelCopilot.Services.Tools;
using JameelCopilot.TaskPane.Chat;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services.Tools
{
    /// <summary>
    /// Setting up a meeting with two first names, in a room: names become the
    /// address book's people, rooms come with whether they are free, and a
    /// rule is only ever proposed. The owner got a meeting request addressed
    /// to three bare first names and no answer about rooms.
    /// </summary>
    public class PeopleAndRuleToolsTests
    {
        private static readonly DateTime Sunday = new DateTime(2026, 11, 15);

        private sealed class Book : IPeopleSurface, IRuleSurface, IOutlookSurface
        {
            public Dictionary<string, PersonMatch> People = new Dictionary<string, PersonMatch>(StringComparer.OrdinalIgnoreCase);
            public IReadOnlyList<RoomInfo> Rooms = new RoomInfo[0];
            public Dictionary<string, string> Folders = new Dictionary<string, string>();
            public List<RuleProposal> Created = new List<RuleProposal>();
            public bool CreateFails;
            public DateTime AskedStart, AskedEnd;
            public string AskedContains;

            public IReadOnlyList<PersonMatch> ResolvePeople(IReadOnlyList<string> names) =>
                names.Select(n => People.TryGetValue(n, out var p) ? new PersonMatch { Query = n, Name = p.Name, Email = p.Email } : new PersonMatch { Query = n }).ToList();

            public IReadOnlyList<RoomInfo> FindRooms(DateTime start, DateTime end, string nameContains, int max)
            {
                AskedStart = start;
                AskedEnd = end;
                AskedContains = nameContains;
                return Rooms;
            }

            public string RuleFolderName(string folderId) => Folders.TryGetValue(folderId ?? "", out var name) ? name : null;
            public bool CreateRule(RuleProposal rule) { Created.Add(rule); return !CreateFails; }

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

        private static async Task<JObject> Run(IOutlookTool tool, string args, IOutlookSurface surface) =>
            JObject.Parse(await tool.ExecuteAsync(args, surface, CancellationToken.None));

        // -- people ------------------------------------------------------------

        [Fact]
        public async Task ResolvePeople_GivesTheAddressBooksPerson_AndSaysWhoWasNotFound()
        {
            var book = new Book();
            book.People["Sara"] = new PersonMatch { Name = "Sara Example", Email = "sara@example.com" };

            var json = await Run(new ResolvePeopleTool(), "{\"names\":[\"Sara\",\"Omar\",\" sara \",\"\"]}", book);

            var people = (JArray)json["people"];
            Assert.Equal(2, people.Count);                                  // the same name twice and the empty one are dropped
            Assert.Equal("Sara", (string)people[0]["asked"]);
            Assert.True((bool)people[0]["found"]);
            Assert.Equal("Sara Example", (string)people[0]["name"]);
            Assert.Equal("sara@example.com", (string)people[0]["email"]);
            Assert.False((bool)people[1]["found"]);
            Assert.Null(people[1]["email"]);                                // nothing the model could take for an address
            Assert.Contains("more than one person", (string)json["note"]);
        }

        [Fact]
        public async Task ResolvePeople_NeedsNames_AndAnAddressBook()
        {
            Assert.Equal("invalid_arguments", (string)(await Run(new ResolvePeopleTool(), "{}", new Book()))["error"]["code"]);
            Assert.Contains("not available", (string)(await Run(new ResolvePeopleTool(), "{\"names\":[\"Sara\"]}", null))["error"]["message"]);
        }

        // -- rooms -------------------------------------------------------------

        [Fact]
        public async Task FindRooms_ListsTheFreeRoomsFirst_ForTheTimeAsked()
        {
            var book = new Book
            {
                Rooms = new[]
                {
                    new RoomInfo { Name = "Room 1", Email = "room1@example.com", Free = false },
                    new RoomInfo { Name = "Room 2", Email = "room2@example.com", Free = null },
                    new RoomInfo { Name = "Room 3", Email = "room3@example.com", Free = true },
                },
            };

            var json = await Run(new FindRoomsTool(), "{\"start\":\"2026-11-15T10:00\",\"end\":\"2026-11-15T10:30\",\"name_contains\":\"Room\"}", book);

            Assert.Equal(new[] { "Room 3", "Room 2", "Room 1" }, json["rooms"].Select(r => (string)r["name"]));
            Assert.Equal(new[] { "yes", "unknown", "no" }, json["rooms"].Select(r => (string)r["free"]));
            Assert.Equal("room3@example.com", (string)json["rooms"][0]["email"]);
            Assert.Equal(Sunday.AddHours(10), book.AskedStart);
            Assert.Equal(Sunday.AddHours(10.5), book.AskedEnd);
            Assert.Equal("Room", book.AskedContains);
            Assert.Null(json["note"]);
        }

        // Not an error line on the page: the model is told to ask the user for the room.
        [Fact]
        public async Task FindRooms_WithNoRoomListInTheAddressBook_SaysToAskTheUser()
        {
            var json = await Run(new FindRoomsTool(), "{\"start\":\"2026-11-15T10:00\",\"end\":\"2026-11-15T10:30\"}", new Book { Rooms = null });

            Assert.Null(json["error"]);
            Assert.Empty(json["rooms"]);
            Assert.Contains("Ask the user for the room's name", (string)json["note"]);
        }

        [Theory]
        [InlineData("{}")]
        [InlineData("{\"start\":\"2026-11-15T10:00\"}")]
        [InlineData("{\"start\":\"2026-11-15T10:00\",\"end\":\"2026-11-15T09:00\"}")]
        [InlineData("{\"start\":\"2026-11-15T10:00\",\"end\":\"2026-11-17T10:00\"}")]
        public async Task FindRooms_NeedsATimeThatMakesSense(string args)
        {
            Assert.Equal("invalid_arguments", (string)(await Run(new FindRoomsTool(), args, new Book()))["error"]["code"]);
        }

        [Theory]
        [InlineData(10, 0, 10, 30, true)]      // the free half hour
        [InlineData(9, 30, 10, 30, false)]     // starts in the busy one
        [InlineData(10, 0, 11, 30, false)]     // runs into the next busy one
        [InlineData(10, 15, 10, 20, true)]     // inside one free slot
        public void IsFree_LooksAtEverySlotTheTimeTouches(int h1, int m1, int h2, int m2, bool expected)
        {
            var grid = new string('0', 48).ToCharArray();
            grid[19] = '2';   // 09:30-10:00
            grid[22] = '2';   // 11:00-11:30

            Assert.Equal(expected, MeetingTimes.IsFree(new string(grid), Sunday, Sunday.AddHours(h1).AddMinutes(m1), Sunday.AddHours(h2).AddMinutes(m2)));
        }

        [Fact]
        public void IsFree_IsUnknown_WithoutAGridThatCoversTheTime()
        {
            Assert.Null(MeetingTimes.IsFree(null, Sunday, Sunday.AddHours(10), Sunday.AddHours(11)));
            Assert.Null(MeetingTimes.IsFree(new string('0', 48), Sunday, Sunday.AddDays(2), Sunday.AddDays(2).AddHours(1)));
            Assert.Null(MeetingTimes.IsFree(new string('0', 48), Sunday, Sunday.AddHours(-2), Sunday.AddHours(1)));
        }

        [Theory]
        [InlineData("All Rooms", true)]
        [InlineData("Meeting rooms - Jeddah", true)]
        [InlineData("\u0642\u0627\u0639\u0627\u062A \u0627\u0644\u0627\u062C\u062A\u0645\u0627\u0639\u0627\u062A", true)]
        [InlineData("All Users", false)]
        [InlineData(null, false)]
        public void RoomList_IsKnownByItsName(string listName, bool expected)
        {
            Assert.Equal(expected, LiveOutlookSurface.IsRoomListName(listName));
        }

        [Fact]
        public async Task ProposeMeeting_CarriesTheRoom()
        {
            var result = await new ProposeMeetingTool().ExecuteAsync(
                "{\"subject\":\"Sync\",\"attendees\":[\"sara@example.com\"],\"room\":\"room3@example.com\",\"slots\":[{\"start\":\"2026-11-15T10:00\",\"end\":\"2026-11-15T10:30\"}]}",
                null, CancellationToken.None);

            Assert.Equal("room3@example.com", (string)JObject.Parse(result)["room"]);
            Assert.Equal("room3@example.com", MeetingProposal.FromResult(result).Room);
            Assert.Equal("", MeetingProposal.FromResult(result.Replace("room3@example.com", "")).Room);
        }

        // -- rules -------------------------------------------------------------

        private const string RuleArgs =
            "{\"name\":\"Vendor invoices\",\"from\":[\"billing@vendor.example\"],\"subject_contains\":[\"invoice\"],\"folder_id\":\"f1\",\"category\":\"Finance\"}";

        [Fact]
        public async Task ProposeRule_IsAProposal_WithOutlooksOwnFolderName()
        {
            var book = new Book();
            book.Folders["f1"] = "Invoices";

            var raw = await new ProposeRuleTool().ExecuteAsync(RuleArgs.Replace("}", ",\"folder\":\"Deleted Items\"}"), book, CancellationToken.None);
            var json = JObject.Parse(raw);
            var rule = RuleProposal.FromResult(raw);

            Assert.Equal("rule_proposal", (string)json["result_type"]);
            Assert.Equal("Invoices", (string)json["folder"]);          // never a name the model gave
            Assert.Contains("Nothing is created yet", (string)json["status"]);
            Assert.Empty(book.Created);
            Assert.Equal("Vendor invoices", rule.Name);
            Assert.Equal(new[] { "billing@vendor.example" }, rule.From);
            Assert.Equal(new[] { "invoice" }, rule.SubjectContains);
            Assert.Equal("f1", rule.FolderId);
            Assert.Equal("Invoices", rule.FolderName);
            Assert.Equal("Finance", rule.Category);
        }

        // Deleted Items and Junk answer null from the surface, like a folder Jameel Copilot never listed.
        [Fact]
        public async Task ProposeRule_RefusesAFolderARuleMayNotUse()
        {
            var json = await Run(new ProposeRuleTool(), RuleArgs, new Book());

            Assert.Equal("invalid_arguments", (string)json["error"]["code"]);
            Assert.Contains("not Deleted Items or Junk", (string)json["error"]["message"]);
        }

        [Theory]
        [InlineData("{\"from\":[\"a@example.com\"],\"category\":\"X\"}")]                      // no name
        [InlineData("{\"name\":\"R\",\"category\":\"X\"}")]                                    // no condition
        [InlineData("{\"name\":\"R\",\"from\":[\"a@example.com\"]}")]                          // no action
        [InlineData("{\"name\":\"R\",\"from\":\"a@example.com\",\"category\":\"X\"}")]         // from is not a list
        public async Task ProposeRule_NeedsAName_ACondition_AndAnAction(string args)
        {
            Assert.Equal("invalid_arguments", (string)(await Run(new ProposeRuleTool(), args, new Book()))["error"]["code"]);
        }

        [Fact]
        public void RuleProposal_IsReadOnlyFromItsOwnResult()
        {
            Assert.Null(RuleProposal.FromResult("{\"result_type\":\"meeting_proposal\",\"name\":\"R\",\"from\":[\"a\"],\"category\":\"X\"}"));
            Assert.Null(RuleProposal.FromResult("not json"));
            Assert.Null(RuleProposal.FromResult(null));
        }

        // The card's button names the card. The rule made is the one Jameel Copilot kept, and one click makes one rule.
        [Fact]
        public async Task CreateRule_MakesTheKeptRule_Once()
        {
            var book = new Book();
            book.Folders["f1"] = "Invoices";
            var scripts = new List<string>();
            var bridge = new RulesBridge(book, s => { scripts.Add(s); return Task.CompletedTask; });
            var result = await new ProposeRuleTool().ExecuteAsync(RuleArgs, book, CancellationToken.None);

            bridge.Offered("call1", ProposeRuleTool.ToolName, true, result);
            bridge.Offered("call2", ProposeMeetingTool.ToolName, true, result);   // another tool's result is no rule
            bridge.Offered("call3", ProposeRuleTool.ToolName, false, result);     // nor is a failed call's

            Assert.False(bridge.Handle("meeting_open", new JObject()));
            Assert.True(bridge.Handle("rule_create", new JObject(new JProperty("call", "call2"), new JProperty("name", "Forward everything"))));
            Assert.True(bridge.Handle("rule_create", new JObject(new JProperty("call", "call3"))));
            Assert.Empty(book.Created);

            Assert.True(bridge.Handle("rule_create", new JObject(new JProperty("call", "call1"))));
            Assert.True(bridge.Handle("rule_create", new JObject(new JProperty("call", "call1"))));

            Assert.Equal("Vendor invoices", Assert.Single(book.Created).Name);
            Assert.Equal("jameelcopilot.markRule(\"call1\", \"created - see Rules in Outlook\", true);", Assert.Single(scripts));
        }

        [Fact]
        public async Task CreateRule_ThatFails_SaysSo()
        {
            var book = new Book { CreateFails = true };
            var scripts = new List<string>();
            var bridge = new RulesBridge(book, s => { scripts.Add(s); return Task.CompletedTask; });
            bridge.Offered("c", ProposeRuleTool.ToolName, true,
                await new ProposeRuleTool().ExecuteAsync("{\"name\":\"R\",\"from\":[\"a@example.com\"],\"category\":\"X\"}", book, CancellationToken.None));

            bridge.Handle("rule_create", new JObject(new JProperty("call", "c")));

            Assert.Contains("could not be created", Assert.Single(scripts));
            Assert.EndsWith(", false);", scripts[0]);
        }
        // What a rule Jameel Copilot creates may do is fixed in the code: file the mail or label it. Never send it on or remove it.
        [Fact]
        public void TheRuleJcCreates_OnlyMovesOrLabels()
        {
            var surface = File.ReadAllText(RepoFiles.Find("VSTO2", "JameelCopilot", "Services", "Tools", "LiveOutlookSurface.cs"));
            var method = surface.Substring(surface.IndexOf("public bool CreateRule(", StringComparison.Ordinal));
            method = method.Substring(0, method.IndexOf("private T Run<T>(", StringComparison.Ordinal));

            Assert.Contains("Actions.MoveToFolder", method);
            Assert.Contains("Actions.AssignToCategory", method);
            Assert.Contains("RuleFolder(proposal.FolderId)", method);   // a folder Jameel Copilot listed, and never Deleted Items or Junk
            foreach (var other in new[] { "Forward", "Redirect", "Delete", "CopyToFolder", "olRuleSend" }) Assert.DoesNotContain(other, method);
        }
    }
}
