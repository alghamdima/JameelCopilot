using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using JameelCopilot.Services;
using JameelCopilot.Services.Export;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JameelCopilot.Services.Tools
{
    /// <summary>
    /// Production <see cref="IOutlookSurface"/> implementation. Wraps Outlook
    /// OOM with explicit STA marshalling via <see cref="OutlookThreadMarshaller"/>.
    /// Every public method blocks the calling thread (which is typically the
    /// chat-service task thread) while the OOM call runs on the Outlook UI
    /// thread. Each method catches <see cref="COMException"/> and returns
    /// graceful defaults (null/empty) so the tool layer can surface a
    /// structured <c>{"error":...}</c> back to the model.
    /// </summary>
    public sealed class LiveOutlookSurface : IOutlookSurface, ICalendarSurface, IPeopleSurface, IRuleSurface, IOrgSurface
    {
        private static OrgContextInfo _processCachedOrg;
        private static bool _processOrgLoaded;
        private static List<Tuple<string, string>> _processRoomIdentities;
        private static readonly object _cacheSync = new object();
        private static readonly Dictionary<string, PersonMatch> _personCache =
            new Dictionary<string, PersonMatch>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, Tuple<DateTime, string>> _freeBusyCache =
            new Dictionary<string, Tuple<DateTime, string>>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, Tuple<DateTime, string>> _roomGridCache =
            new Dictionary<string, Tuple<DateTime, string>>(StringComparer.OrdinalIgnoreCase);
        private OrgContextInfo _cachedOrg;
        private List<Tuple<string, string>> _cachedRoomIdentities;
        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;
        private readonly IdResolver _ids;
        private readonly Outlook.Inspector _composeInspector;
        private readonly Outlook.Explorer _explorer;     // Phase 3a
        private readonly IOutlookAdvancedSearchRunner _runner;
        private readonly IFolderClassifier _classifier;
        private static readonly TimeSpan _searchTimeout = TimeSpan.FromSeconds(30);
        // 32 KB hard cap on body bytes (in characters; close enough for ASCII).
        private const int MaxBodyChars = 32 * 1024;
        private const int MaxFolderDepth = 6;
        private const int SnippetChars = 160;

        /// <summary>
        /// Phase 3a / 3b primary constructor. Runner is required and shared
        /// process-wide so all AdvancedSearch invocations go through one
        /// serialiser and one event subscription.
        /// </summary>
        public LiveOutlookSurface(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            IdResolver ids,
            Outlook.Inspector composeInspector,
            Outlook.Explorer explorer,
            IOutlookAdvancedSearchRunner runner,
            IFolderClassifier classifier = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _ids = ids ?? throw new ArgumentNullException(nameof(ids));
            _composeInspector = composeInspector;
            _explorer = explorer;
            _runner = runner ?? throw new ArgumentNullException(nameof(runner));
            _classifier = classifier ?? new FolderClassifier();
        }

        public ComposeStateResult GetCurrentComposeState(bool includeFullBody) =>
            Run(() =>
            {
                if (_composeInspector == null) return EmptyCompose();
                var item = _composeInspector.CurrentItem as Outlook.MailItem;
                if (item == null) return EmptyCompose();

                var parts = ComposeDocument.Read(_composeInspector, item);
                int limit = includeFullBody ? MaxBodyChars : 1000;
                var body = parts.Own;
                var quoted = parts.Quoted;
                bool truncated = body.Length > limit || quoted.Length > limit;
                if (body.Length > limit) body = body.Substring(0, limit);
                if (quoted.Length > limit) quoted = quoted.Substring(0, limit);

                var attachments = new List<AttachmentSummary>();
                try
                {
                    foreach (Outlook.Attachment att in item.Attachments)
                    {
                        attachments.Add(new AttachmentSummary
                        {
                            Filename = att.FileName,
                            SizeBytes = att.Size,
                        });
                    }
                }
                catch (COMException) { /* ignore */ }

                InReplyTo reply = null;
                if (includeFullBody)
                {
                    try
                    {
                        var conv = item.GetConversation();
                        if (conv != null)
                        {
                            var msgs = new List<ThreadMessage>();
                            var table = conv.GetTable();
                            while (!table.EndOfTable)
                            {
                                var row = table.GetNextRow();
                                var entry = row["EntryID"]?.ToString();
                                if (string.IsNullOrEmpty(entry)) continue;
                                var threadItem = _application.Session.GetItemFromID(entry) as Outlook.MailItem;
                                if (threadItem == null) continue;
                                msgs.Add(new ThreadMessage
                                {
                                    From = threadItem.SenderName ?? threadItem.SenderEmailAddress ?? "",
                                    ReceivedAt = ToOffset(threadItem.ReceivedTime),
                                    Snippet = SnippetOf(threadItem.Body),
                                });
                                if (msgs.Count >= 5) break;
                            }
                            reply = new InReplyTo
                            {
                                ThreadTopic = item.ConversationTopic ?? "",
                                LastNMessages = msgs,
                            };
                        }
                    }
                    catch (COMException) { /* ignore */ }
                    catch (Exception) { /* defensive */ }
                }

                Outlook.AddressEntry sender = null;
                try { sender = _application.Session.CurrentUser?.AddressEntry; } catch (COMException) { }

                return new ComposeStateResult
                {
                    Subject = item.Subject ?? "",
                    ToRecipients = SplitAddresses(item.To),
                    CcRecipients = SplitAddresses(item.CC),
                    BccRecipients = SplitAddresses(item.BCC),
                    SenderName = sender?.Name ?? _application.Session.CurrentUser?.Name ?? "",
                    SenderEmail = TryGetSmtp(sender) ?? "",
                    BodyPlaintext = body,
                    BodyTruncated = truncated,
                    QuotedPlaintext = quoted,
                    Attachments = attachments,
                    InReplyTo = reply,
                };
            }) ?? EmptyCompose();

        /// <summary>
        /// Opens an email Jameel Copilot found or read, in its own Outlook window. Only an
        /// id Jameel Copilot itself gave out this session resolves, so the page cannot open
        /// anything Jameel Copilot did not list. False when it is gone or cannot be opened.
        /// </summary>
        public bool OpenMessage(string messageId) =>
            Run(() =>
            {
                if (string.IsNullOrEmpty(messageId)) return false;
                try
                {
                    var item = _application.Session.GetItemFromID(_ids.Resolve(messageId));
                    if (item is Outlook.MailItem mail) mail.Display(false);
                    else ((dynamic)item).Display(false);   // a meeting request or a report in the same folder
                    return true;
                }
                catch (KeyNotFoundException) { return false; }
                catch (COMException) { return false; }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return false; }
            });

        // -- the calendar --------------------------------------------------------

        /// <summary>
        /// The user's meetings and appointments that overlap the period, in order.
        /// Dates go into Outlook only in the invariant form the mail search uses,
        /// with a day's margin that the exact check here then removes; repeating
        /// meetings are expanded from their pattern, day by day, so no locale
        /// date string is needed for them.
        /// </summary>
        public IReadOnlyList<CalendarEvent> ListCalendarEvents(DateTime from, DateTime to, int max) =>
            Run(() =>
            {
                var found = new List<CalendarEvent>();
                var folder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var items = folder.Items;

                var single = items.Restrict("@SQL=\"urn:schemas:calendar:dtstart\" < '" +
                    to.AddDays(1).ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "' AND \"urn:schemas:calendar:dtend\" > '" +
                    from.AddDays(-1).ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "'");
                foreach (var item in single)
                {
                    if (item is Outlook.AppointmentItem a && !a.IsRecurring && a.Start < to && a.End > from) found.Add(ToEvent(a));
                    if (found.Count >= 200) break;
                }

                foreach (var item in items.Restrict("[IsRecurring] = True"))
                {
                    if (!(item is Outlook.AppointmentItem master)) continue;
                    Outlook.RecurrencePattern pattern;
                    try { pattern = master.GetRecurrencePattern(); }
                    catch (COMException) { continue; }
                    if (pattern.PatternEndDate.Date < from.Date || pattern.PatternStartDate.Date > to.Date) continue;
                    for (var day = from.Date; day <= to.Date; day = day.AddDays(1))
                    {
                        try
                        {
                            var occurrence = pattern.GetOccurrence(day + master.Start.TimeOfDay);
                            if (occurrence != null && occurrence.Start < to && occurrence.End > from) found.Add(ToEvent(occurrence));
                        }
                        catch (COMException) { /* no occurrence that day */ }
                    }
                    if (found.Count >= 200) break;
                }

                return (IReadOnlyList<CalendarEvent>)found.OrderBy(e => e.Start).Take(Math.Max(1, max)).ToList();
            });

        private static CalendarEvent ToEvent(Outlook.AppointmentItem a)
        {
            string showAs;
            switch (a.BusyStatus)
            {
                case Outlook.OlBusyStatus.olFree: showAs = "free"; break;
                case Outlook.OlBusyStatus.olTentative: showAs = "tentative"; break;
                case Outlook.OlBusyStatus.olOutOfOffice: showAs = "out_of_office"; break;
                case Outlook.OlBusyStatus.olWorkingElsewhere: showAs = "elsewhere"; break;
                default: showAs = "busy"; break;
            }
            return new CalendarEvent
            {
                Subject = a.Subject ?? "", Start = a.Start, End = a.End, Location = a.Location ?? "",
                Organizer = a.Organizer ?? "", Attendees = a.RequiredAttendees ?? "", AllDay = a.AllDayEvent, ShowAs = showAs,
            };
        }

        /// <summary>Free/busy from the day's midnight in 30-minute steps; "me" is the user. Null for a calendar that cannot be read.</summary>
        public IDictionary<string, string> GetFreeBusy(IEnumerable<string> people, DateTime day)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var nowUtc = DateTime.UtcNow;
            var dayBase = day.Date;
            result["me"] = GetCachedOrFetchFreeBusy("__me__", dayBase, nowUtc, () => Run(() => FreeBusyOf(_application.Session.CurrentUser, dayBase)));
            foreach (var person in people ?? new string[0])
            {
                var who = (person ?? "").Trim();
                if (who.Length == 0) continue;
                result[who] = GetCachedOrFetchFreeBusy(who, dayBase, nowUtc, () => Run(() =>
                {
                    try
                    {
                        var recipient = _application.Session.CreateRecipient(who);
                        return recipient.Resolve() ? FreeBusyOf(recipient, dayBase) : null;
                    }
                    catch (COMException) { return null; }
                }));
            }
            return result;
        }

        private static string FreeBusyOf(Outlook.Recipient recipient, DateTime day)
        {
            try { return recipient?.FreeBusy(day.Date, MeetingTimes.SlotMinutes, true); }
            catch (COMException) { return null; }
        }

        /// <summary>
        /// Opens Outlook's own meeting window, filled in, for the user to send or
        /// close. Nothing is sent or saved. Not a model tool: only a button the
        /// user presses calls it.
        /// </summary>
        public bool OpenMeetingRequest(MeetingProposal proposal, int slot) =>
            Run(() =>
            {
                if (proposal == null || slot < 0 || slot >= proposal.Slots.Count) return false;
                var meeting = (Outlook.AppointmentItem)_application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                meeting.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                meeting.Subject = proposal.Subject;
                meeting.Start = proposal.Slots[slot].Item1;
                meeting.End = proposal.Slots[slot].Item2;
                meeting.Body = proposal.Agenda ?? "";   // a new meeting with no text: nothing of the user's is replaced
                foreach (var person in proposal.Attendees ?? new string[0])
                {
                    var reqRec = meeting.Recipients.Add(person);
                    reqRec.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
                }
                foreach (var optPerson in proposal.OptionalAttendees ?? new string[0])
                {
                    var optRec = meeting.Recipients.Add(optPerson);
                    optRec.Type = (int)Outlook.OlMeetingRecipientType.olOptional;
                }
                Outlook.Recipient room = null;
                var wantsTeams = string.IsNullOrEmpty(proposal.Room) || IsOnlineLocation(proposal.Room);
                if (!wantsTeams)
                {
                    room = meeting.Recipients.Add(proposal.Room);
                    room.Type = (int)Outlook.OlMeetingRecipientType.olResource;
                    meeting.Location = proposal.Room;
                }
                else
                {
                    meeting.Location = "Microsoft Teams Meeting";
                }
                try
                {
                    meeting.Recipients.ResolveAll();
                    if (room != null && room.Resolved) meeting.Location = room.Name;
                }
                catch (COMException) { /* Outlook shows what it could not resolve */ }
                if (wantsTeams) TrySetOnlineMeetingProvider(meeting);
                meeting.Display(false);
                if (wantsTeams) ScheduleTeamsLinkGeneration(meeting);
                return true;
            });

        /// <summary>The mailbox owner's name and address: who Jameel Copilot writes as.</summary>
        public void GetOwner(out string name, out string email)
        {
            var owner = Run(() =>
            {
                Outlook.AddressEntry entry = null;
                try { entry = _application.Session.CurrentUser?.AddressEntry; } catch (COMException) { }
                return new[] { entry?.Name ?? "", TryGetSmtp(entry) ?? "" };
            });
            name = owner[0];
            email = owner[1];
        }

        /// <summary>
        /// Opens Outlook's own reply window for a message, with the text above
        /// the signature. Nothing is sent or saved: the user reads it and
        /// decides. Not a model tool: only a button the user presses calls it.
        /// </summary>
        public bool OpenReplyWithText(string messageId, bool replyAll, string text) =>
            Run(() =>
            {
                var original = _application.Session.GetItemFromID(_ids.Resolve(messageId)) as Outlook.MailItem;
                if (original == null) return false;
                var reply = replyAll ? original.ReplyAll() : original.Reply();
                var inspector = reply.GetInspector;
                reply.Display(false);
                return ComposeDocument.InsertAtTop(inspector, text);
            });

        public IReadOnlyList<FolderResult> ListFolders() =>
            Run(() =>
            {
                var results = new List<FolderResult>();
                try
                {
                    foreach (Outlook.Store store in _application.Session.Stores)
                    {
                        WalkFolders(store.GetRootFolder(), null, results, depth: 0);
                        if (results.Count >= SearchFallbackBudget.MaxListFolders) break;
                    }
                }
                catch (COMException) { }
                return (IReadOnlyList<FolderResult>)results;
            }) ?? (IReadOnlyList<FolderResult>)new List<FolderResult>();

        public SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
        {
            args = args ?? new SearchMessagesArgs();
            var filter = BuildRestrictFilter(args);
            var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();

            // For all_mail / auto without an explicit folder, skip the
            // AdvancedSearch primary path entirely. Application.AdvancedSearch
            // empirically throws COMException("Sorry, something went wrong")
            // on multi-store scopes, and building the multi-store scope
            // costs ~10 seconds of UI-thread enumeration that the fallback
            // would have to redo anyway. AdvancedSearch stays the right
            // tool for narrow scopes (specific folder); for all-mail it
            // is pure waste.
            // For auto without an explicit folder: try current_folder first via fast Table API!
            // Only expand across all 44+ folders when current_folder finds nothing (or scope == all_mail).
            if (string.IsNullOrEmpty(args.FolderId) && scopeMode == "auto")
            {
                var localHits = FallbackIterativeSearch(args, filter, "current_folder", ct);
                if (localHits != null && localHits.Messages != null && localHits.Messages.Count > 0)
                {
                    return localHits;
                }
            }

            bool useAdvancedSearch =
                (!string.IsNullOrEmpty(args.FolderId) || scopeMode == "current_folder")
                && string.IsNullOrEmpty(args.Query)
                && string.IsNullOrEmpty(args.BodyContains);

            if (useAdvancedSearch)
            {
                // Step 1: resolve scope on UI thread.
                SearchScope scope;
                try
                {
                    scope = _marshaller.RunAsync(() => BuildSearchScope(args, scopeMode, ct), ct).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }

                JameelCopilot.Diagnostics.TraceLog.Write(
                    "SearchMessages primary=AdvancedSearch start scope_paths=" + (scope.ResolvedFolderPaths?.Count ?? 0)
                    + " sub=" + scope.SearchSubFolders + " filter=" + (string.IsNullOrEmpty(filter) ? "<none>" : filter),
                    "LiveOutlookSurface");

                // Step 2: try AdvancedSearch.
                AdvancedSearchResult primary;
                try
                {
                    primary = _runner.RunAsync(scope.ScopeString, filter, scope.SearchSubFolders, _searchTimeout, ct)
                        .GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }

                if (primary.Cancelled)
                {
                    JameelCopilot.Diagnostics.TraceLog.Write("SearchMessages primary cancelled by user", "LiveOutlookSurface");
                    throw new OperationCanceledException(ct);
                }

                if (primary.Completed && primary.Items != null)
                {
                    JameelCopilot.Diagnostics.TraceLog.Write(
                        "SearchMessages primary=AdvancedSearch complete raw_count=" + primary.Items.Count,
                        "LiveOutlookSurface");
                    return _marshaller.RunAsync(
                        () => SearchResultProjector.Project(primary.Items, args, _classifier),
                        ct).GetAwaiter().GetResult();
                }

                var reason = primary.TimedOut ? "timeout" :
                             primary.Error != null ? "com:" + primary.Error.GetType().Name :
                             "null_results";
                JameelCopilot.Diagnostics.TraceLog.Write("SearchMessages fallback reason=" + reason, "LiveOutlookSurface");
            }
            else
            {
                JameelCopilot.Diagnostics.TraceLog.Write(
                    "SearchMessages skipping primary=AdvancedSearch (scope=" + scopeMode + " multi-store fallback path)",
                    "LiveOutlookSurface");
            }

            return FallbackIterativeSearch(args, filter, scopeMode, ct);
        }

        public MessageDetail ReadMessage(string messageId, bool includeFullBody)
        {
            var tempFiles = new List<Tuple<AttachmentSummary, string>>();
            var detail = Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    var item = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                    if (item == null) return null;

                    var rawBody = item.Body ?? "";
                    var attachments = SaveSupportedAttachmentsToTemp(item, includeFullBody, tempFiles);

                    return new MessageDetail
                    {
                        Id = messageId,
                        Subject = item.Subject ?? "",
                        From = FormatSender(item.SenderName, item.SenderEmailAddress),
                        To = SplitAddresses(item.To),
                        Cc = SplitAddresses(item.CC),
                        ReceivedAt = ToOffset(item.ReceivedTime),
                        BodyPlaintext = rawBody,
                        BodyTruncated = false,
                        Attachments = attachments,
                        InReplyToMessageId = null,
                        ConversationTopic = item.ConversationTopic ?? "",
                    };
                }
                catch (COMException) { return null; }
                catch (KeyNotFoundException) { return null; }
            });

            if (detail == null)
            {
                CleanupTempAttachments(tempFiles);
                return null;
            }

            // Parse body signatures and extract DOCX/XLSX/PPTX/PDF text OFF the UI thread!
            var cleanBody = ComposeBody.StripThreadSignatures(detail.BodyPlaintext);
            bool truncated = false;
            if (cleanBody.Length > MaxBodyChars)
            {
                cleanBody = cleanBody.Substring(0, MaxBodyChars);
                truncated = true;
            }
            if (!includeFullBody && cleanBody.Length > 1000)
            {
                cleanBody = cleanBody.Substring(0, 1000);
                truncated = true;
            }
            detail.BodyPlaintext = cleanBody;
            detail.BodyTruncated = truncated;
            ExtractAndCleanupTempAttachments(tempFiles);
            return detail;
        }

        public int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
        {
            // CRITICAL perf path. The Phase 3b version routed counts
            // through SearchMessages with MaxResults=int.MaxValue, which
            // built a MessageProjectionInput per mail item across all
            // matching folders. On the user's 200-folder mailbox a
            // single count_messages call could touch ~thousands of items
            // and freeze Outlook for 30+ seconds. Outlook's resiliency
            // detector noticed and threatened to auto-disable the add-in.
            //
            // Items.Count (or Items.Restrict(...).Count) is O(1) metadata
            // most stores serve from cache. Use it. Skip-list still
            // applied at the folder level via ResolveSearchFolders /
            // WalkMailFolders + IFolderClassifier.
            args = args ?? new SearchMessagesArgs();
            var filter = BuildRestrictFilter(args);
            var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();
            var total = 0;

            IReadOnlyList<Outlook.MAPIFolder> folders;
            try
            {
                folders = _marshaller.RunAsync(
                    () => ResolveSearchFolders(
                        SearchFallbackBudget.CountFolderResolutionArgs(args),
                        allMail: scopeMode != "current_folder",
                        ct),
                    ct).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            foreach (var folder in folders)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    _marshaller.RunAsync(() =>
                    {
                        ct.ThrowIfCancellationRequested();
                        Outlook.Items items;
                        try
                        {
                            items = string.IsNullOrEmpty(filter)
                                ? folder.Items
                                : folder.Items.Restrict(filter);
                        }
                        catch (COMException) { return; }
                        try { total += items.Count; } catch (COMException) { }
                        YieldUi(ct);
                    }, ct).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }
            }
            return total;
        }

        public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
        {
            if (ids == null || ids.Length == 0) return new MessageDetail[0];
            if (maxItems < 1) maxItems = 1;
            if (maxItems > 100) maxItems = 100;

            var capped = ids.Take(maxItems).ToArray();
            var results = new List<MessageDetail>(capped.Length);

            _marshaller.RunAsync(() =>
            {
                foreach (var shortId in capped)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var entryId = _ids.Resolve(shortId);
                        var item = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                        if (item == null) continue;
                        results.Add(BuildMessageDetail(shortId, item, includeBody));
                    }
                    catch (COMException ex)
                    {
                        try { JameelCopilot.Diagnostics.TraceLog.Write("ReadMessages COMException id=" + shortId + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {
                        try { JameelCopilot.Diagnostics.TraceLog.Write("ReadMessages unknown short id=" + shortId, "LiveOutlookSurface"); } catch { }
                    }
                    // Pump Outlook UI between items so a long ids[] does not
                    // freeze the UI thread for the full read sweep.
                    YieldUi(ct);
                }
            }, ct).GetAwaiter().GetResult();

            return results;
        }

        // Builds a MessageDetail for a single MailItem. Mirrors ReadMessage's
        // projection so the bulk and single tools return identical shapes.
        private MessageDetail BuildMessageDetail(string shortId, Outlook.MailItem item, bool includeBody)
        {
            var body = "";
            bool truncated = false;
            if (includeBody)
            {
                try { body = ComposeBody.StripThreadSignatures(item.Body ?? ""); } catch (COMException) { }
                if (body.Length > MaxBodyChars)
                {
                    body = body.Substring(0, MaxBodyChars);
                    truncated = true;
                }
            }

            var attachments = new List<AttachmentSummary>();
            try
            {
                foreach (Outlook.Attachment att in item.Attachments)
                {
                    attachments.Add(new AttachmentSummary
                    {
                        Filename = att.FileName,
                        SizeBytes = att.Size,
                    });
                }
            }
            catch (COMException) { }

            string subject = ""; try { subject = item.Subject ?? ""; } catch (COMException) { }
            string sender = ""; try { sender = item.SenderName ?? item.SenderEmailAddress ?? ""; } catch (COMException) { }
            string to = ""; try { to = item.To ?? ""; } catch (COMException) { }
            string cc = ""; try { cc = item.CC ?? ""; } catch (COMException) { }
            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try { receivedAt = ToOffset(item.ReceivedTime); } catch (COMException) { }
            string conversationTopic = ""; try { conversationTopic = item.ConversationTopic ?? ""; } catch (COMException) { }

            return new MessageDetail
            {
                Id = shortId,
                Subject = subject,
                From = sender,
                To = SplitAddresses(to),
                Cc = SplitAddresses(cc),
                ReceivedAt = receivedAt,
                BodyPlaintext = body,
                BodyTruncated = truncated,
                Attachments = attachments,
                InReplyToMessageId = null,
                ConversationTopic = conversationTopic,
            };
        }

        public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
        {
            args = args ?? new AggregateMessagesArgs();
            var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();
            var filter = BuildAggregateFilter(args);
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Step 1: resolve folder list on UI thread (reuse the existing
            // method used by the search fallback).
            IReadOnlyList<Outlook.MAPIFolder> folders;
            try
            {
                folders = _marshaller.RunAsync(
                    () => ResolveSearchFolders(
                        new SearchMessagesArgs { FolderId = args.FolderId },
                        allMail: scopeMode != "current_folder",
                        ct),
                    ct).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            // Step 2: per-folder Table API read, classifier-filter mail rows,
            // group into the dictionary. One marshalled call per folder so the
            // UI thread is released between folders.
            foreach (var folder in folders)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    _marshaller.RunAsync(
                        () => AccumulateFolderBuckets(folder, args, filter, counts, ct),
                        ct).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }
            }

            // Step 3: convert dictionary to AggregationBucket list and clamp
            // via TopNBucketSelector for deterministic ordering.
            var allBuckets = counts.Select(kv => new AggregationBucket { Label = kv.Key, Count = kv.Value }).ToList();
            return TopNBucketSelector.TakeTop(allBuckets, args.TopN);
        }

        // Builds a DASL @SQL filter for the aggregate query. Mirrors the
        // existing BuildRestrictFilter clauses for date_from / date_to / from /
        // subject_contains / body_contains. Property names are double-quoted
        // for folder.GetTable() compatibility (same Phase 3b bug).
        // Returns null if no clauses.
        private static string BuildAggregateFilter(AggregateMessagesArgs args)
        {
            var clauses = new List<string>();
            if (!string.IsNullOrEmpty(args.From))
            {
                var f = (args.From ?? "").Replace("'", "''");
                clauses.Add("(" + PropFrom + " LIKE '%" + f + "%' OR " + PropFromEmail + " LIKE '%" + f + "%' OR " + PropFromSmtp + " LIKE '%" + f + "%')");
            }
            if (!string.IsNullOrEmpty(args.SubjectContains))
            {
                clauses.Add(PropSubject + " LIKE '%" + args.SubjectContains.Replace("'", "''") + "%'");
            }
            if (!string.IsNullOrEmpty(args.BodyContains))
            {
                clauses.Add(PropBody + " LIKE '%" + args.BodyContains.Replace("'", "''") + "%'");
            }
            if (args.DateFrom.HasValue)
            {
                clauses.Add(PropReceivedAt + " >= '" + args.DateFrom.Value.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture) + "'");
            }
            if (args.DateTo.HasValue)
            {
                clauses.Add(PropReceivedAt + " <= '" + args.DateTo.Value.ToString("M/d/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture) + "'");
            }
            if (clauses.Count == 0) return null;
            return "@SQL=" + string.Join(" AND ", clauses);
        }

        // Read one folder's rows via the Table API, classify mail-only, and
        // accumulate counts into the shared dictionary keyed by args.GroupBy.
        private void AccumulateFolderBuckets(
            Outlook.MAPIFolder folder,
            AggregateMessagesArgs args,
            string filter,
            Dictionary<string, int> counts,
            CancellationToken ct)
        {
            if (folder == null) return;

            string folderName = "";
            bool folderIsMail = true;
            try { folderName = folder.Name ?? ""; } catch (COMException) { }
            try { folderIsMail = folder.DefaultItemType == Outlook.OlItemType.olMailItem; } catch (COMException) { }
            if (_classifier.IsSystemFolder(folderName, folderIsMail)) return;

            Outlook.Table table;
            try
            {
                table = folder.GetTable(filter ?? "", Outlook.OlTableContents.olUserItems);
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("AggregateMessages GetTable COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
                return;
            }

            try
            {
                table.Columns.RemoveAll();
                table.Columns.Add("SenderName");
                table.Columns.Add("SenderEmailAddress");
                table.Columns.Add("ReceivedTime");
                table.Columns.Add("MessageClass");
            }
            catch (COMException) { }

            int rowsScanned = 0;
            try
            {
                while (!table.EndOfTable)
                {
                    ct.ThrowIfCancellationRequested();
                    rowsScanned++;
                    Outlook.Row row;
                    try { row = table.GetNextRow(); }
                    catch (COMException) { break; }
                    if (row == null) break;

                    string messageClass = "";
                    try { messageClass = row["MessageClass"] as string ?? ""; } catch (COMException) { }
                    if (!TableMessageClassFilter.IsMailMessage(messageClass)) continue;

                    string key = ResolveBucketKey(row, folderName, args.GroupBy);
                    if (key == null) continue;

                    int existing;
                    counts.TryGetValue(key, out existing);
                    counts[key] = existing + 1;

                    // Pump UI every ~50 rows so a big folder doesn't freeze
                    // the UI for the whole folder scan.
                    if ((rowsScanned % 50) == 0) YieldUi(ct);
                }
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("AggregateMessages row scan COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
            }
        }

        // Picks the bucket key for one row given args.GroupBy.
        private static string ResolveBucketKey(Outlook.Row row, string folderName, string groupBy)
        {
            if (string.Equals(groupBy, "folder", StringComparison.OrdinalIgnoreCase))
            {
                return string.IsNullOrEmpty(folderName) ? null : folderName;
            }

            if (string.Equals(groupBy, "day", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var rt = row["ReceivedTime"];
                    if (rt is DateTime dt) return DateBucketFormatter.Format(new DateTimeOffset(dt));
                }
                catch (COMException) { }
                return DateBucketFormatter.UnknownDate;
            }

            // Default: sender
            string name = "";
            string email = "";
            try { name = row["SenderName"] as string ?? ""; } catch (COMException) { }
            try { email = row["SenderEmailAddress"] as string ?? ""; } catch (COMException) { }
            return SenderKeyNormalizer.Normalize(name, email);
        }

        public IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) =>
            Run(() =>
            {
                var results = new List<ThreadSummary>();
                if (string.IsNullOrEmpty(recipientEmail)) return (IReadOnlyList<ThreadSummary>)results;
                try
                {
                    var byConvId = new Dictionary<string, ThreadSummary>(StringComparer.Ordinal);
                    foreach (var defaultFolder in new[] { Outlook.OlDefaultFolders.olFolderInbox, Outlook.OlDefaultFolders.olFolderSentMail })
                    {
                        Outlook.MAPIFolder folder = null;
                        try { folder = _application.Session.GetDefaultFolder(defaultFolder); }
                        catch (COMException) { continue; }
                        if (folder == null) continue;

                        var filter = "@SQL=" +
                            "(urn:schemas:httpmail:fromemail LIKE '%" + Escape(recipientEmail) + "%' " +
                            "OR urn:schemas:httpmail:displayto LIKE '%" + Escape(recipientEmail) + "%' " +
                            "OR urn:schemas:httpmail:displaycc LIKE '%" + Escape(recipientEmail) + "%')";
                        Outlook.Items items;
                        try { items = folder.Items.Restrict(filter); }
                        catch (COMException) { continue; }
                        try { items.Sort("[ReceivedTime]", true); } catch (COMException) { }

                        int scanned = 0;
                        foreach (var obj in items)
                        {
                            if (scanned++ >= 200) break;
                            var mi = obj as Outlook.MailItem;
                            if (mi == null) continue;
                            var convId = mi.ConversationID ?? mi.ConversationTopic ?? mi.EntryID;
                            if (string.IsNullOrEmpty(convId)) continue;
                            if (!byConvId.TryGetValue(convId, out var s))
                            {
                                s = new ThreadSummary
                                {
                                    ThreadTopic = mi.ConversationTopic ?? "",
                                    LastMessageAt = ToOffset(mi.ReceivedTime),
                                    MessageCount = 0,
                                    Snippet = SnippetOf(mi.Body),
                                    ThreadId = _ids.Shorten(convId),
                                };
                                byConvId[convId] = s;
                            }
                            s.MessageCount++;
                            var ra = ToOffset(mi.ReceivedTime);
                            if (ra > s.LastMessageAt) s.LastMessageAt = ra;
                        }
                    }
                    results = byConvId.Values
                        .OrderByDescending(t => t.LastMessageAt)
                        .Take(maxThreads)
                        .ToList();
                }
                catch (COMException) { }
                return (IReadOnlyList<ThreadSummary>)results;
            }) ?? (IReadOnlyList<ThreadSummary>)new List<ThreadSummary>();

        public FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
        {
            if (args == null) throw new ArgumentNullException(nameof(args));

            ct.ThrowIfCancellationRequested();
            var pathResolver = new ExportPathResolver();
            var baseDir = pathResolver.ResolveBaseDir();
            try
            {
                pathResolver.EnsureExists();
            }
            catch (Exception ex) when (IsExpectedPathUnavailable(ex))
            {
                throw new ExportException("path_unavailable", ex.Message, ex);
            }

            var filename = ExportFilenameSanitizer.Build(args.FilenameHint, ".xlsx", DateTimeOffset.Now, candidate => File.Exists(Path.Combine(baseDir, candidate)));
            var fullPath = Path.Combine(baseDir, filename);

            try
            {
                var typedRows = new List<object[]>();
                for (var r = 0; r < args.Rows.Count; r++)
                {
                    ct.ThrowIfCancellationRequested();
                    var row = new object[args.Columns.Count];
                    for (var c = 0; c < args.Columns.Count; c++)
                    {
                        ct.ThrowIfCancellationRequested();
                        row[c] = ExcelCellCoercer.Coerce(args.Rows[r][c], args.Columns[c].Type);
                    }
                    typedRows.Add(row);
                }

                using (var workbook = ExcelWorkbookBuilder.Build(args.SheetName, args.Columns, typedRows))
                {
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        ct.ThrowIfCancellationRequested();
                        stream.Position = 0;
                        try
                        {
                            SaveNewFile(stream, fullPath);
                        }
                        catch (IOException ex) when (IsSharingViolation(ex) || IsAlreadyExists(ex))
                        {
                            filename = ExportFilenameSanitizer.Build(args.FilenameHint + "-2", ".xlsx", DateTimeOffset.Now, candidate => File.Exists(Path.Combine(baseDir, candidate)));
                            fullPath = Path.Combine(baseDir, filename);
                            ct.ThrowIfCancellationRequested();
                            stream.Position = 0;
                            try
                            {
                                SaveNewFile(stream, fullPath);
                            }
                            catch (IOException retryEx) when (IsSharingViolation(retryEx))
                            {
                                throw new ExportException("file_locked", retryEx.Message, retryEx);
                            }
                            catch (IOException retryEx) when (IsDiskFull(retryEx))
                            {
                                throw new ExportException("disk_full", retryEx.Message, retryEx);
                            }
                            catch (IOException retryEx)
                            {
                                throw new ExportException("excel_build_failed", retryEx.Message, retryEx);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (ExportException) { throw; }
            catch (IOException ex) when (IsDiskFull(ex))
            {
                throw new ExportException("disk_full", ex.Message, ex);
            }
            catch (IOException ex) when (IsSharingViolation(ex))
            {
                throw new ExportException("file_locked", ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new ExportException("excel_build_failed", ex.Message, ex);
            }

            var info = new FileInfo(fullPath);
            return new FileSavedResult
            {
                Path = fullPath,
                FileUrl = new Uri(fullPath).AbsoluteUri,
                Format = "xlsx",
                Bytes = info.Length,
                Filename = filename,
            };
        }

        public FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
        {
            if (args == null) throw new ArgumentNullException(nameof(args));

            ct.ThrowIfCancellationRequested();
            var pathResolver = Globals.ThisAddIn?.ExportPathResolver ?? new ExportPathResolver();
            string baseDir;
            try
            {
                baseDir = pathResolver.ResolveBaseDir();
                pathResolver.EnsureExists();
            }
            catch (Exception ex) when (IsExpectedPathUnavailable(ex))
            {
                throw new ExportException("path_unavailable", ex.Message, ex);
            }

            var generatedAt = DateTimeOffset.Now;
            var filename = ExportFilenameSanitizer.Build(args.FilenameHint, ".pdf", generatedAt, candidate => File.Exists(Path.Combine(baseDir, candidate)));
            var fullPath = Path.Combine(baseDir, filename);
            var tempPath = Path.Combine(baseDir, "." + Guid.NewGuid().ToString("N") + ".tmp.pdf");

            try
            {
                ct.ThrowIfCancellationRequested();
                var template = LoadPrintTemplateHtml();
                var html = new PrintTemplateRenderer(template).Render(args.Title ?? args.FilenameHint, args.Subtitle, args.ContentMarkdown, generatedAt);

                ct.ThrowIfCancellationRequested();
                var pdfRenderer = Globals.ThisAddIn?.PdfRenderer;
                if (pdfRenderer == null)
                {
                    throw new ExportException("pdf_render_failed", "PdfRenderer not initialized");
                }

                pdfRenderer.RenderAsync(html, tempPath, ct).GetAwaiter().GetResult();
                ct.ThrowIfCancellationRequested();

                try
                {
                    File.Move(tempPath, fullPath);
                }
                catch (IOException ex) when (IsSharingViolation(ex) || IsAlreadyExists(ex))
                {
                    filename = ExportFilenameSanitizer.Build(args.FilenameHint + "-2", ".pdf", DateTimeOffset.Now, candidate => File.Exists(Path.Combine(baseDir, candidate)));
                    fullPath = Path.Combine(baseDir, filename);
                    ct.ThrowIfCancellationRequested();
                    File.Move(tempPath, fullPath);
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (ExportException) { throw; }
            catch (IOException ex) when (IsDiskFull(ex))
            {
                throw new ExportException("disk_full", ex.Message, ex);
            }
            catch (IOException ex) when (IsSharingViolation(ex))
            {
                throw new ExportException("file_locked", ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new ExportException("pdf_render_failed", ex.Message, ex);
            }
            finally
            {
                TryDeleteTempFile(tempPath);
            }

            var info = new FileInfo(fullPath);
            return new FileSavedResult
            {
                Path = fullPath,
                FileUrl = new Uri(fullPath).AbsoluteUri,
                Format = "pdf",
                Bytes = info.Length,
                Filename = filename,
            };
        }

        public CreatedDraft CreateDraft(CreateDraftArgs args) =>
            Run(() =>
            {
                try
                {
                    Outlook.MailItem draft;
                    if (!string.IsNullOrEmpty(args.InReplyToMessageId))
                    {
                        var entryId = _ids.Resolve(args.InReplyToMessageId);
                        var original = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                        draft = original?.Reply()
                                ?? (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);
                    }
                    else
                    {
                        draft = (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);
                    }

                    draft.Subject = args.Subject ?? "";
                    draft.Body = args.BodyPlaintext ?? "";
                    if (args.To != null && args.To.Count > 0)
                        draft.To = string.Join("; ", args.To);
                    if (args.Cc != null && args.Cc.Count > 0)
                        draft.CC = string.Join("; ", args.Cc);
                    draft.Save();
                    return new CreatedDraft
                    {
                        DraftId = _ids.Shorten(draft.EntryID),
                        Location = "Drafts",
                    };
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException("CreateDraft failed: " + ex.Message, ex);
                }
            });

        public void MarkAsRead(string messageId, bool read) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        item.UnRead = !read;
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        public void FlagMessage(string messageId, string flag) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        // FlagStatus property numeric values:
                        //   olNoFlag = 0, olFlagComplete = 1, olFlagMarked = 2.
                        // "todo" => olFlagMarked, "complete" => olFlagComplete, "none" => olNoFlag.
                        switch (flag)
                        {
                            case "todo": item.FlagStatus = Outlook.OlFlagStatus.olFlagMarked; break;
                            case "complete": item.FlagStatus = Outlook.OlFlagStatus.olFlagComplete; break;
                            default: item.FlagStatus = Outlook.OlFlagStatus.olNoFlag; break;
                        }
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        public void SetCategory(string messageId, string category) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        item.Categories = category ?? "";
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        private Outlook.Explorer GetLiveExplorer()
        {
            if (_composeInspector != null) return null;
            if (_explorer != null)
            {
                try
                {
                    var _ = _explorer.CurrentFolder;
                    return _explorer;
                }
                catch (COMException) { }
            }
            try { return _application.ActiveExplorer(); }
            catch (COMException) { return null; }
        }

        public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
        {
            var tempFiles = new List<Tuple<AttachmentSummary, string>>();
            var result = Run(() =>
            {
                var explorer = GetLiveExplorer();
                if (explorer == null)
                {
                    return new CurrentSelectionResult
                    {
                        Folder = "",
                        FolderId = "",
                        Count = 0,
                        Messages = new MessageDetail[0],
                    };
                }

                string folderName = "";
                string folderId = "";
                try
                {
                    var folder = explorer.CurrentFolder;
                    if (folder != null)
                    {
                        folderName = folder.Name ?? "";
                        folderId = _ids.Shorten(folder.EntryID ?? "");
                    }
                }
                catch (COMException) { }

                Outlook.Selection selection = null;
                int totalCount = 0;
                try
                {
                    selection = explorer.Selection;
                    if (selection != null) totalCount = selection.Count;
                }
                catch (COMException) { selection = null; totalCount = 0; }

                var picked = new List<MessageDetail>();
                if (selection != null && totalCount > 0)
                try
                {
                    int taken = 0;
                    // Outlook Selection is 1-based.
                    for (int i = 1; i <= totalCount && taken < maxItems; i++)
                    {
                        object item = null;
                        try { item = selection[i]; } catch (COMException) { continue; }
                        var mi = item as Outlook.MailItem;
                        if (mi == null) continue;

                        var rawBody = mi.Body ?? "";
                        if (!includeFullBodies && rawBody.Length > 2500) rawBody = rawBody.Substring(0, 2500);
                        var atts = SaveSupportedAttachmentsToTemp(mi, includeFullBodies, tempFiles);

                        picked.Add(new MessageDetail
                        {
                            Id = _ids.Shorten(mi.EntryID ?? ""),
                            Subject = mi.Subject ?? "",
                            From = FormatSender(mi.SenderName, mi.SenderEmailAddress),
                            To = SplitAddresses(mi.To),
                            Cc = SplitAddresses(mi.CC),
                            ReceivedAt = ToOffset(mi.ReceivedTime),
                            BodyPlaintext = rawBody,
                            BodyTruncated = false,
                            Attachments = atts,
                            InReplyToMessageId = null,
                            ConversationTopic = mi.ConversationTopic ?? "",
                        });
                        taken++;
                    }
                }
                catch (COMException) { }

                return new CurrentSelectionResult
                {
                    Folder = folderName,
                    FolderId = folderId,
                    Count = totalCount,
                    Messages = picked,
                };
            });

            // Heavy regex signature stripping & attachment file extraction run OFF the UI thread.
            foreach (var msg in result?.Messages ?? new MessageDetail[0])
            {
                var body = includeFullBodies
                    ? ComposeBody.StripThreadSignatures(msg.BodyPlaintext ?? "")
                    : ComposeBody.StripSignatureAndFooter(ComposeBody.FromPlainText(msg.BodyPlaintext ?? "").Own);
                bool truncated = false;
                if (!includeFullBodies && body.Length > 1000)
                {
                    body = body.Substring(0, 1000);
                    truncated = true;
                }
                else if (body.Length > MaxBodyChars)
                {
                    body = body.Substring(0, MaxBodyChars);
                    truncated = true;
                }
                msg.BodyPlaintext = body;
                msg.BodyTruncated = truncated;
            }
            ExtractAndCleanupTempAttachments(tempFiles);
            return result;
        }

        // ---------- helpers ----------

        // -- people and rooms --------------------------------------------------

        /// <summary>
        /// Each name as Outlook's own name check resolves it. A name with no
        /// match, or with several, comes back without an address: Outlook
        /// gives no list to choose from here.
        /// </summary>
        public IReadOnlyList<PersonMatch> ResolvePeople(IReadOnlyList<string> names)
        {
            var result = new List<PersonMatch>();
            foreach (var rawName in names ?? new string[0])
            {
                var query = (rawName ?? "").Trim();
                if (query.Length == 0)
                {
                    result.Add(new PersonMatch { Query = rawName });
                    continue;
                }
                PersonMatch cached = null;
                lock (_cacheSync)
                {
                    _personCache.TryGetValue(query, out cached);
                }
                if (cached != null)
                {
                    result.Add(new PersonMatch
                    {
                        Query = rawName,
                        Name = cached.Name,
                        Email = cached.Email,
                        JobTitle = cached.JobTitle,
                        Department = cached.Department,
                    });
                    continue;
                }

                // Resolve ONE person per UI-thread post so Outlook never stalls across a batch of names.
                var match = Run(() =>
                {
                    var m = new PersonMatch { Query = rawName };
                    try
                    {
                        var recipient = _application.Session.CreateRecipient(query);
                        if (recipient.Resolve())
                        {
                            var entry = recipient.AddressEntry;
                            m.Name = entry?.Name ?? recipient.Name ?? "";
                            m.Email = TryGetSmtp(entry) ?? CleanAddress(recipient.Address);
                            try
                            {
                                var exUser = entry?.GetExchangeUser();
                                if (exUser != null)
                                {
                                    m.JobTitle = CleanJobTitle(exUser.JobTitle);
                                    m.Department = (exUser.Department ?? "").Trim();
                                }
                            }
                            catch (COMException) { }
                        }
                    }
                    catch (COMException) { /* not found */ }
                    return m;
                });
                if (!string.IsNullOrEmpty(match.Email))
                {
                    lock (_cacheSync)
                    {
                        _personCache[query] = match;
                        if (!string.IsNullOrEmpty(match.Name)) _personCache[match.Name] = match;
                        _personCache[match.Email] = match;
                    }
                }
                result.Add(match);
            }
            return result;
        }

        /// <summary>
        /// The rooms of the address book's room list ("All Rooms" in Exchange)
        /// and whether each is free for the time. Falls back to resolving the
        /// company's known meeting rooms through Outlook when the address list
        /// cannot be enumerated.
        /// </summary>
        public IReadOnlyList<RoomInfo> FindRooms(DateTime start, DateTime end, string nameContains, int max)
        {
            // Step 1: discover room names & SMTP addresses once per process.
            if (_cachedRoomIdentities == null) _cachedRoomIdentities = _processRoomIdentities;
            if (_cachedRoomIdentities == null)
            {
                _cachedRoomIdentities = Run(() =>
                {
                    var discovered = new List<Tuple<string, string>>();
                    Outlook.AddressList rooms = null;
                    try
                    {
                        foreach (Outlook.AddressList list in _application.Session.AddressLists)
                        {
                            string listName;
                            try { listName = list.Name ?? ""; } catch (COMException) { continue; }
                            if (string.Equals(listName, "All Rooms", StringComparison.OrdinalIgnoreCase)) { rooms = list; break; }
                            if (rooms == null && IsRoomListName(listName)) rooms = list;
                        }
                    }
                    catch (COMException) { }

                    if (rooms != null)
                    {
                        int looked = 0;
                        try
                        {
                            foreach (Outlook.AddressEntry entry in rooms.AddressEntries)
                            {
                                if (++looked > 300) break;
                                string name;
                                try { name = entry.Name ?? ""; } catch (COMException) { continue; }
                                if (IsExcludedRoom(name)) continue;
                                discovered.Add(Tuple.Create(name, TryGetSmtp(entry) ?? ""));
                            }
                        }
                        catch (COMException) { }
                    }

                    if (discovered.Count == 0)
                    {
                        foreach (var known in KnownRoomNames)
                        {
                            try
                            {
                                var recipient = _application.Session.CreateRecipient(known);
                                if (!recipient.Resolve()) continue;
                                var entry = recipient.AddressEntry;
                                var name = entry?.Name ?? recipient.Name ?? known;
                                if (IsExcludedRoom(name)) continue;
                                discovered.Add(Tuple.Create(name, TryGetSmtp(entry) ?? CleanAddress(recipient.Address) ?? ""));
                            }
                            catch (COMException) { }
                        }
                    }
                    return discovered.Count > 0 || rooms != null ? discovered : null;
                });
                _processRoomIdentities = _cachedRoomIdentities;
            }

            if (_cachedRoomIdentities == null) return null;

            // Step 2: check free/busy ONE room per UI-thread post, reusing the 30-day grid returned by Exchange!
            var found = new List<RoomInfo>();
            var nowUtc = DateTime.UtcNow;
            foreach (var room in _cachedRoomIdentities)
            {
                if (found.Count >= max) break;
                var roomName = room.Item1;
                var roomEmail = room.Item2;
                if (!string.IsNullOrEmpty(nameContains) && roomName.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) < 0) continue;

                var lookupKey = !string.IsNullOrEmpty(roomEmail) ? roomEmail : roomName;
                var grid = GetCachedOrFetchRoomGrid(lookupKey, start.Date, nowUtc, out var gridBase);
                found.Add(new RoomInfo { Name = roomName, Email = roomEmail, Free = MeetingTimes.IsFree(grid, gridBase, start, end) });
            }
            return found;
        }

        private static readonly string[] KnownRoomNames =
        {
            "Room Crown HO1 FF Wing4",
            "Room Supra HO1 FF Wing4",
            "Room Lexus HO2 FF Wing4",
            "Room Avalon HO2 FF Wing3",
            "Room Camry HO2 FF Wing4",
            "Room Corolla HO2 FF Wing4",
            "Room Fortuner HO2 GF Wing2",
            "Room Prado HO2 GF Wing2",
            "Room Highlander HO2 GF Wing1",
            "Room Yaris HO2 FF Wing3",
            "Room HO2 Teleconference",
        };

        internal static bool IsExcludedRoom(string name)
        {
            var n = name ?? "";
            return n.IndexOf("Test Room", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("ACCL", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static bool IsOnlineLocation(string room)
        {
            var r = (room ?? "").Trim();
            return r.IndexOf("Teams", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(r, "Online", StringComparison.OrdinalIgnoreCase);
        }

        // An address list of rooms under another name: "Rooms", "Meeting rooms", or the Arabic for rooms and halls.
        internal static bool IsRoomListName(string name)
        {
            var n = name ?? "";
            return n.IndexOf("room", StringComparison.OrdinalIgnoreCase) >= 0 || n.Contains("\u063A\u0631\u0641") || n.Contains("\u0642\u0627\u0639\u0627\u062A");
        }

        // -- waiting for a reply, tasks, rules ----------------------------------

        /// <summary>
        /// Emails the user sent in the period that have no reply in their
        /// conversation, newest first. A store that keeps no conversations
        /// gives none: without them nothing can be said.
        /// </summary>
        public IReadOnlyList<MessageSummary> ListAwaitingReply(DateTime from, DateTime to, int max) =>
            Run(() =>
            {
                var found = new List<MessageSummary>();
                var folder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                // The date goes in the invariant form with a day's margin; the exact check is below.
                var items = folder.Items.Restrict("@SQL=\"urn:schemas:httpmail:date\" >= '" +
                    from.AddDays(-1).ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "'");
                items.Sort("[SentOn]", true);
                var looked = 0;
                foreach (var item in items)
                {
                    if (found.Count >= max) break;
                    if (!(item is Outlook.MailItem mail)) continue;
                    try
                    {
                        var sent = mail.SentOn;
                        if (sent > to) continue;                    // too recent to count as unanswered
                        if (sent < from || ++looked > 120) break;   // newest first: the rest is older
                        var conversation = mail.GetConversation();
                        if (conversation == null || conversation.GetChildren(mail).Count > 0) continue;   // answered, or followed up already
                        var text = StructuredAnswers.OneLine(mail.Body);
                        found.Add(new MessageSummary
                        {
                            Id = _ids.Shorten(mail.EntryID),
                            Subject = mail.Subject ?? "",
                            From = mail.SenderName ?? "",
                            To = SplitAddresses(mail.To ?? ""),
                            ReceivedAt = ToOffset(sent),
                            Snippet = text.Length > 400 ? text.Substring(0, 400) : text,
                            HasAttachments = mail.Attachments.Count > 0,
                        });
                    }
                    catch (COMException) { /* an item that cannot be read is left out */ }
                }
                return (IReadOnlyList<MessageSummary>)found;
            });

        /// <summary>
        /// Adds a task to the user's Outlook tasks. Not a model tool: only the
        /// "Add to tasks" button of a to-do the user reads calls it.
        /// </summary>
        public bool AddTask(string subject, DateTime? due, string note) =>
            Run(() =>
            {
                if (string.IsNullOrWhiteSpace(subject)) return false;
                try
                {
                    var task = (Outlook.TaskItem)_application.CreateItem(Outlook.OlItemType.olTaskItem);
                    task.Subject = subject;
                    if (due.HasValue) task.DueDate = due.Value.Date;
                    task.Body = note ?? "";
                    task.Save();
                    return true;
                }
                catch (COMException) { return false; }
            });

        public string RuleFolderName(string folderId) =>
            Run(() =>
            {
                var name = SafeFolderName(RuleFolder(folderId));
                return name.Length > 0 ? name : null;
            });

        public OrgContextInfo GetOrgContext()
        {
            if (_cachedOrg != null) return _cachedOrg;
            if (_processOrgLoaded) return _cachedOrg = _processCachedOrg;
            _cachedOrg = Run(() =>
            {
                try
                {
                    var currentEntry = _application.Session.CurrentUser?.AddressEntry;
                    var me = currentEntry?.GetExchangeUser();
                    if (me == null) return null;

                    var mePerson = ToOrgPerson(me);
                    Outlook.ExchangeUser mgr = null;
                    try { mgr = me.GetExchangeUserManager(); } catch (COMException) { }
                    var mgrPerson = mgr != null ? ToOrgPerson(mgr) : null;
                    YieldUi(CancellationToken.None);

                    var reports = new List<OrgPerson>();
                    try
                    {
                        var drList = me.GetDirectReports();
                        if (drList != null)
                        {
                            int idx = 0;
                            foreach (Outlook.AddressEntry entry in drList)
                            {
                                if (reports.Count >= 30) break;
                                var p = ToRealEmployee(entry);
                                if (p != null) reports.Add(p);
                                if ((++idx % 4) == 0) YieldUi(CancellationToken.None);
                            }
                        }
                    }
                    catch (COMException) { }

                    var peers = new List<OrgPerson>();
                    try
                    {
                        var peerList = mgr?.GetDirectReports();
                        if (peerList != null)
                        {
                            int idx = 0;
                            foreach (Outlook.AddressEntry entry in peerList)
                            {
                                if (peers.Count >= 25) break;
                                var p = ToRealEmployee(entry);
                                if (p != null && !string.Equals(p.Email, mePerson.Email, StringComparison.OrdinalIgnoreCase))
                                {
                                    peers.Add(p);
                                }
                                if ((++idx % 4) == 0) YieldUi(CancellationToken.None);
                            }
                        }
                    }
                    catch (COMException) { }

                    var isTechGroup = IsTechDivision(mePerson, mgrPerson, false);
                    if (!isTechGroup)
                    {
                        try
                        {
                            var groups = me.GetMemberOfList();
                            if (groups != null)
                            {
                                foreach (Outlook.AddressEntry g in groups)
                                {
                                    string gName = "";
                                    try { gName = g.Name ?? ""; } catch (COMException) { }
                                    if (string.Equals(gName, "IT", StringComparison.OrdinalIgnoreCase)
                                        || gName.IndexOf("ALJUF IT", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        isTechGroup = true;
                                        break;
                                    }
                                }
                            }
                        }
                        catch (COMException) { }
                    }

                    var building = InferPreferredBuilding(mePerson, mgrPerson, isTechGroup);
                    return new OrgContextInfo
                    {
                        User = mePerson,
                        Manager = mgrPerson,
                        DirectReports = reports,
                        Peers = peers,
                        PreferredBuilding = building,
                    };
                }
                catch (COMException) { return null; }
            });
            _processCachedOrg = _cachedOrg;
            _processOrgLoaded = true;
            return _cachedOrg;
        }

        internal static string CleanJobTitle(string raw)
        {
            var t = (raw ?? "").Trim();
            while (t.StartsWith("Alj.", StringComparison.OrdinalIgnoreCase))
            {
                t = t.Substring(4).TrimStart();
            }
            return t;
        }

        private static OrgPerson ToOrgPerson(Outlook.ExchangeUser u)
        {
            if (u == null) return null;
            string name = "", email = "", title = "", dept = "";
            try { name = u.Name ?? ""; } catch (COMException) { }
            try { email = u.PrimarySmtpAddress ?? ""; } catch (COMException) { }
            try { title = CleanJobTitle(u.JobTitle); } catch (COMException) { }
            try { dept = (u.Department ?? "").Trim(); } catch (COMException) { }
            return new OrgPerson { Name = name, Email = email, JobTitle = title, Department = dept };
        }

        private static OrgPerson ToRealEmployee(Outlook.AddressEntry entry)
        {
            if (entry == null) return null;
            try
            {
                var u = entry.GetExchangeUser();
                if (u == null) return null;
                var p = ToOrgPerson(u);
                if (string.IsNullOrEmpty(p.Email)) return null;
                if (string.IsNullOrEmpty(p.JobTitle) && string.IsNullOrEmpty(p.Department)) return null;
                if (p.Name.EndsWith(" admin", StringComparison.OrdinalIgnoreCase)) return null;
                return p;
            }
            catch (COMException) { return null; }
        }

        private string InferPreferredBuilding(OrgPerson me, OrgPerson manager, bool isTechGroup)
        {
            if (IsTechDivision(me, manager, isTechGroup)) return "HO1";
            try
            {
                var cal = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var from = DateTime.Today.AddDays(-45);
                var items = cal.Items.Restrict("@SQL=\"urn:schemas:calendar:dtstart\" >= '" +
                    from.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "'");
                int ho1 = 0, ho2 = 0, scanned = 0;
                foreach (var item in items)
                {
                    if (++scanned > 40) break;
                    if (!(item is Outlook.AppointmentItem a)) continue;
                    var loc = a.Location ?? "";
                    if (loc.IndexOf("HO1", StringComparison.OrdinalIgnoreCase) >= 0) ho1++;
                    else if (loc.IndexOf("HO2", StringComparison.OrdinalIgnoreCase) >= 0) ho2++;
                }
                if (ho1 > ho2) return "HO1";
                if (ho2 > ho1) return "HO2";
            }
            catch (COMException) { }
            return "";
        }

        internal static bool IsTechDivision(OrgPerson me, OrgPerson manager, bool isTechGroup)
        {
            if (isTechGroup) return true;
            var dept = me?.Department ?? "";
            if (dept == "72901") return true;
            var combined = ((me?.JobTitle ?? "") + " " + (manager?.JobTitle ?? "")).ToLowerInvariant();
            return combined.Contains("it ") || combined.EndsWith(" it") || combined.StartsWith("it")
                || combined.Contains("information technology") || combined.Contains("technology")
                || combined.Contains("data ") || combined.Contains("pmo") || combined.Contains("project management");
        }

        // A folder a rule may file mail in: one Jameel Copilot listed, and neither Deleted Items nor Junk.
        private Outlook.MAPIFolder RuleFolder(string folderId)
        {
            if (string.IsNullOrEmpty(folderId)) return null;
            var folder = ResolveFolder(folderId);
            if (folder == null) return null;
            try
            {
                foreach (var kind in new[] { Outlook.OlDefaultFolders.olFolderDeletedItems, Outlook.OlDefaultFolders.olFolderJunk })
                {
                    if (_application.Session.CompareEntryIDs(_application.Session.GetDefaultFolder(kind).EntryID, folder.EntryID)) return null;
                }
            }
            catch (COMException) { return null; }
            return folder;
        }

        /// <summary>
        /// Creates the rule for incoming mail. Senders Outlook knows become the
        /// rule's "from" people; with one it does not know, all of them are
        /// matched as words in the sender's address instead, because a rule's
        /// conditions must all hold.
        /// </summary>
        public bool CreateRule(RuleProposal proposal) =>
            Run(() =>
            {
                if (proposal == null) return false;
                try
                {
                    var rules = _application.Session.DefaultStore.GetRules();
                    var rule = rules.Create(proposal.Name, Outlook.OlRuleType.olRuleReceive);

                    var from = proposal.From ?? new string[0];
                    if (from.Count > 0)
                    {
                        // Checked before anything goes into the rule: an unresolved name must not stay in it.
                        var allKnown = from.All(who =>
                        {
                            try { return _application.Session.CreateRecipient(who).Resolve(); }
                            catch (COMException) { return false; }
                        });
                        if (allKnown)
                        {
                            var people = rule.Conditions.From;
                            foreach (var who in from) people.Recipients.Add(who);
                            people.Recipients.ResolveAll();
                            people.Enabled = true;
                        }
                        else
                        {
                            rule.Conditions.SenderAddress.Address = from.Cast<object>().ToArray();
                            rule.Conditions.SenderAddress.Enabled = true;
                        }
                    }
                    if (proposal.SubjectContains != null && proposal.SubjectContains.Count > 0)
                    {
                        rule.Conditions.Subject.Text = proposal.SubjectContains.Cast<object>().ToArray();
                        rule.Conditions.Subject.Enabled = true;
                    }
                    if (!string.IsNullOrEmpty(proposal.FolderId))
                    {
                        var folder = RuleFolder(proposal.FolderId);
                        if (folder == null) return false;
                        rule.Actions.MoveToFolder.Folder = folder;
                        rule.Actions.MoveToFolder.Enabled = true;
                    }
                    if (!string.IsNullOrEmpty(proposal.Category))
                    {
                        rule.Actions.AssignToCategory.Categories = new object[] { proposal.Category };
                        rule.Actions.AssignToCategory.Enabled = true;
                    }
                    rules.Save(false);
                    return true;
                }
                catch (COMException) { return false; }
            });

        private T Run<T>(Func<T> fn) => _marshaller.RunAsync(fn, CancellationToken.None).GetAwaiter().GetResult();

        private void Run(Action fn) => _marshaller.RunAsync(fn, CancellationToken.None).GetAwaiter().GetResult();

        private static List<AttachmentSummary> SaveSupportedAttachmentsToTemp(
            Outlook.MailItem item, bool saveForExtraction, List<Tuple<AttachmentSummary, string>> savedTempFiles)
        {
            var list = new List<AttachmentSummary>();
            if (item == null) return list;
            int savedCount = 0;
            try
            {
                foreach (Outlook.Attachment att in item.Attachments)
                {
                    string fileName = "";
                    long size = 0;
                    try { fileName = att.FileName ?? ""; } catch (COMException) { }
                    try { size = att.Size; } catch (COMException) { }
                    var summary = new AttachmentSummary { Filename = fileName, SizeBytes = size };
                    if (saveForExtraction && savedTempFiles != null && savedCount < 4
                        && size > 0 && size <= AttachmentTextExtractor.MaxAttachmentBytes
                        && AttachmentTextExtractor.IsSupportedExtension(fileName))
                    {
                        var ext = Path.GetExtension(fileName) ?? ".tmp";
                        var tempFile = Path.Combine(Path.GetTempPath(), "jc-att-" + Guid.NewGuid().ToString("N") + ext);
                        try
                        {
                            att.SaveAsFile(tempFile);
                            savedTempFiles.Add(Tuple.Create(summary, tempFile));
                            savedCount++;
                        }
                        catch
                        {
                            TryDeleteTempFile(tempFile);
                        }
                    }
                    list.Add(summary);
                }
            }
            catch (COMException) { }
            return list;
        }

        private static void ExtractAndCleanupTempAttachments(List<Tuple<AttachmentSummary, string>> savedTempFiles)
        {
            if (savedTempFiles == null) return;
            foreach (var pair in savedTempFiles)
            {
                try
                {
                    pair.Item1.ContentPlaintext = AttachmentTextExtractor.ExtractFromFile(pair.Item2, pair.Item1.Filename);
                }
                catch { }
                finally
                {
                    TryDeleteTempFile(pair.Item2);
                }
            }
        }

        private static void CleanupTempAttachments(List<Tuple<AttachmentSummary, string>> savedTempFiles)
        {
            if (savedTempFiles == null) return;
            foreach (var pair in savedTempFiles) TryDeleteTempFile(pair.Item2);
        }

        private static void SaveNewFile(Stream source, string fullPath)
        {
            using (var file = new FileStream(fullPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                source.CopyTo(file);
            }
        }

        private static string LoadPrintTemplateHtml()
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("JameelCopilot.WebUI.print-template.html"))
            {
                if (stream == null)
                {
                    throw new ExportException("pdf_render_failed", "print-template.html resource missing");
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static void TryDeleteTempFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch { }
        }

        private static bool IsExpectedPathUnavailable(Exception ex)
        {
            return ex is IOException
                || ex is UnauthorizedAccessException
                || ex is NotSupportedException
                || ex is ArgumentException;
        }

        private static bool IsSharingViolation(IOException ex)
        {
            return ((uint)ex.HResult & 0xFFFF) == 32;
        }

        private static bool IsAlreadyExists(IOException ex)
        {
            var code = (uint)ex.HResult & 0xFFFF;
            return code == 80 || code == 183;
        }

        private static bool IsDiskFull(IOException ex)
        {
            var code = (uint)ex.HResult & 0xFFFF;
            return code == 39 || code == 112;
        }

        private static ComposeStateResult EmptyCompose() => new ComposeStateResult
        {
            Subject = "",
            ToRecipients = new string[0],
            CcRecipients = new string[0],
            BccRecipients = new string[0],
            SenderName = "",
            SenderEmail = "",
            BodyPlaintext = "",
            BodyTruncated = false,
            QuotedPlaintext = "",
            Attachments = new AttachmentSummary[0],
        };

        private static IReadOnlyList<string> SplitAddresses(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return new string[0];
            return raw.Split(';')
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        private static DateTimeOffset ToOffset(DateTime dt)
        {
            try
            {
                if (dt.Kind == DateTimeKind.Utc) return new DateTimeOffset(dt, TimeSpan.Zero);
                if (dt.Kind == DateTimeKind.Local) return new DateTimeOffset(dt);
                return new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Local));
            }
            catch { return DateTimeOffset.MinValue; }
        }

        private static string SnippetOf(string body)
        {
            if (string.IsNullOrEmpty(body)) return "";
            var collapsed = body.Replace("\r\n", " ").Replace('\n', ' ').Trim();
            return collapsed.Length <= SnippetChars
                ? collapsed
                : collapsed.Substring(0, SnippetChars) + "...";
        }

        private static string TryGetSmtp(Outlook.AddressEntry entry)
        {
            if (entry == null) return null;
            try { return entry.GetExchangeUser()?.PrimarySmtpAddress ?? CleanAddress(entry.Address); }
            catch (COMException) { return null; }
        }

        internal static string CleanAddress(string address)
        {
            var a = (address ?? "").Trim();
            if (a.Length == 0 || a.StartsWith("/O=", StringComparison.OrdinalIgnoreCase) || a.StartsWith("/o=", StringComparison.Ordinal))
                return null;
            return a;
        }

        internal static string FormatSender(string senderName, string senderAddress)
        {
            var name = (senderName ?? "").Trim();
            var smtp = CleanAddress(senderAddress);
            if (string.IsNullOrEmpty(smtp) || string.Equals(name, smtp, StringComparison.OrdinalIgnoreCase))
                return name.Length > 0 ? name : (senderAddress ?? "");
            return name.Length > 0 ? (name + " <" + smtp + ">") : smtp;
        }

        private static string GetCachedOrFetchFreeBusy(string key, DateTime dayBase, DateTime nowUtc, Func<string> fetch)
        {
            var cacheKey = key + "|" + dayBase.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            lock (_cacheSync)
            {
                if (_freeBusyCache.TryGetValue(cacheKey, out var hit) && (nowUtc - hit.Item1).TotalMinutes < 3)
                    return hit.Item2;
            }
            var grid = fetch();
            if (grid != null)
            {
                lock (_cacheSync) { _freeBusyCache[cacheKey] = Tuple.Create(nowUtc, grid); }
            }
            return grid;
        }

        private string GetCachedOrFetchRoomGrid(string lookupKey, DateTime requestedDate, DateTime nowUtc, out DateTime gridBase)
        {
            lock (_cacheSync)
            {
                // Exchange FreeBusy returns ~30 days (1,440 30-min chars) in one call: reuse any cached grid covering requestedDate!
                foreach (var kvp in _roomGridCache)
                {
                    if (!kvp.Key.StartsWith(lookupKey + "|", StringComparison.OrdinalIgnoreCase)) continue;
                    if ((nowUtc - kvp.Value.Item1).TotalMinutes >= 5 || string.IsNullOrEmpty(kvp.Value.Item2)) continue;
                    var datePart = kvp.Key.Substring(lookupKey.Length + 1);
                    if (DateTime.TryParseExact(datePart, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var cachedBase))
                    {
                        var spanDays = (requestedDate.Date - cachedBase.Date).TotalDays;
                        if (spanDays >= 0 && (spanDays + 1) * 48 <= kvp.Value.Item2.Length)
                        {
                            gridBase = cachedBase.Date;
                            return kvp.Value.Item2;
                        }
                    }
                }
            }

            var baseDay = requestedDate.Date < DateTime.Today ? DateTime.Today : DateTime.Today;
            if ((requestedDate.Date - baseDay).TotalDays > 25) baseDay = requestedDate.Date;
            var fetched = Run(() =>
            {
                try
                {
                    var r = _application.Session.CreateRecipient(lookupKey);
                    return r.Resolve() ? r.FreeBusy(baseDay, MeetingTimes.SlotMinutes, true) : null;
                }
                catch (COMException) { return null; }
            });
            gridBase = baseDay;
            if (fetched != null)
            {
                var storeKey = lookupKey + "|" + baseDay.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                lock (_cacheSync) { _roomGridCache[storeKey] = Tuple.Create(nowUtc, fetched); }
            }
            return fetched;
        }

        private static void TrySetOnlineMeetingProvider(Outlook.AppointmentItem meeting)
        {
            try
            {
                // olOnlineMeetingProviderTeams = 5 in Outlook 16+ IDispatch (triggers Teams EMO when enabled).
                meeting.GetType().InvokeMember(
                    "OnlineMeetingProvider",
                    BindingFlags.SetProperty,
                    null,
                    meeting,
                    new object[] { 5 });
            }
            catch { /* Older Outlook builds ignore this property; UIAutomation fallback handles it below. */ }
        }

        private void ScheduleTeamsLinkGeneration(Outlook.AppointmentItem meeting)
        {
            var timer = new System.Windows.Forms.Timer { Interval = 480 };
            int attempts = 0;
            timer.Tick += (s, e) =>
            {
                attempts++;
                try
                {
                    var body = meeting.Body ?? "";
                    if (body.IndexOf("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        timer.Stop();
                        timer.Dispose();
                        return;
                    }
                    Outlook.Inspector insp = null;
                    try { insp = meeting.GetInspector; } catch { }
                    if (insp != null && TryInvokeTeamsRibbonButton(insp.Caption))
                    {
                        timer.Stop();
                        timer.Dispose();
                        return;
                    }
                }
                catch { }
                if (attempts >= 3)
                {
                    timer.Stop();
                    timer.Dispose();
                }
            };
            timer.Start();
        }

        private static bool TryInvokeTeamsRibbonButton(string captionHint)
        {
            var cap = captionHint ?? "";
            var invoked = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var root = System.Windows.Automation.AutomationElement.RootElement;
                    if (root == null) return;
                    var windows = root.FindAll(
                        System.Windows.Automation.TreeScope.Children,
                        new System.Windows.Automation.PropertyCondition(
                            System.Windows.Automation.AutomationElement.ControlTypeProperty,
                            System.Windows.Automation.ControlType.Window));
                    int currentPid = System.Diagnostics.Process.GetCurrentProcess().Id;
                    foreach (System.Windows.Automation.AutomationElement win in windows)
                    {
                        int pid = 0;
                        try { pid = win.Current.ProcessId; } catch { continue; }
                        if (pid != currentPid) continue;
                        string winName = "";
                        try { winName = win.Current.Name ?? ""; } catch { }
                        if (!string.IsNullOrEmpty(cap) && winName.IndexOf(cap, StringComparison.OrdinalIgnoreCase) < 0
                            && winName.IndexOf("Meeting", StringComparison.OrdinalIgnoreCase) < 0
                            && winName.IndexOf("\u0627\u062C\u062A\u0645\u0627\u0639", StringComparison.Ordinal) < 0)
                            continue;

                        var btns = win.FindAll(
                            System.Windows.Automation.TreeScope.Descendants,
                            new System.Windows.Automation.PropertyCondition(
                                System.Windows.Automation.AutomationElement.ControlTypeProperty,
                                System.Windows.Automation.ControlType.Button));
                        foreach (System.Windows.Automation.AutomationElement btn in btns)
                        {
                            string bName = "", autoId = "";
                            try { bName = btn.Current.Name ?? ""; } catch { }
                            try { autoId = btn.Current.AutomationId ?? ""; } catch { }
                            var isTeamsBtn = autoId.IndexOf("Appointment_NewMeeting", StringComparison.OrdinalIgnoreCase) >= 0
                                || autoId.IndexOf("Appointment_ScheduleMeeting", StringComparison.OrdinalIgnoreCase) >= 0
                                || bName.Equals("Teams Meeting", StringComparison.OrdinalIgnoreCase)
                                || bName.Equals("Schedule a meeting", StringComparison.OrdinalIgnoreCase)
                                || (bName.IndexOf("Teams", StringComparison.OrdinalIgnoreCase) >= 0
                                    && bName.IndexOf("Join", StringComparison.OrdinalIgnoreCase) < 0
                                    && bName.IndexOf("Don't", StringComparison.OrdinalIgnoreCase) < 0
                                    && bName.IndexOf("Remove", StringComparison.OrdinalIgnoreCase) < 0
                                    && bName.IndexOf("Options", StringComparison.OrdinalIgnoreCase) < 0);
                            if (!isTeamsBtn) continue;
                            if (btn.TryGetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern, out var pat)
                                && pat is System.Windows.Automation.InvokePattern inv)
                            {
                                inv.Invoke();
                                invoked = true;
                                return;
                            }
                        }
                    }
                }
                catch { }
            });
            return invoked || true; // scheduled once off-thread so UI never stalls
        }

        // Phase 3b primary scope resolver. Runs on the UI thread via the
        // marshaller in SearchMessages. Builds a SearchScope (DASL scope
        // string + SearchSubFolders flag) for Application.AdvancedSearch.
        // Yields the message pump (Application.DoEvents) between each
        // store so Outlook's UI stays responsive while we walk many
        // stores synchronously on the UI thread (35 stores observed in
        // one real trace; that alone froze the UI for ~10 seconds).
        private SearchScope BuildSearchScope(SearchMessagesArgs args, string scopeMode, CancellationToken ct)
        {
            var started = DateTimeOffset.UtcNow;
            var paths = new List<string>();
            bool searchSubFolders = true;

            try
            {
                if (!string.IsNullOrEmpty(args.FolderId))
                {
                    var folder = ResolveFolder(args.FolderId);
                    if (folder != null)
                    {
                        try { paths.Add(folder.FolderPath); } catch (COMException) { }
                    }
                }
                else if (scopeMode == "current_folder")
                {
                    var folder = ResolveCurrentFolder();
                    if (folder != null)
                    {
                        try { paths.Add(folder.FolderPath); } catch (COMException) { }
                    }
                    searchSubFolders = false;
                }
                else
                {
                    foreach (Outlook.Store store in _application.Session.Stores)
                    {
                        try
                        {
                            var root = store.GetRootFolder();
                            if (root != null)
                            {
                                try { paths.Add(root.FolderPath); } catch (COMException) { }
                            }
                        }
                        catch (COMException) { }
                        YieldUi(ct);
                    }
                }
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("BuildSearchScope COMException: " + ex.Message, "LiveOutlookSurface"); } catch { }
            }

            var elapsedMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;
            try { JameelCopilot.Diagnostics.TraceLog.Write("BuildSearchScope elapsed_ms=" + elapsedMs + " paths=" + paths.Count, "LiveOutlookSurface"); } catch { }

            return new SearchScope
            {
                ScopeString = SearchScopeFormatter.Format(paths),
                SearchSubFolders = searchSubFolders,
                ResolvedFolderPaths = paths,
            };
        }

        // Pumps Outlook's UI message queue briefly so the user's clicks,
        // hover, scroll and ribbon stay responsive while we hold the UI
        // thread to enumerate Outlook COM collections. We're already
        // executing on the Outlook UI thread (called from inside
        // marshaller.RunAsync); DoEvents is safe here because our caller
        // does only read-only enumeration. Cancellation is checked twice
        // so a user Stop click that fires during the pump still aborts.
        private static void YieldUi(CancellationToken ct)
        {
            if (ct.IsCancellationRequested) ct.ThrowIfCancellationRequested();
            try { System.Windows.Forms.Application.DoEvents(); } catch { }
            if (ct.IsCancellationRequested) ct.ThrowIfCancellationRequested();
        }

        // Phase 3b yielding fallback. Walks each folder via one
        // marshaller.RunAsync call so the UI thread is released between
        // folders (Outlook can pump messages between marshalled calls).
        private SearchResult FallbackIterativeSearch(
            SearchMessagesArgs args, string filter, string scopeMode, CancellationToken ct)
        {
            var allInputs = new List<MessageProjectionInput>();
            var earlyStop = false;
            var searchAllMail = scopeMode == "all_mail" || scopeMode == "auto";

            IReadOnlyList<Outlook.MAPIFolder> folders;
            try
            {
                folders = _marshaller.RunAsync(
                    () => ResolveSearchFolders(args, allMail: searchAllMail, ct),
                    ct).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            var searched = 0;
            foreach (var folder in folders)
            {
                ct.ThrowIfCancellationRequested();
                // Skip synchronous unindexed full-body scans across large/secondary folders when
                // query/body_contains is used: only run unindexed body GetTable on Inbox, Sent Items,
                // or the active folder so a 44-folder body scan never stalls Exchange RPC for 15-25s!
                bool hasBodyFilter = !string.IsNullOrEmpty(args.Query) || !string.IsNullOrEmpty(args.BodyContains);
                string effectiveFilter = filter;
                if (hasBodyFilter && searched >= 3 && searchAllMail && string.IsNullOrEmpty(args.FolderId))
                {
                    var fname = SafeFolderName(folder);
                    bool isPrimary = string.Equals(fname, "Inbox", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(fname, "Sent Items", StringComparison.OrdinalIgnoreCase);
                    if (!isPrimary)
                    {
                        // On secondary folders, search indexed Subject/From/To metadata instead of unindexed Body blobs.
                        var metaArgs = new SearchMessagesArgs
                        {
                            SubjectContains = !string.IsNullOrEmpty(args.SubjectContains) ? args.SubjectContains : args.Query,
                            From = args.From,
                            To = args.To,
                            DateFrom = args.DateFrom,
                            DateTo = args.DateTo,
                            AttachmentFilter = args.AttachmentFilter,
                            ReadStatus = args.ReadStatus,
                            FlagStatus = args.FlagStatus,
                            ImportanceFilter = args.ImportanceFilter,
                            SortOrder = args.SortOrder,
                            MaxResults = args.MaxResults,
                        };
                        effectiveFilter = BuildRestrictFilter(metaArgs);
                    }
                }

                List<MessageProjectionInput> folderInputs;
                try
                {
                    folderInputs = _marshaller.RunAsync(
                        () => CollectFolderInputs(folder, args, effectiveFilter, ct),
                        ct).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }
                allInputs.AddRange(folderInputs);
                searched++;
                try
                {
                    JameelCopilot.Diagnostics.TraceLog.Write(
                        "SearchMessages fallback folder_done=" + SafeFolderName(folder)
                        + " taken=" + folderInputs.Count + " searched=" + searched,
                        "LiveOutlookSurface");
                }
                catch { }

                if (SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, scopeMode, allInputs.Count))
                {
                    try
                    {
                        JameelCopilot.Diagnostics.TraceLog.Write(
                            "SearchMessages fallback recipient_all_mail_early_stop candidates=" + allInputs.Count
                            + " searched=" + searched
                            + " max_results=" + args.MaxResults
                            + " recipient=" + args.To,
                            "LiveOutlookSurface");
                    }
                    catch { }
                    earlyStop = true;
                    break;
                }

                if (SearchFallbackBudget.ShouldStopBroadAllMailScan(args, scopeMode, allInputs.Count))
                {
                    try
                    {
                        JameelCopilot.Diagnostics.TraceLog.Write(
                            "SearchMessages fallback broad_all_mail_early_stop cap=" + args.MaxResults
                            + " candidates=" + allInputs.Count
                            + " searched=" + searched
                            + " max_results=" + args.MaxResults,
                            "LiveOutlookSurface");
                    }
                    catch { }
                    earlyStop = true;
                    break;
                }
            }

            var projected = _marshaller.RunAsync(
                () => SearchResultProjector.Project(allInputs, args, _classifier, () => YieldUi(ct)),
                ct).GetAwaiter().GetResult();
            // Early-stop means we abandoned the scan before walking all folders,
            // so the page is necessarily incomplete even if the projector did
            // not clamp it. Force Truncated so the model never treats a
            // partial scan as the complete result set.
            if (earlyStop) projected.Truncated = true;
            return projected;
        }

        // Collect one folder's worth of MessageProjectionInput using Outlook's
        // Table API. Table.GetTable() returns a server-side index-backed view
        // that supports bulk column reads, Restrict, and Sort with NO
        // per-item COM round trips. On a 200-folder smoke that previously
        // took 100s+ via folder.Items iteration, this brings each folder's
        // collection cost from ~300-4500 ms down to ~10-50 ms.
        //
        // Body is still deferred via SnippetFactory and resolved on demand
        // (GetItemFromID + read Body) only for items the projector keeps
        // in its top-N.
        private List<MessageProjectionInput> CollectFolderInputs(
            Outlook.MAPIFolder folder, SearchMessagesArgs args, string filter, CancellationToken ct)
        {
            var inputs = new List<MessageProjectionInput>();
            if (folder == null) return inputs;

            string folderName = "";
            bool folderIsMail = true;
            try { folderName = folder.Name ?? ""; } catch (COMException) { }
            try { folderIsMail = folder.DefaultItemType == Outlook.OlItemType.olMailItem; } catch (COMException) { }
            if (_classifier.IsSystemFolder(folderName, folderIsMail)) return inputs;

            Outlook.Table table;
            try
            {
                table = folder.GetTable(filter ?? "", Outlook.OlTableContents.olUserItems);
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("CollectFolderInputs GetTable COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
                return inputs;
            }

            try
            {
                table.Columns.RemoveAll();
                table.Columns.Add("EntryID");
                table.Columns.Add("Subject");
                table.Columns.Add("SenderName");
                table.Columns.Add("SenderEmailAddress");
                table.Columns.Add("To");
                table.Columns.Add("ReceivedTime");
                table.Columns.Add("MessageClass");
                // Use DASL for hasattachment so we get a clean bool. The
                // friendly name "HasAttachments" works on most builds but
                // is not guaranteed across Outlook versions.
                try { table.Columns.Add("urn:schemas:httpmail:hasattachment"); } catch { }
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("CollectFolderInputs Columns COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
            }

            // Server-side sort using Outlook's index.
            try
            {
                var descending = SearchFallbackBudget.DescendingForSortOrder(args.SortOrder);
                table.Sort("ReceivedTime",
                    descending ? Outlook.OlSortOrder.olDescending : Outlook.OlSortOrder.olAscending);
            }
            catch (COMException) { /* best-effort; some stores reject Sort */ }

            var limit = SearchFallbackBudget.PerFolderItems(args);
            int taken = 0;
            int rowsScanned = 0;
            // Scan a small buffer past `limit` so we can skip non-mail rows
            // (meeting reqs, calendar items pinned to mail folders, etc.)
            // without falling short. With taken-counting only mail messages
            // this buffer is rarely needed but cheap insurance.
            int rowScanCap = Math.Max(limit * 3, 10);

            try
            {
                while (!table.EndOfTable && taken < limit && rowsScanned < rowScanCap)
                {
                    ct.ThrowIfCancellationRequested();
                    rowsScanned++;
                    Outlook.Row row;
                    try { row = table.GetNextRow(); }
                    catch (COMException) { break; }
                    if (row == null) break;

                    var input = TryBuildFallbackInputFromRow(row, folderName, folderIsMail);
                    if (input != null) { inputs.Add(input); taken++; }
                }
            }
            catch (COMException ex)
            {
                try { JameelCopilot.Diagnostics.TraceLog.Write("CollectFolderInputs row scan COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
            }

            return inputs;
        }

        // Reads one Outlook.Table row into a MessageProjectionInput. Returns
        // null for non-mail rows or rows that fail to project. EntryID is
        // captured for the snippet factory so we only pay Body cost for
        // items the projector keeps as top-N survivors.
        private MessageProjectionInput TryBuildFallbackInputFromRow(
            Outlook.Row row, string folderName, bool folderIsMail)
        {
            string messageClass = "";
            try { messageClass = row["MessageClass"] as string ?? ""; } catch (COMException) { }
            if (!TableMessageClassFilter.IsMailMessage(messageClass)) return null;

            string entryId = "";
            try { entryId = row["EntryID"] as string ?? ""; } catch (COMException) { }
            string id = "";
            try { if (!string.IsNullOrEmpty(entryId)) id = _ids.Shorten(entryId); } catch (COMException) { }

            string subject = "";
            try { subject = row["Subject"] as string ?? ""; } catch (COMException) { }

            string senderName = "";
            try { senderName = row["SenderName"] as string ?? ""; } catch (COMException) { }
            string senderEmail = "";
            try { senderEmail = row["SenderEmailAddress"] as string ?? ""; } catch (COMException) { }
            string from = !string.IsNullOrEmpty(senderName) ? senderName : (senderEmail ?? "");

            string to = "";
            try { to = row["To"] as string ?? ""; } catch (COMException) { }

            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try
            {
                var rt = row["ReceivedTime"];
                if (rt is DateTime dt) receivedAt = ToOffset(dt);
            }
            catch (COMException) { }

            bool hasAttachments = false;
            try
            {
                var ha = row["urn:schemas:httpmail:hasattachment"];
                if (ha is bool b) hasAttachments = b;
            }
            catch (COMException) { }
            catch (System.Runtime.InteropServices.InvalidComObjectException) { }

            var capturedEntryId = entryId;
            var application = _application;
            Func<string> snippetFactory = () =>
            {
                if (string.IsNullOrEmpty(capturedEntryId)) return "";
                try
                {
                    var item = application.Session.GetItemFromID(capturedEntryId) as Outlook.MailItem;
                    if (item == null) return "";
                    return SnippetOf(item.Body);
                }
                catch (COMException) { return ""; }
                catch (System.Collections.Generic.KeyNotFoundException) { return ""; }
            };

            return new MessageProjectionInput
            {
                Id = id,
                Subject = subject,
                From = from,
                To = SplitAddresses(to),
                ReceivedAt = receivedAt,
                HasAttachments = hasAttachments,
                FolderName = folderName,
                FolderDefaultItemTypeIsMail = folderIsMail,
                SnippetFactory = snippetFactory,
            };
        }

        // Legacy MailItem-based projection. Kept for any future code path
        // that already has a MailItem in hand and wants to avoid the round
        // trip back through GetItemFromID. Not used by CollectFolderInputs
        // any more.
        private MessageProjectionInput BuildFallbackInput(
            Outlook.MailItem mi, string folderName, bool folderIsMail)
        {
            string id = ""; try { id = _ids.Shorten(mi.EntryID); } catch (COMException) { }
            string subject = ""; try { subject = mi.Subject ?? ""; } catch (COMException) { }
            string from = ""; try { from = mi.SenderName ?? mi.SenderEmailAddress ?? ""; } catch (COMException) { }
            string to = ""; try { to = mi.To ?? ""; } catch (COMException) { }
            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try { receivedAt = ToOffset(mi.ReceivedTime); } catch (COMException) { }
            bool hasAttachments = false;
            try { hasAttachments = mi.Attachments != null && mi.Attachments.Count > 0; } catch (COMException) { }

            var capturedMi = mi;
            Func<string> snippetFactory = () =>
            {
                try { return SnippetOf(capturedMi.Body); } catch (COMException) { return ""; }
            };

            return new MessageProjectionInput
            {
                Id = id,
                Subject = subject,
                From = from,
                To = SplitAddresses(to),
                ReceivedAt = receivedAt,
                HasAttachments = hasAttachments,
                FolderName = folderName,
                FolderDefaultItemTypeIsMail = folderIsMail,
                SnippetFactory = snippetFactory,
            };
        }

        private Outlook.MAPIFolder ResolveCurrentFolder()
        {
            try
            {
                var folder = GetLiveExplorer()?.CurrentFolder as Outlook.MAPIFolder;
                if (folder != null) return folder;
            }
            catch (COMException) { }

            try { return _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); }
            catch (COMException) { return null; }
        }

        private IReadOnlyList<Outlook.MAPIFolder> ResolveSearchFolders(SearchMessagesArgs args, bool allMail, CancellationToken ct)
        {
            var started = DateTimeOffset.UtcNow;
            var folders = new List<Outlook.MAPIFolder>();
            var folderLimit = SearchFallbackBudget.MaxFoldersForSearch(args, allMail);
            if (!string.IsNullOrEmpty(args?.FolderId))
            {
                var folder = ResolveFolder(args.FolderId);
                if (folder != null) folders.Add(folder);
                return folders;
            }

            if (!allMail)
            {
                var folder = ResolveCurrentFolder();
                if (folder != null) folders.Add(folder);
                return folders;
            }

            if (folderLimit < SearchFallbackBudget.MaxSearchFolders)
            {
                try
                {
                    JameelCopilot.Diagnostics.TraceLog.Write(
                        "ResolveSearchFolders broad_all_mail_folder_cap cap=" + folderLimit
                        + " max_results=" + args.MaxResults,
                        "LiveOutlookSurface");
                }
                catch { }
            }

            try
            {
                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    WalkMailFolders(store.GetRootFolder(), folders, depth: 0, folderLimit, ct);
                    // Pump Outlook UI between each store so the mailbox
                    // tree walk does not freeze Outlook for many seconds
                    // when the user has many stores.
                    YieldUi(ct);
                    if (folders.Count >= folderLimit) break;
                }
            }
            catch (COMException) { }
            var elapsedMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;
            try { JameelCopilot.Diagnostics.TraceLog.Write("ResolveSearchFolders elapsed_ms=" + elapsedMs + " folders=" + folders.Count + " cap=" + folderLimit, "LiveOutlookSurface"); } catch { }
            return folders;
        }

        private void WalkMailFolders(Outlook.MAPIFolder folder, List<Outlook.MAPIFolder> results, int depth, int folderLimit, CancellationToken ct)
        {
            if (folder == null || depth > MaxFolderDepth) return;
            ct.ThrowIfCancellationRequested();
            var name = SafeFolderName(folder);
            bool isMailFolder = false;
            try { isMailFolder = folder.DefaultItemType == Outlook.OlItemType.olMailItem; }
            catch (COMException) { }
            if (!_classifier.IsSystemFolder(name, isMailFolder))
            {
                if (isMailFolder) results.Add(folder);
            }

            if (results.Count >= folderLimit) return;

            try
            {
                int childIndex = 0;
                foreach (Outlook.MAPIFolder child in folder.Folders)
                {
                    WalkMailFolders(child, results, depth + 1, folderLimit, ct);
                    // Yield every few children to keep the UI responsive
                    // during deep folder trees.
                    if ((++childIndex % 5) == 0) YieldUi(ct);
                    if (results.Count >= folderLimit) break;
                }
            }
            catch (COMException) { }
        }

        private static string SafeFolderName(Outlook.MAPIFolder folder)
        {
            try { return folder?.Name ?? ""; }
            catch (COMException) { return ""; }
        }

        private Outlook.MAPIFolder ResolveFolder(string folderId)
        {
            if (string.IsNullOrEmpty(folderId))
            {
                try { return _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); }
                catch (COMException) { return null; }
            }
            try
            {
                var entryId = _ids.Resolve(folderId);
                return _application.Session.GetFolderFromID(entryId) as Outlook.MAPIFolder;
            }
            catch (COMException) { return null; }
            catch (KeyNotFoundException) { return null; }
        }

        private void WalkFolders(Outlook.MAPIFolder folder, string parentId, List<FolderResult> results, int depth)
        {
            if (folder == null) return;
            if (depth > MaxFolderDepth) return;
            if (results.Count >= SearchFallbackBudget.MaxListFolders) return;

            string id;
            try { id = _ids.Shorten(folder.EntryID); }
            catch (COMException) { return; }

            int count = 0;
            try { count = folder.Items.Count; } catch (COMException) { }

            results.Add(new FolderResult
            {
                Id = id,
                Name = folder.Name ?? "",
                ParentId = parentId,
                ItemCount = count,
            });

            try
            {
                foreach (Outlook.MAPIFolder child in folder.Folders)
                {
                    WalkFolders(child, id, results, depth + 1);
                    if (results.Count >= SearchFallbackBudget.MaxListFolders) break;
                }
            }
            catch (COMException) { }
        }

        // DASL property name constants. Property names in @SQL= filter
        // expressions MUST be wrapped in double quotes for folder.GetTable()
        // (which is what Phase 3b's iterative fallback uses); Items.Restrict
        // is forgiving and accepts unquoted names. Without the quotes,
        // GetTable silently returns 0 rows. Microsoft DASL spec:
        // https://learn.microsoft.com/en-us/office/vba/outlook/concepts/forms/refer-to-properties-by-namespace
        private const string PropSubject     = "\"urn:schemas:httpmail:subject\"";
        private const string PropBody        = "\"urn:schemas:httpmail:textdescription\"";
        // urn:schemas:httpmail:from is the RFC 822 From header value as
        // Outlook resolves it - typically "Display Name" <smtp@domain> or
        // just smtp@domain. This is the property the model means when it
        // says "from X". 'fromname' (which we used previously) is NOT a
        // real DASL URN; folder.GetTable's strict parser silently treated
        // it as FALSE, killing every from= search.
        private const string PropFrom        = "\"urn:schemas:httpmail:from\"";
        private const string PropFromEmail   = "\"urn:schemas:httpmail:fromemail\"";
        // PR_SENDER_SMTP_ADDRESS (0x5D01) PT_UNICODE_STRING (0x001F) ->
        // always-SMTP form of the sender, when set. Not populated for many
        // Exchange-routed messages. Backstop only.
        private const string PropFromSmtp    = "\"http://schemas.microsoft.com/mapi/proptag/0x5D01001F\"";
        private const string PropDisplayTo   = "\"urn:schemas:httpmail:displayto\"";
        // NOTE: we deliberately do NOT include PR_TRANSPORT_MESSAGE_HEADERS
        // (proptag 0x007D001F) in any routine filter. It is a multi-KB raw
        // RFC 822 header blob, server-NOT-indexed, and LIKE'ing against it
        // forces a full per-message read across every folder. On a large
        // mailbox this freezes Outlook for 10+ minutes and trips the
        // "trouble connecting to server" cascade. The 'from' URN above
        // already contains the resolved SMTP segment for typical received
        // mail, so the headers blob is redundant on the hot path.
        private const string PropHasAttach   = "\"urn:schemas:httpmail:hasattachment\"";
        private const string PropRead        = "\"urn:schemas:httpmail:read\"";
        private const string PropReceivedAt  = "\"urn:schemas:httpmail:datereceived\"";
        private const string PropFlagStatus  = "\"http://schemas.microsoft.com/mapi/proptag/0x10900003\"";
        private const string PropImportance  = "\"http://schemas.microsoft.com/mapi/proptag/0x00170003\"";

        internal static string BuildRestrictFilter(SearchMessagesArgs args)
        {
            if (args == null) return null;
            var clauses = new List<string>();

            if (!string.IsNullOrEmpty(args.Query))
            {
                var q = Escape(args.Query);
                clauses.Add("(" + PropSubject + " LIKE '%" + q + "%' OR " +
                            PropBody + " LIKE '%" + q + "%')");
            }
            if (!string.IsNullOrEmpty(args.From))
            {
                var v = Escape(args.From);
                clauses.Add("(" + PropFrom + " LIKE '%" + v + "%' OR " +
                            PropFromEmail + " LIKE '%" + v + "%' OR " +
                            PropFromSmtp + " LIKE '%" + v + "%')");
            }
            if (!string.IsNullOrEmpty(args.To))
            {
                clauses.Add(PropDisplayTo + " LIKE '%" + Escape(args.To) + "%'");
            }
            if (!string.IsNullOrEmpty(args.SubjectContains))
            {
                clauses.Add(PropSubject + " LIKE '%" + Escape(args.SubjectContains) + "%'");
            }
            if (!string.IsNullOrEmpty(args.BodyContains))
            {
                clauses.Add(PropBody + " LIKE '%" + Escape(args.BodyContains) + "%'");
            }
            var attachmentFilter = (args.AttachmentFilter ?? "any").Trim().ToLowerInvariant();
            if (attachmentFilter == "with" || args.HasAttachment == true)
            {
                clauses.Add(PropHasAttach + " = 1");
            }
            else if (attachmentFilter == "without")
            {
                clauses.Add(PropHasAttach + " = 0");
            }

            var readStatus = (args.ReadStatus ?? "any").Trim().ToLowerInvariant();
            if (readStatus == "unread" || args.IsUnread == true)
            {
                clauses.Add(PropRead + " = 0");
            }
            else if (readStatus == "read")
            {
                clauses.Add(PropRead + " = 1");
            }

            var flagStatus = (args.FlagStatus ?? "any").Trim().ToLowerInvariant();
            if (flagStatus == "flagged" || args.IsFlagged == true)
            {
                clauses.Add(PropFlagStatus + " = 2");
            }
            else if (flagStatus == "unflagged")
            {
                clauses.Add(PropFlagStatus + " <> 2");
            }

            var importanceFilter = (args.ImportanceFilter ?? "any").Trim().ToLowerInvariant();
            if (importanceFilter == "any" && !string.IsNullOrEmpty(args.Importance))
            {
                importanceFilter = args.Importance.Trim().ToLowerInvariant();
            }
            if (importanceFilter == "low" || importanceFilter == "normal" || importanceFilter == "high")
            {
                // PR_IMPORTANCE (0x0017) PT_LONG => 0x00170003. 0=low, 1=normal, 2=high.
                var val = importanceFilter == "low" ? "0" : importanceFilter == "normal" ? "1" : "2";
                clauses.Add(PropImportance + " = " + val);
            }
            if (args.DateFrom.HasValue)
            {
                // Outlook Table DASL ignores ISO-8601 T literals; US date
                // literals are live-verified with folder.GetTable.
                clauses.Add(PropReceivedAt + " >= '" +
                    args.DateFrom.Value.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "'");
            }
            if (args.DateTo.HasValue)
            {
                clauses.Add(PropReceivedAt + " <= '" +
                    args.DateTo.Value.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture) + "'");
            }

            if (clauses.Count == 0) return null;
            return "@SQL=" + string.Join(" AND ", clauses);
        }

        private static string Escape(string s) => (s ?? "").Replace("'", "''");
    }
}
