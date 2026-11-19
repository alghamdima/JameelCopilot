using System.Collections.Generic;
using System.Linq;
using System.Text;
using JameelCopilot.Services;
using JameelCopilot.Services.Tools;

namespace JameelCopilot.TaskPane.InboxCopilot
{
    /// <summary>
    /// What the Inbox pane tells the model about the mailbox for one turn:
    /// the folder in view and the selection, with each selected email's id so
    /// the tools can read it. Sent as context (data), after the fixed prompt.
    /// Sender-written fields stay on one line. A pure function, so it is
    /// tested without the controller.
    /// </summary>
    public static class InboxCopilotPromptBuilder
    {
        public static string BuildContext(
            string folderName,
            int unreadCount,
            int totalCount,
            CurrentSelectionResult selection,
            OrgContextInfo org = null)
        {
            var sb = new StringBuilder();
            var orgBlock = FormatOrgContext(org);
            if (!string.IsNullOrEmpty(orgBlock))
            {
                sb.AppendLine(orgBlock);
            }
            sb.Append("- Folder: ").Append(StructuredAnswers.OneLine(folderName ?? "Inbox"));
            sb.Append(" (").Append(unreadCount).Append(" unread, ").Append(totalCount).Append(" total)");
            sb.AppendLine();

            if (selection != null && selection.Count > 0)
            {
                if (selection.Count == 1 && selection.Messages != null && selection.Messages.Count > 0)
                {
                    var m = selection.Messages[0];
                    sb.Append("- Selected: ").AppendLine(StructuredAnswers.OneLine(m.Subject));
                    if (!string.IsNullOrEmpty(m.Id))
                    {
                        sb.Append("  Id (for outlook_read_message): ").AppendLine(StructuredAnswers.OneLine(m.Id));
                    }
                    if (!string.IsNullOrEmpty(m.From))
                    {
                        sb.Append("  From: ").AppendLine(StructuredAnswers.OneLine(m.From));
                    }
                    if (m.To != null && m.To.Count > 0)
                    {
                        sb.Append("  To: ").AppendLine(StructuredAnswers.OneLine(string.Join(", ", m.To.Take(12))));
                    }
                    if (m.Cc != null && m.Cc.Count > 0)
                    {
                        sb.Append("  Cc: ").AppendLine(StructuredAnswers.OneLine(string.Join(", ", m.Cc.Take(12))));
                    }
                    sb.Append("  Received: ").AppendLine(m.ReceivedAt.ToString("o"));
                    var snippet = StructuredAnswers.OneLine(m.BodyPlaintext);
                    if (snippet.Length > 200) snippet = snippet.Substring(0, 200);
                    if (!string.IsNullOrEmpty(snippet))
                    {
                        sb.Append("  Snippet: ").AppendLine(snippet);
                    }
                }
                else
                {
                    sb.Append("- ").Append(selection.Count).AppendLine(" messages selected");
                    foreach (var m in (selection.Messages ?? new MessageDetail[0]).Where(x => !string.IsNullOrEmpty(x.Id)).Take(5))
                    {
                        sb.Append("  Id ").Append(StructuredAnswers.OneLine(m.Id)).Append(": ")
                          .Append(StructuredAnswers.OneLine(m.From)).Append(" - ").AppendLine(StructuredAnswers.OneLine(m.Subject));
                    }
                }
            }

            return sb.ToString();
        }

        public static string FormatOrgContext(OrgContextInfo org)
        {
            if (org?.User == null) return "";
            var sb = new StringBuilder();
            sb.Append("- User profile: ").Append(FormatPerson(org.User));
            if (!string.IsNullOrEmpty(org.PreferredBuilding))
            {
                sb.Append(" | Primary building: ").Append(StructuredAnswers.OneLine(org.PreferredBuilding));
            }
            if (org.Manager != null && !string.IsNullOrEmpty(org.Manager.Name))
            {
                sb.AppendLine();
                sb.Append("  Manager: ").Append(FormatPerson(org.Manager));
            }
            if (org.DirectReports != null && org.DirectReports.Count > 0)
            {
                sb.AppendLine();
                sb.Append("  Direct team (").Append(org.DirectReports.Count).Append("): ")
                  .Append(string.Join("; ", org.DirectReports.Take(20).Select(FormatPerson)));
            }
            if (org.Peers != null && org.Peers.Count > 0)
            {
                sb.AppendLine();
                sb.Append("  Peers under same manager (").Append(org.Peers.Count).Append("): ")
                  .Append(string.Join("; ", org.Peers.Take(15).Select(FormatPerson)));
            }
            return sb.ToString();
        }

        public static Newtonsoft.Json.Linq.JArray BuildPeopleHints(OrgContextInfo org, IEnumerable<string> extraAddresses = null)
        {
            var list = new Newtonsoft.Json.Linq.JArray();
            var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var myEmail = org?.User?.Email ?? "";

            void AddOrgPerson(OrgPerson p, string role)
            {
                if (p == null) return;
                var name = StructuredAnswers.OneLine(p.Name ?? "");
                var email = StructuredAnswers.OneLine(p.Email ?? "");
                if (name.Length == 0 && email.Length == 0) return;
                if (email.Length > 0 && string.Equals(email, myEmail, System.StringComparison.OrdinalIgnoreCase)) return;
                var key = email.Length > 0 ? email : name;
                if (!seen.Add(key)) return;
                var meta = StructuredAnswers.OneLine(p.JobTitle ?? "");
                if (!string.IsNullOrEmpty(role))
                {
                    meta = string.IsNullOrEmpty(meta) ? role : (role + " - " + meta);
                }
                list.Add(new Newtonsoft.Json.Linq.JObject(
                    new Newtonsoft.Json.Linq.JProperty("name", name),
                    new Newtonsoft.Json.Linq.JProperty("email", email),
                    new Newtonsoft.Json.Linq.JProperty("meta", meta)));
            }

            if (extraAddresses != null)
            {
                foreach (var raw in extraAddresses)
                {
                    var clean = StructuredAnswers.OneLine(raw ?? "");
                    if (clean.Length == 0) continue;
                    string nm = clean, em = "";
                    int lt = clean.IndexOf('<'), gt = clean.IndexOf('>');
                    if (lt >= 0 && gt > lt)
                    {
                        nm = clean.Substring(0, lt).Trim().Trim('"');
                        em = clean.Substring(lt + 1, gt - lt - 1).Trim();
                    }
                    else if (clean.Contains("@"))
                    {
                        em = clean;
                        nm = clean.Substring(0, clean.IndexOf('@'));
                    }
                    if (em.Length > 0 && string.Equals(em, myEmail, System.StringComparison.OrdinalIgnoreCase)) continue;
                    var key = em.Length > 0 ? em : nm;
                    if (!seen.Add(key)) continue;
                    list.Add(new Newtonsoft.Json.Linq.JObject(
                        new Newtonsoft.Json.Linq.JProperty("name", nm),
                        new Newtonsoft.Json.Linq.JProperty("email", em),
                        new Newtonsoft.Json.Linq.JProperty("meta", "Thread")));
                }
            }

            AddOrgPerson(org?.Manager, "Manager");
            foreach (var r in org?.DirectReports ?? new OrgPerson[0]) AddOrgPerson(r, "Team");
            foreach (var p in org?.Peers ?? new OrgPerson[0]) AddOrgPerson(p, "Peer");
            return list;
        }

        private static string FormatPerson(OrgPerson p)
        {
            if (p == null) return "";
            var s = StructuredAnswers.OneLine(p.Name ?? "");
            if (!string.IsNullOrEmpty(p.Email)) s += " <" + StructuredAnswers.OneLine(p.Email) + ">";
            if (!string.IsNullOrEmpty(p.JobTitle)) s += " (" + StructuredAnswers.OneLine(p.JobTitle) + ")";
            return s;
        }
    }
}
