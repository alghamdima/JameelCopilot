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
            CurrentSelectionResult selection)
        {
            var sb = new StringBuilder();
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
    }
}
