using System;
using System.Runtime.InteropServices;
using JameelCopilot.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace JameelCopilot.Services.Tools
{
    /// <summary>
    /// The compose window through its Word editor. Outlook bookmarks the
    /// automatic signature and the start of the quoted message, so this is
    /// where the user's own text can be told from both, and changed without
    /// touching either: the signature keeps its picture and formatting, the
    /// thread its layout. Call on the UI thread.
    /// </summary>
    internal static class ComposeDocument
    {
        private const string SignatureMark = "_MailAutoSig";
        private const string QuotedMark = "_MailOriginal";

        public static ComposeBody Read(Outlook.Inspector inspector, Outlook.MailItem item)
        {
            try
            {
                var doc = Editor(inspector);
                if (doc != null)
                {
                    Split(doc, out int ownEnd, out int quotedStart);
                    TraceLog.Write("Read: own text ends at " + ownEnd + ", quoted message starts at " + quotedStart +
                                   " of " + doc.Content.End, "ComposeDocument");
                    return new ComposeBody(
                        ComposeBody.StripSignatureAndFooter(ComposeBody.FromWord(doc.Range(0, ownEnd).Text)),
                        ComposeBody.StripThreadSignatures(ComposeBody.FromWord(doc.Range(quotedStart, doc.Content.End).Text)));
                }
            }
            catch (Exception ex) when (ex is COMException || ex is InvalidCastException)
            {
                TraceLog.Write("Read: editor unavailable (" + ex.GetType().Name + "), using the plain body", "ComposeDocument");
            }

            // shortcut: plain text has no mark for the signature, so it stays in Own; the editor path has none of this.
            string body = "";
            try { body = item?.Body ?? ""; } catch (COMException) { }
            return ComposeBody.FromPlainText(body);
        }

        /// <summary>Puts the text above everything the window holds. False when the window has no editor.</summary>
        public static bool InsertAtTop(Outlook.Inspector inspector, string text)
        {
            var doc = Editor(inspector);
            if (doc == null) return false;
            doc.Range(0, 0).InsertBefore(ComposeBody.ToWord(text) + "\r");
            return true;
        }

        /// <summary>Replaces what the user wrote; the signature and the quoted message stay as they are.</summary>
        public static bool ReplaceOwnText(Outlook.Inspector inspector, string text)
        {
            var doc = Editor(inspector);
            if (doc == null) return false;
            Split(doc, out int ownEnd, out _);
            int end = doc.Content.End;
            // A signature or quoted message follows: the new text ends its own
            // paragraph. Otherwise the document's last paragraph mark stays.
            bool followed = ownEnd < end;
            doc.Range(0, followed ? ownEnd : end - 1).Text = ComposeBody.ToWord(text) + (followed ? "\r" : "");
            return true;
        }

        /// <summary>
        /// The text the user selected in their own part of the email (not the
        /// signature, not the quoted message), or null.
        /// </summary>
        public static ComposeSelection ReadSelection(Outlook.Inspector inspector)
        {
            var doc = Editor(inspector);
            if (doc == null) return null;
            var selection = doc.ActiveWindow?.Selection;
            if (selection == null || selection.End <= selection.Start) return null;
            Split(doc, out int ownEnd, out _);
            if (selection.Start < 0 || selection.End > ownEnd) return null;
            var text = ComposeBody.FromWord(doc.Range(selection.Start, selection.End).Text);
            return text.Length == 0 ? null : new ComposeSelection(selection.Start, selection.End, text);
        }

        /// <summary>
        /// Replaces the selected text, if it is still what it was when Jameel Copilot read
        /// it. False when the user changed it since, or there is no editor.
        /// </summary>
        public static bool ReplaceSelection(Outlook.Inspector inspector, ComposeSelection selection, string text)
        {
            var doc = Editor(inspector);
            if (doc == null || selection == null || selection.End > doc.Content.End) return false;
            var range = doc.Range(selection.Start, selection.End);
            if (ComposeBody.FromWord(range.Text) != selection.Text) return false;
            range.Text = ComposeBody.ToWord(text);
            return true;
        }

        private static Word.Document Editor(Outlook.Inspector inspector)
        {
            return inspector?.WordEditor as Word.Document;
        }

        // Where the user's own text ends and where the quoted message starts.
        private static void Split(Word.Document doc, out int ownEnd, out int quotedStart)
        {
            int end = doc.Content.End;
            quotedStart = StartOf(doc, QuotedMark, end);
            if (quotedStart == end)
            {
                // No bookmark (a draft saved by another client): find the header in the
                // text, and trust the place only if the text before it is that long in Word too.
                string text = doc.Content.Text ?? "";
                int at = ComposeBody.QuotedIndex(text.Replace('\r', '\n'));
                if (at >= 0 && (doc.Range(0, at).Text ?? "") == text.Substring(0, at)) quotedStart = at;
            }
            int sigStart = StartOf(doc, SignatureMark, end);
            ownEnd = Math.Min(sigStart, quotedStart);
            if (sigStart == end && ownEnd > 0)
            {
                // Manual signature without _MailAutoSig bookmark: find where the clean text ends
                // so ReplaceOwnText preserves the manual signature block below it.
                try
                {
                    string rawOwn = doc.Range(0, ownEnd).Text ?? "";
                    string cleanOwn = ComposeBody.StripSignatureAndFooter(ComposeBody.FromWord(rawOwn));
                    if (cleanOwn.Length == 0 && rawOwn.Trim().Length > 0)
                    {
                        // Nothing typed yet, only a manual signature sits above the quoted message.
                        ownEnd = 0;
                    }
                    else if (cleanOwn.Length > 0)
                    {
                        string[] cleanLines = cleanOwn.Split('\n');
                        string lastLine = cleanLines[cleanLines.Length - 1].Trim();
                        if (lastLine.Length > 0)
                        {
                            int idx = rawOwn.IndexOf(lastLine, StringComparison.Ordinal);
                            if (idx >= 0)
                            {
                                int candidateEnd = idx + lastLine.Length;
                                if (candidateEnd < ownEnd && (doc.Range(0, candidateEnd).Text ?? "") == rawOwn.Substring(0, candidateEnd))
                                {
                                    ownEnd = candidateEnd;
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private static int StartOf(Word.Document doc, string mark, int none)
        {
            return doc.Bookmarks.Exists(mark) ? doc.Bookmarks[mark].Start : none;
        }
    }

    /// <summary>A stretch of the user's own text, by its place in the editor and what it said.</summary>
    internal sealed class ComposeSelection
    {
        public ComposeSelection(int start, int end, string text)
        {
            Start = start;
            End = end;
            Text = text ?? "";
        }

        public int Start { get; }
        public int End { get; }
        public string Text { get; }
    }
}
