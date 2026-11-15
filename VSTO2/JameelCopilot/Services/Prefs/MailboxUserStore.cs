using System;
using System.Runtime.InteropServices;
using System.Threading;
using JameelCopilot.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JameelCopilot.Services.Prefs
{
    /// <summary>
    /// Keeps the user's Jameel Copilot data in hidden items in their own mailbox
    /// (Outlook's storage for add-in data, in the Inbox of the default
    /// store), so it follows them to another PC and never leaves the
    /// company's mail system. A profile that cannot hold such items (some
    /// shared or file-only profiles) simply answers null and false.
    /// </summary>
    public sealed class MailboxUserStore : IUserStore
    {
        private const string Prefix = "JameelCopilot.";

        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;

        public MailboxUserStore(Outlook.Application application, OutlookThreadMarshaller marshaller)
        {
            _application = application;
            _marshaller = marshaller;
        }

        public string Load(string name)
        {
            return OnOutlookThread("load", item => item.Size == 0 ? "" : (item.Body ?? ""), null, name);
        }

        public bool Save(string name, string json)
        {
            return OnOutlookThread("save", item =>
            {
                item.Body = json ?? "";
                item.Save();
                return true;
            }, false, name);
        }

        public bool Delete(string name)
        {
            return OnOutlookThread("delete", item =>
            {
                item.Delete();
                return true;
            }, false, name);
        }

        private T OnOutlookThread<T>(string what, Func<Outlook.StorageItem, T> use, T unavailable, string name)
        {
            try
            {
                if (_application == null || _marshaller == null) return unavailable;
                return _marshaller.RunAsync(() => WithItem(name, use), CancellationToken.None).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                // The name of the failure only: the item holds the user's own data.
                TraceLog.Write("User store " + what + " failed: " + ex.GetType().Name, "UserStore");
                return unavailable;
            }
        }

        private T WithItem<T>(string name, Func<Outlook.StorageItem, T> use)
        {
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.StorageItem item = null;
            try
            {
                session = _application.Session;
                inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                item = inbox.GetStorage(Prefix + name, Outlook.OlStorageIdentifierType.olIdentifyBySubject);
                return use(item);
            }
            finally
            {
                if (item != null) Marshal.ReleaseComObject(item);
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (session != null) Marshal.ReleaseComObject(session);
            }
        }
    }
}
