using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using JameelCopilot.Diagnostics;
using JameelCopilot.Services;
using JameelCopilot.Services.Chat;
using JameelCopilot.Services.Prefs;
using JameelCopilot.Services.Tools;
using JameelCopilot.TaskPane.Chat;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JameelCopilot.TaskPane
{
    /// <summary>
    /// The Jameel Copilot pane of one compose window. It only hosts the chat page: writing
    /// and editing the email, the chat about its thread, saved chats and
    /// memory are all in that page (<see cref="ChatController"/>).
    /// </summary>
    public partial class AITaskPane : UserControl
    {
        private ChatController _chatController;

        public AITaskPane()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Binds the pane to its compose window. Called by
        /// <see cref="ThisAddIn.ShowTaskPane"/> right after construction.
        /// </summary>
        public void Bind(Outlook.Inspector inspector)
        {
            using (TraceLog.Scope("Bind", "AITaskPane"))
            {
                try
                {
                    var addIn = Globals.ThisAddIn;
                    if (addIn?.ChatService == null || addIn.OutlookMarshaller == null || addIn.IdResolver == null
                        || addIn.Application == null || addIn.AdvancedSearchRunner == null)
                    {
                        TraceLog.Write("Bind: services missing; no chat", "AITaskPane");
                        return;
                    }

                    var surface = new LiveOutlookSurface(addIn.Application, addIn.OutlookMarshaller, addIn.IdResolver,
                        composeInspector: inspector, explorer: null,
                        runner: addIn.AdvancedSearchRunner, classifier: addIn.FolderClassifier);
                    var toolHost = new OutlookToolHost(surface, Config.WriteToolsEnabled);
                    ThreadOf(inspector, out var scope, out var title);
                    _chatController = new ChatController(
                        chatHost, addIn.ChatService, toolHost, surface, new ConversationStore(), addIn.UserPrefs,
                        addIn.Chats, addIn.Memory, scope, title, inspector, addIn.Prompts);

                    // Not awaited: the controller shows its own message if WebView2 cannot start.
                    _chatController.InitializeAsync().ContinueWith(t =>
                    {
                        if (t.IsFaulted) TraceLog.Write("ChatController.InitializeAsync faulted: " + t.Exception?.GetBaseException().GetType().Name, "AITaskPane");
                    }, TaskScheduler.Default);
                }
                catch (Exception ex)
                {
                    TraceLog.Write("Bind failed: " + ex.GetType().Name, "AITaskPane");
                }
            }
        }

        // Which saved chat belongs to this window: a reply or forward shares its
        // thread's chat; a new email starts one of its own. The thread is known
        // by the store's conversation id only. The subject would do as a key,
        // but whoever sends a mail chooses its subject: a mail named like
        // another thread would get that thread's chat when the user replies.
        private static void ThreadOf(Outlook.Inspector inspector, out string scope, out string title)
        {
            scope = ChatsBridge.ComposeScope;
            title = null;
            try
            {
                var mail = inspector?.CurrentItem as Outlook.MailItem;
                if (string.IsNullOrWhiteSpace(mail?.ConversationTopic)) return;   // a new email
                title = mail.Subject;
                var id = mail.ConversationID;
                if (!string.IsNullOrWhiteSpace(id)) scope = ChatArchive.ThreadScope(id);
            }
            catch (COMException ex)
            {
                TraceLog.Write("ThreadOf failed: " + ex.GetType().Name, "AITaskPane");
            }
        }

        partial void DisposeCustomResources()
        {
            try { _chatController?.Dispose(); } catch { }
        }
    }
}
