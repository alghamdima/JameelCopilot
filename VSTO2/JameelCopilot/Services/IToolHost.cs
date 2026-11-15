using System.Threading;
using System.Threading.Tasks;

namespace JameelCopilot.Services
{
    /// <summary>
    /// Seam between CustomApiChatService and Outlook tool execution. Production
    /// implementation is OutlookToolHost (Task 20). Tests use FakeToolHost.
    /// </summary>
    public interface IToolHost
    {
        Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct);
    }
}
