namespace JameelCopilot.Services.Chat
{
    /// <summary>
    /// Why <see cref="JameelCopilot.Services.CustomApiChatService"/> stopped a turn.
    /// </summary>
    public enum StopReason
    {
        Completed,
        Cancelled,
        MaxRoundsReached,
        Error,
    }
}
