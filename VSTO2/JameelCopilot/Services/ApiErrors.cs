using System;
using System.Net;

namespace JameelCopilot.Services
{
    /// <summary>
    /// What the user is told when the AI service can't answer. The service's
    /// own error text is never shown: it may repeat mail content or internals.
    /// </summary>
    internal static class ApiErrors
    {
        public const string NotConfigured = "Jameel Copilot is not configured. Contact IT.";
        public const string SignIn = "Jameel Copilot could not sign in to the AI service. Contact IT.";
        public const string CannotReach = "Jameel Copilot cannot reach the AI service.";
        public const string ErrorInAnswer = "The AI service returned an error.";
        public const string EmptyAnswer = "The AI service returned an empty answer.";
        public const string CutAtLimit = "[The answer reached the length limit and stops here.]";

        public static string ForStatus(HttpStatusCode status, TimeSpan? retryAfter)
        {
            var code = (int)status;
            if (code == 401 || code == 403) return SignIn;
            if (code == 429)
            {
                return retryAfter.HasValue && retryAfter.Value > TimeSpan.Zero
                    ? "The AI service is busy. Try again in " + (int)Math.Ceiling(retryAfter.Value.TotalSeconds) + " seconds."
                    : "The AI service is busy. Try again shortly.";
            }
            return "The AI service returned an error (HTTP " + code + ").";
        }

        /// <summary>A request that never got an answer: name resolution, TLS, connection or timeout.</summary>
        public static string ForException(Exception ex)
        {
            return CannotReach;
        }
    }
}
