using System;
using JameelCopilot.Services;
using WixToolset.Dtf.WindowsInstaller;

namespace JameelCopilot.Installer.Actions
{
    public static class ProtectApiKeyAction
    {
        /// <summary>
        /// Turns the API_KEY given on the command line into API_KEY_PROTECTED,
        /// encrypted for this machine, and clears API_KEY, so only the
        /// encrypted form is ever written to the registry. Neither value is
        /// written to the log.
        /// </summary>
        [CustomAction]
        public static ActionResult ProtectApiKey(Session session)
        {
            try
            {
                var key = session["API_KEY"];
                if (string.IsNullOrEmpty(key)) return ActionResult.Success;

                session["API_KEY_PROTECTED"] = ApiKeyProtector.Protect(key);
                session["API_KEY"] = "";
                session.Log("Jameel Copilot: the API key was encrypted for this machine.");
                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log("Jameel Copilot: the API key could not be encrypted (" + ex.GetType().Name + ").");
                return ActionResult.Failure;
            }
        }
    }
}
