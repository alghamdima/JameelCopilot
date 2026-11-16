using System;
using System.IO;
using JameelCopilot.Services;
using WixToolset.Dtf.WindowsInstaller;

namespace JameelCopilot.Installer.Actions
{
    public static class ProtectApiKeyAction
    {
        /// <summary>
        /// Turns the API_KEY given on the command line, or the key in the file
        /// API_KEY_FILE names, into API_KEY_PROTECTED, encrypted for this
        /// machine, and clears API_KEY, so only the encrypted form is ever
        /// written to the registry. The key is never written to the log.
        /// </summary>
        [CustomAction]
        public static ActionResult ProtectApiKey(Session session)
        {
            try
            {
                var key = session["API_KEY"];
                var file = session["API_KEY_FILE"];
                if (string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(file))
                {
                    // A file keeps the key out of the command line, which process auditing (event 4688) records.
                    // A relative name is next to the MSI: this action's own working folder is System32.
                    key = File.ReadAllText(Path.Combine(Path.GetDirectoryName(session["OriginalDatabase"]), file)).Trim();
                    if (key.Length == 0)
                    {
                        session.Log("Jameel Copilot: the file API_KEY_FILE names is empty.");
                        return ActionResult.Failure;
                    }
                }
                if (string.IsNullOrEmpty(key)) return ActionResult.Success;

                session["API_KEY_PROTECTED"] = ApiKeyProtector.Protect(key);
                session["API_KEY"] = "";
                session.Log("Jameel Copilot: the API key was encrypted for this machine.");
                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log("Jameel Copilot: the API key could not be read or encrypted (" + ex.GetType().Name + ").");
                return ActionResult.Failure;
            }
        }
    }
}
