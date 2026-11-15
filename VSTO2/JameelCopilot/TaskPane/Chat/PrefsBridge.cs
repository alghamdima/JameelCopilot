using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services;
using JameelCopilot.Services.Prefs;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// Connects the chat page's response-style picker to the user's stored
    /// preference: tells the page which style is in use, takes the page's
    /// "set_style" message, and gives a turn its temperature.
    /// </summary>
    internal sealed class PrefsBridge
    {
        private readonly UserPrefsService _prefs;
        private readonly Func<string, Task> _runScript;

        public PrefsBridge(UserPrefsService prefs, Func<string, Task> runScript)
        {
            _prefs = prefs ?? new UserPrefsService(null);
            _runScript = runScript ?? throw new ArgumentNullException(nameof(runScript));
        }

        /// <summary>The temperature for the next turn; null sends none.</summary>
        public double? Temperature => ResponseStyles.TemperatureFor(_prefs.Style);

        /// <summary>Shows the picker with the stored style, or hides it when temperature is switched off.</summary>
        public Task PushAsync()
        {
            // The style is one of three fixed names, so it is safe inside the script.
            return _runScript(Config.Temperature.HasValue
                ? "jameelcopilot.setStyle(\"" + _prefs.Style + "\");"
                : "jameelcopilot.setStyle(null);");
        }

        /// <summary>True when the message was the picker's.</summary>
        public bool Handle(string type, JObject payload)
        {
            if (type != "set_style") return false;
            var style = payload?["style"];
            _prefs.SetStyle(style != null && style.Type == JTokenType.String ? (string)style : null);
            return true;
        }
    }
}
