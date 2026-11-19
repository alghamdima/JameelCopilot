using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JameelCopilot.Services.Prefs
{
    /// <summary>
    /// The user's own Jameel Copilot preferences: for now the response style. Loaded
    /// from the store on first use and kept in memory; a store that is
    /// missing, fails or throws leaves the defaults and never stops a turn.
    /// </summary>
    public sealed class UserPrefsService
    {
        private const string ItemName = "Preferences";

        private readonly IUserStore _store;
        private readonly object _lock = new object();
        private string _style;   // null until a valid read or explicit SetStyle

        /// <param name="store">Null keeps the choice for this session only.</param>
        public UserPrefsService(IUserStore store)
        {
            _store = store;
        }

        /// <summary>One of the <c>ResponseStyles</c> names.</summary>
        public string Style
        {
            get
            {
                lock (_lock)
                {
                    if (_style != null) return _style;
                    var loaded = Load(out var storeAnswered);
                    if (storeAnswered) _style = loaded;
                    return loaded;
                }
            }
        }

        public void SetStyle(string style)
        {
            var normalized = ResponseStyles.Normalize(style);
            lock (_lock)
            {
                _style = normalized;
                try
                {
                    _store?.Save(ItemName, new JObject(new JProperty("v", 1), new JProperty("style", normalized)).ToString(Formatting.None));
                }
                catch (Exception)
                {
                    // The choice holds for this session.
                }
            }
        }

        private string Load(out bool storeAnswered)
        {
            storeAnswered = _store == null;
            try
            {
                var json = _store?.Load(ItemName);
                if (json == null) return ResponseStyles.Balanced;
                storeAnswered = true;
                if (string.IsNullOrWhiteSpace(json)) return ResponseStyles.Balanced;
                var style = (JToken.Parse(json) as JObject)?["style"];
                return ResponseStyles.Normalize(style != null && style.Type == JTokenType.String ? (string)style : null);
            }
            catch (Exception)
            {
                return ResponseStyles.Balanced;
            }
        }
    }
}
