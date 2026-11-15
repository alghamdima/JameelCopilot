using System;
using System.Collections.Generic;
using JameelCopilot.Services.Prefs;

namespace JameelCopilot.Tests.Helpers
{
    /// <summary>A mailbox store in memory, which can be made to fail the way the real one does.</summary>
    internal sealed class FakeUserStore : IUserStore
    {
        public readonly Dictionary<string, string> Items = new Dictionary<string, string>(StringComparer.Ordinal);
        public bool Unavailable;   // answers null and false
        public bool Throws;        // the real one never does; a faulty one must not break Jameel Copilot either
        public int FailNextLoads;  // that many reads fail (null) while writes still work

        public string Load(string name)
        {
            if (Throws) throw new InvalidOperationException("store unavailable");
            if (Unavailable) return null;
            if (FailNextLoads > 0)
            {
                FailNextLoads--;
                return null;
            }
            // Nothing stored under the name is an empty text; null means the read failed.
            return Items.TryGetValue(name, out var json) ? json : "";
        }

        public bool Save(string name, string json)
        {
            if (Throws) throw new InvalidOperationException("store unavailable");
            if (Unavailable) return false;
            Items[name] = json;
            return true;
        }

        public bool Delete(string name)
        {
            if (Throws) throw new InvalidOperationException("store unavailable");
            if (Unavailable) return false;
            return Items.Remove(name);
        }
    }
}
