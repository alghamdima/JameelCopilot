namespace JameelCopilot.Services.Prefs
{
    /// <summary>
    /// A place that keeps small pieces of the user's own Jameel Copilot data by name.
    /// Implementations never throw: a store that is unavailable answers
    /// null and false.
    /// </summary>
    public interface IUserStore
    {
        /// <summary>
        /// The stored text; "" when nothing is stored under the name; null when
        /// the store could not be read. The two are different: a caller that
        /// keeps a list must not write it back after a read that failed.
        /// </summary>
        string Load(string name);

        /// <summary>False when it could not be saved.</summary>
        bool Save(string name, string json);

        /// <summary>False when it could not be removed.</summary>
        bool Delete(string name);
    }
}
