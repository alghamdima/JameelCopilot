using System;
using JameelCopilot.Services;

namespace JameelCopilot.Tests.Helpers
{
    /// <summary>
    /// Applies machine settings for one test and puts the previous ones back.
    /// Use inside <c>[Collection("Config")]</c>: Config is static.
    /// </summary>
    internal sealed class ConfigStateScope : IDisposable
    {
        private readonly MachineSettings _before;

        public ConfigStateScope(MachineSettings settings = null)
        {
            _before = Config.Current;
            Config.Apply(settings ?? TestSettings.Configured());
        }

        public void Dispose()
        {
            Config.Apply(_before);
        }
    }
}
