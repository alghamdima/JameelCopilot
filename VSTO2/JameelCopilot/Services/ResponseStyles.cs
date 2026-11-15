using System;

namespace JameelCopilot.Services
{
    /// <summary>
    /// The three response styles a user can pick, and the temperature each
    /// one means within the limits the administrator set: Precise is 0.2 (or
    /// the administrator's temperature if that is lower), Balanced is the
    /// administrator's temperature, Creative is the administrator's ceiling.
    /// </summary>
    internal static class ResponseStyles
    {
        public const string Precise = "precise";
        public const string Balanced = "balanced";
        public const string Creative = "creative";

        /// <summary>One of the three names; anything else is Balanced.</summary>
        public static string Normalize(string style)
        {
            var name = (style ?? "").Trim().ToLowerInvariant();
            return name == Precise || name == Creative ? name : Balanced;
        }

        /// <summary>Null when the administrator switched temperature off: then none is sent.</summary>
        public static double? TemperatureFor(string style)
        {
            var balanced = Config.Temperature;
            if (!balanced.HasValue) return null;

            switch (Normalize(style))
            {
                case Precise:
                    return Math.Min(MachineSettings.PreciseTemperature, balanced.Value);
                case Creative:
                    return Config.TemperatureMax;
                default:
                    return balanced;
            }
        }
    }
}
