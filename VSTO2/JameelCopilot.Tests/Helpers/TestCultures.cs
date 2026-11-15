using System;
using System.Globalization;
using System.Threading;

namespace JameelCopilot.Tests.Helpers
{
    internal static class TestCultures
    {
        /// <summary>
        /// Runs the test as a PC whose regional settings write decimals with
        /// the Arabic decimal separator (U+066B). Plain ar-SA proves nothing: in .NET
        /// it writes a dot, so a test under it passes whatever the code does.
        /// </summary>
        public static IDisposable ArabicDecimals()
        {
            var culture = (CultureInfo)new CultureInfo("ar-SA").Clone();
            culture.NumberFormat.NumberDecimalSeparator = "\u066B";
            Assert(0.5.ToString(culture) != "0.5");
            return new Scope(culture);
        }

        private static void Assert(bool condition)
        {
            if (!condition) throw new InvalidOperationException("The test culture still writes decimals with a dot.");
        }

        private sealed class Scope : IDisposable
        {
            private readonly CultureInfo _before = Thread.CurrentThread.CurrentCulture;

            public Scope(CultureInfo culture)
            {
                Thread.CurrentThread.CurrentCulture = culture;
            }

            public void Dispose()
            {
                Thread.CurrentThread.CurrentCulture = _before;
            }
        }
    }
}
