using System;
using System.Collections.Generic;
using System.Security;
using System.Security.Cryptography;
using JameelCopilot.Services;
using JameelCopilot.Services.Tools;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    public class MachineSettingsTests
    {
        private static Dictionary<string, object> Valid() => new Dictionary<string, object>
        {
            ["ApiBaseUrl"] = "https://llm.example.internal/v1",
            ["ApiKeyProtected"] = "BLOB",
            ["Model"] = "m1",
        };

        // The fake key store: only "BLOB" decrypts.
        private static MachineSettings Read(Dictionary<string, object> values) =>
            MachineSettings.Read(
                name => values.TryGetValue(name, out var value) ? value : null,
                blob => blob == "BLOB" ? "fake-key-for-tests" : null);

        private static MachineSettings ReadWith(string name, object value)
        {
            var values = Valid();
            values[name] = value;
            return Read(values);
        }

        [Fact]
        public void AllPresent_IsConfigured()
        {
            var settings = Read(Valid());

            Assert.True(settings.IsConfigured);
            Assert.Null(settings.Problem);
            Assert.Equal("https://llm.example.internal/v1", settings.ApiBaseUrl.AbsoluteUri);
            Assert.Equal("fake-key-for-tests", settings.ApiKey);
            Assert.Equal("m1", settings.Model);
            Assert.Null(settings.ReasoningEffort);
        }

        [Fact]
        public void TrailingSlashAndSpaces_AreTrimmed()
        {
            var settings = ReadWith("ApiBaseUrl", "  https://llm.example.internal/v1/ ");

            Assert.True(settings.IsConfigured);
            Assert.Equal("https://llm.example.internal/v1", settings.ApiBaseUrl.AbsoluteUri);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("http://llm.example.internal/v1")]
        [InlineData("llm.example.internal/v1")]
        [InlineData("ftp://llm.example.internal/v1")]
        [InlineData("https://")]
        public void BadUrl_NotConfigured(string url)
        {
            var settings = ReadWith("ApiBaseUrl", url);

            Assert.Null(settings.ApiBaseUrl);
            Assert.False(settings.IsConfigured);
            Assert.Contains("ApiBaseUrl", settings.Problem);
        }

        [Fact]
        public void UndecryptableKey_NotConfigured()
        {
            var settings = ReadWith("ApiKeyProtected", "OTHER-BLOB");

            Assert.Null(settings.ApiKey);
            Assert.False(settings.IsConfigured);
            Assert.Contains("ApiKeyProtected", settings.Problem);
            Assert.DoesNotContain("OTHER-BLOB", settings.Problem);
        }

        [Fact]
        public void MissingKey_NotConfigured()
        {
            var values = Valid();
            values.Remove("ApiKeyProtected");

            var settings = Read(values);

            Assert.False(settings.IsConfigured);
            Assert.Contains("ApiKeyProtected", settings.Problem);
        }

        [Fact]
        public void KeyStoreThrows_NotConfigured()
        {
            var values = Valid();

            var settings = MachineSettings.Read(
                name => values[name == "ApiBaseUrl" || name == "ApiKeyProtected" || name == "Model" ? name : "Model"],
                blob => throw new CryptographicException("bad data"));

            Assert.Null(settings.ApiKey);
            Assert.False(settings.IsConfigured);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("  ")]
        public void BlankModel_NotConfigured(string model)
        {
            var settings = ReadWith("Model", model);

            Assert.Null(settings.Model);
            Assert.False(settings.IsConfigured);
            Assert.Contains("Model", settings.Problem);
        }

        [Fact]
        public void ModelAndEffort_AreTrimmed()
        {
            var values = Valid();
            values["Model"] = " m1 ";
            values["ReasoningEffort"] = " low ";

            var settings = Read(values);

            Assert.Equal("m1", settings.Model);
            Assert.Equal("low", settings.ReasoningEffort);
        }

        [Fact]
        public void BlankEffort_IsNull()
        {
            Assert.Null(ReadWith("ReasoningEffort", "  ").ReasoningEffort);
        }

        [Fact]
        public void WriteTools_DefaultEmpty()
        {
            Assert.Empty(Read(Valid()).EnabledWriteTools);
        }

        [Fact]
        public void WriteTools_UnknownNamesDropped()
        {
            var settings = ReadWith("EnabledWriteTools", "outlook_create_draft, bogus ,, OUTLOOK_FLAG_MESSAGE");

            Assert.Equal(new[] { "outlook_create_draft" }, settings.EnabledWriteTools);
        }

        [Fact]
        public void WriteTools_AllFourAccepted()
        {
            var settings = ReadWith("EnabledWriteTools", string.Join(",", MachineSettings.AllWriteTools));

            Assert.Equal(4, MachineSettings.AllWriteTools.Length);
            Assert.True(settings.EnabledWriteTools.SetEquals(MachineSettings.AllWriteTools));
        }

        [Fact]
        public void MaxRows_DefaultsTo2000()
        {
            Assert.Equal(2000, Read(Valid()).MaxBulkExportRows);
            Assert.Equal(2000, MachineSettings.DefaultMaxBulkExportRows);
        }

        [Theory]
        [InlineData(0, 1)]
        [InlineData(-5, 1)]
        [InlineData(500, 500)]
        [InlineData(999999, BulkExportRowCap.Max)]
        public void MaxRows_Clamped(int given, int expected)
        {
            Assert.Equal(expected, ReadWith("MaxBulkExportRows", given).MaxBulkExportRows);
        }

        [Fact]
        public void MaxRows_AcceptsNumericString()
        {
            Assert.Equal(500, ReadWith("MaxBulkExportRows", " 500 ").MaxBulkExportRows);
        }

        [Theory]
        [InlineData("abc")]
        [InlineData("")]
        public void MaxRows_NonNumeric_UsesDefault(string value)
        {
            Assert.Equal(2000, ReadWith("MaxBulkExportRows", value).MaxBulkExportRows);
        }

        // -- model settings ---------------------------------------------------

        [Fact]
        public void ModelSettings_Defaults()
        {
            var settings = Read(Valid());

            Assert.Equal(0.5, settings.Temperature);
            Assert.Equal(0.9, settings.TemperatureMax);
            Assert.Null(settings.MaxOutputTokens);
        }

        [Theory]
        [InlineData("0.3", 0.3)]
        [InlineData(" 0.3 ", 0.3)]
        [InlineData("0", 0.0)]
        [InlineData("5", 2.0)]
        [InlineData("-1", 0.0)]
        [InlineData("abc", 0.5)]
        [InlineData("", 0.5)]
        [InlineData("0,3", 0.5)]
        [InlineData("NaN", 0.5)]
        public void Temperature_Parsed(string value, double expected)
        {
            Assert.Equal(expected, ReadWith("Temperature", value).Temperature);
        }

        // For a model that rejects the parameter.
        [Theory]
        [InlineData("off")]
        [InlineData("OFF")]
        [InlineData(" Off ")]
        public void Temperature_Off_IsNull(string value)
        {
            var settings = ReadWith("Temperature", value);

            Assert.Null(settings.Temperature);
            Assert.True(settings.IsConfigured);
        }

        // Review Focus 1: Arabic regional settings write decimals with another separator.
        [Fact]
        public void Temperature_OnAnArabicPc_IsStillReadWithADot()
        {
            using (TestCultures.ArabicDecimals())
            {
                // Not the default (0.5): a value that failed to parse would fall back to it.
                var settings = ReadWith("Temperature", "0.3");

                Assert.Equal(0.3, settings.Temperature);
                Assert.Equal(0.9, settings.TemperatureMax);
                Assert.True(settings.IsConfigured);
            }
        }

        [Theory]
        [InlineData("0.7", "0.4", 0.7)]
        [InlineData("0.3", "1.2", 1.2)]
        [InlineData("0.3", "9", 2.0)]
        [InlineData("0.3", "x", 0.9)]
        [InlineData("1.5", "", 1.5)]
        [InlineData("off", "0.8", 0.8)]
        public void TemperatureMax_IsNeverBelowTemperature(string temperature, string max, double expected)
        {
            var values = Valid();
            values["Temperature"] = temperature;
            values["TemperatureMax"] = max;

            Assert.Equal(expected, Read(values).TemperatureMax);
        }

        [Theory]
        [InlineData("800", 800)]
        [InlineData(" 800 ", 800)]
        [InlineData("0", null)]
        [InlineData("-5", null)]
        [InlineData("x", null)]
        [InlineData("", null)]
        [InlineData("1.5", null)]
        [InlineData("999999999", 100000)]
        public void MaxOutputTokens_Parsed(string value, int? expected)
        {
            Assert.Equal(expected, ReadWith("MaxOutputTokens", value).MaxOutputTokens);
        }

        // The installer writes strings, but a baseline may write numbers.
        [Fact]
        public void MaxOutputTokens_AcceptsARegistryNumber()
        {
            Assert.Equal(800, ReadWith("MaxOutputTokens", 800).MaxOutputTokens);
        }

        // A value of the wrong registry type (here REG_BINARY) must not stop Outlook from starting.
        [Fact]
        public void WrongRegistryTypes_DoNotThrow()
        {
            var bytes = new byte[] { 1 };
            var values = new Dictionary<string, object>
            {
                ["ApiBaseUrl"] = bytes, ["ApiKeyProtected"] = bytes, ["Model"] = bytes,
                ["ReasoningEffort"] = bytes, ["EnabledWriteTools"] = bytes, ["MaxBulkExportRows"] = bytes,
                ["Temperature"] = bytes, ["TemperatureMax"] = bytes, ["MaxOutputTokens"] = bytes,
            };

            var settings = Read(values);

            Assert.False(settings.IsConfigured);
            Assert.NotNull(settings.Problem);
            Assert.Empty(settings.EnabledWriteTools);
            Assert.Equal(2000, settings.MaxBulkExportRows);
            Assert.Equal(0.5, settings.Temperature);
            Assert.Equal(0.9, settings.TemperatureMax);
            Assert.Null(settings.MaxOutputTokens);
        }

        [Fact]
        public void ReaderThrows_NotConfigured()
        {
            var settings = MachineSettings.Read(
                name => throw new SecurityException("denied"),
                blob => "fake-key-for-tests");

            Assert.False(settings.IsConfigured);
            Assert.NotNull(settings.Problem);
            Assert.Empty(settings.EnabledWriteTools);
            Assert.Equal(2000, settings.MaxBulkExportRows);
        }

        [Fact]
        public void NothingSet_ProblemNamesTheFirstMissingSetting()
        {
            var settings = Read(new Dictionary<string, object>());

            Assert.False(settings.IsConfigured);
            Assert.Contains("ApiBaseUrl", settings.Problem);
        }

        [Fact]
        public void RegistryPath_IsJcOutlook()
        {
            Assert.Equal(@"SOFTWARE\JameelCopilot", MachineSettings.RegistryPath);
        }
    }

    public class ApiKeyProtectorTests
    {
        [Fact]
        public void RoundTrip()
        {
            Assert.Equal("fake-key", ApiKeyProtector.Unprotect(ApiKeyProtector.Protect("fake-key")));
        }

        [Fact]
        public void Protect_OutputIsBase64AndNotThePlainKey()
        {
            var blob = ApiKeyProtector.Protect("fake-key-0123456789");

            Assert.DoesNotContain("fake-key", blob);
            Assert.NotEmpty(Convert.FromBase64String(blob));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("not base64!")]
        [InlineData("AQID")] // base64 of bytes 1,2,3: valid base64, not a protected blob
        public void Unprotect_BadInput_ReturnsNull(string input)
        {
            Assert.Null(ApiKeyProtector.Unprotect(input));
        }
    }
}
