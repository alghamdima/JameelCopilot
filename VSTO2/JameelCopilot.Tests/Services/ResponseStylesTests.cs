using JameelCopilot.Services;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>The user's three response styles, within the limits the administrator set.</summary>
    [Collection("Config")]
    public class ResponseStylesTests
    {
        [Theory]
        [InlineData("precise", 0.2)]
        [InlineData("balanced", 0.5)]
        [InlineData("creative", 0.9)]
        [InlineData(null, 0.5)]
        [InlineData("loud", 0.5)]
        public void TemperatureFor_Defaults(string style, double expected)
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                Assert.Equal(expected, ResponseStyles.TemperatureFor(style));
            }
        }

        [Fact]
        public void Precise_IsNeverAboveTheAdminsTemperature()
        {
            using (new ConfigStateScope(TestSettings.Configured(temperature: "0.1")))
            {
                Assert.Equal(0.1, ResponseStyles.TemperatureFor("precise"));
            }
        }

        [Fact]
        public void Creative_IsTheAdminsCeiling()
        {
            using (new ConfigStateScope(TestSettings.Configured(temperature: "0.3", temperatureMax: "0.6")))
            {
                Assert.Equal(0.6, ResponseStyles.TemperatureFor("creative"));
                Assert.Equal(0.3, ResponseStyles.TemperatureFor("balanced"));
            }
        }

        // For a model that rejects the parameter: no style sends one.
        [Theory]
        [InlineData("precise")]
        [InlineData("balanced")]
        [InlineData("creative")]
        public void TemperatureOff_GivesNoTemperature(string style)
        {
            using (new ConfigStateScope(TestSettings.Configured(temperature: "off")))
            {
                Assert.Null(ResponseStyles.TemperatureFor(style));
            }
        }

        [Theory]
        [InlineData("PRECISE", "precise")]
        [InlineData(" creative ", "creative")]
        [InlineData("balanced", "balanced")]
        [InlineData("", "balanced")]
        [InlineData(null, "balanced")]
        [InlineData("loud", "balanced")]
        public void Normalize(string given, string expected)
        {
            Assert.Equal(expected, ResponseStyles.Normalize(given));
        }
    }
}
