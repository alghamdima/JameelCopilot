using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using JameelCopilot.Services.Prefs;
using JameelCopilot.TaskPane.Chat;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.TaskPane.Chat
{
    /// <summary>Between the chat page's style picker and the user's stored preference.</summary>
    [Collection("Config")]
    public class PrefsBridgeTests
    {
        private readonly List<string> _scripts = new List<string>();
        private readonly UserPrefsService _prefs = new UserPrefsService(null);

        private PrefsBridge NewBridge() => new PrefsBridge(_prefs, script =>
        {
            _scripts.Add(script);
            return Task.CompletedTask;
        });

        [Fact]
        public async Task Push_SendsTheStoredStyle()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                _prefs.SetStyle("creative");

                await NewBridge().PushAsync();

                Assert.Equal(new[] { "jameelcopilot.setStyle(\"creative\");" }, _scripts);
            }
        }

        // The administrator switched temperature off: there is nothing to choose.
        [Fact]
        public async Task Push_TemperatureOff_HidesThePicker()
        {
            using (new ConfigStateScope(TestSettings.Configured(temperature: "off")))
            {
                await NewBridge().PushAsync();

                Assert.Equal(new[] { "jameelcopilot.setStyle(null);" }, _scripts);
            }
        }

        [Fact]
        public void Handle_SetStyle_SavesIt_AndReturnsTrue()
        {
            var handled = NewBridge().Handle("set_style", JObject.Parse("{\"style\":\"precise\"}"));

            Assert.True(handled);
            Assert.Equal("precise", _prefs.Style);
        }

        [Theory]
        [InlineData("send")]
        [InlineData("ready")]
        [InlineData("")]
        [InlineData(null)]
        public void Handle_OtherMessages_ReturnsFalse(string type)
        {
            _prefs.SetStyle("creative");

            Assert.False(NewBridge().Handle(type, JObject.Parse("{\"style\":\"precise\"}")));
            Assert.Equal("creative", _prefs.Style);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("{}")]
        [InlineData("{\"style\":7}")]
        [InlineData("{\"style\":{\"x\":1}}")]
        [InlineData("{\"style\":\"<script>\"}")]
        public void Handle_SetStyle_WithNoUsableStyle_FallsBackToBalanced(string payload)
        {
            _prefs.SetStyle("creative");

            var handled = NewBridge().Handle("set_style", payload == null ? null : JObject.Parse(payload));

            Assert.True(handled);
            Assert.Equal("balanced", _prefs.Style);
        }

        [Fact]
        public void Temperature_FollowsTheStyle()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                var bridge = NewBridge();

                _prefs.SetStyle("precise");
                Assert.Equal(0.2, bridge.Temperature);
                _prefs.SetStyle("creative");
                Assert.Equal(0.9, bridge.Temperature);
                _prefs.SetStyle("balanced");
                Assert.Equal(0.5, bridge.Temperature);
            }
        }

        [Fact]
        public void NoPrefsService_StillWorks()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                var bridge = new PrefsBridge(null, script => Task.CompletedTask);

                Assert.True(bridge.Handle("set_style", JObject.Parse("{\"style\":\"creative\"}")));
                Assert.Equal(0.9, bridge.Temperature);
            }
        }
    }
}
