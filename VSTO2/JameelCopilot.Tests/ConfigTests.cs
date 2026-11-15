using System.Linq;
using JameelCopilot.Services;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    [Collection("Config")]
    public class ConfigTests
    {
        [Fact]
        public void Apply_CopiesEveryValue()
        {
            using (new ConfigStateScope(TestSettings.Configured(
                model: "m9", effort: "low", writeTools: "outlook_create_draft,outlook_flag_message", maxRows: 500)))
            {
                Assert.True(Config.IsConfigured);
                Assert.Null(Config.Problem);
                Assert.Equal(TestSettings.BaseUrl, Config.ApiBaseUrl.AbsoluteUri);
                Assert.Equal(TestSettings.Key, Config.ApiKey);
                Assert.Equal("m9", Config.Model);
                Assert.Equal("low", Config.ReasoningEffort);
                Assert.Equal(new[] { "outlook_create_draft", "outlook_flag_message" }, Config.EnabledWriteTools.OrderBy(t => t));
                Assert.Equal(500, Config.MaxBulkExportRows);
            }
        }

        [Fact]
        public void WriteToolsEnabled_FollowsEnabledWriteTools()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                Assert.False(Config.WriteToolsEnabled);
            }
            using (new ConfigStateScope(TestSettings.Configured(writeTools: "outlook_set_category")))
            {
                Assert.True(Config.WriteToolsEnabled);
            }
        }

        [Fact]
        public void Apply_Unconfigured_SetsProblem()
        {
            using (new ConfigStateScope(TestSettings.Unconfigured()))
            {
                Assert.False(Config.IsConfigured);
                Assert.NotNull(Config.Problem);
                Assert.Null(Config.ApiKey);
                Assert.Empty(Config.EnabledWriteTools);
            }
        }

        [Fact]
        public void Scope_RestoresThePreviousSettings()
        {
            using (new ConfigStateScope(TestSettings.Configured(model: "outer")))
            {
                using (new ConfigStateScope(TestSettings.Configured(model: "inner")))
                {
                    Assert.Equal("inner", Config.Model);
                }
                Assert.Equal("outer", Config.Model);
            }
        }

        [Fact]
        public void AllWriteTools_AreTheFourSafeWrites()
        {
            Assert.Equal(MachineSettings.AllWriteTools, Config.AllWriteTools);
        }
    }
}
