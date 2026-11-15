using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>The read-only About dialog that replaced the Settings editor.</summary>
    [Collection("Config")]
    public class SettingsFormTests
    {
        [Fact]
        public void Describe_ShowsHostModelAndTools_NotTheKey()
        {
            using (new ConfigStateScope(TestSettings.Configured(model: "m1")))
            {
                var lines = SettingsForm.DescribeSettings();

                Assert.Contains(lines, l => l.StartsWith("Version: "));
                Assert.Contains("AI service: llm.example.internal", lines);
                Assert.Contains("Model: m1", lines);
                Assert.Contains("Write tools: none", lines);
                Assert.DoesNotContain(lines, l => l.Contains(TestSettings.Key));
                Assert.DoesNotContain(lines, l => l.Contains("/v1") || l.Contains("https://"));
                Assert.DoesNotContain(lines, l => l.StartsWith("Not configured"));
            }
        }

        [Fact]
        public void Describe_ShowsTheModelSettings()
        {
            using (new ConfigStateScope(TestSettings.Configured()))
            {
                var lines = SettingsForm.DescribeSettings();

                Assert.Contains("Temperature: 0.5 (creative up to 0.9)", lines);
                Assert.DoesNotContain(lines, l => l.StartsWith("Answer length limit"));
            }
            using (new ConfigStateScope(TestSettings.Configured(temperature: "off", maxOutputTokens: "800")))
            {
                var lines = SettingsForm.DescribeSettings();

                Assert.Contains("Temperature: not sent", lines);
                Assert.Contains("Answer length limit: 800 tokens", lines);
            }
        }

        // Review Focus 1: the dialog shows what will be sent, with a dot.
        [Fact]
        public void Describe_OnAnArabicPc_ShowsTheTemperatureWithADot()
        {
            using (TestCultures.ArabicDecimals())
            using (new ConfigStateScope(TestSettings.Configured(temperature: "0.25", temperatureMax: "1.5")))
            {
                Assert.Contains("Temperature: 0.25 (creative up to 1.5)", SettingsForm.DescribeSettings());
            }
        }

        // ShowDialog ends by replacing the thread's WinForms context with a plain
        // one. After About Jameel Copilot was closed, an await on the UI thread resumed on the
        // thread pool and an open pane stopped answering.
        [Fact]
        public void Open_LeavesTheUiThreadsSynchronizationContextInPlace()
        {
            System.Threading.SynchronizationContext before = null, after = null;
            Sta.Run(() =>
            {
                using (new ConfigStateScope(TestSettings.Configured()))
                {
                    before = new System.Windows.Forms.WindowsFormsSynchronizationContext();
                    System.Threading.SynchronizationContext.SetSynchronizationContext(before);
                    var closer = new System.Windows.Forms.Timer { Interval = 200 };
                    closer.Tick += (s, e) =>
                    {
                        foreach (var open in System.Windows.Forms.Application.OpenForms.Cast<System.Windows.Forms.Form>().ToList()) open.Close();
                    };
                    closer.Start();
                    try
                    {
                        SettingsForm.Open(null);
                    }
                    finally
                    {
                        closer.Stop();
                        closer.Dispose();
                    }
                    after = System.Threading.SynchronizationContext.Current;
                }
            });

            Assert.Same(before, after);
        }

        // Every line of the summary is visible: the label is tall enough for the longest summary.
        [Fact]
        public void Dialog_ShowsEveryLineOfTheSummary()
        {
            Sta.Run(() =>
            {
                using (new ConfigStateScope(TestSettings.Configured(maxOutputTokens: "800")))
                using (var form = new SettingsForm(null))
                {
                    var summary = Descendants(form).OfType<Label>().First(l => l.Text.StartsWith("Version: "));
                    var needed = TextRenderer.MeasureText(summary.Text, summary.Font, new System.Drawing.Size(summary.Width, 0),
                        TextFormatFlags.WordBreak).Height;

                    Assert.True(summary.Height >= needed, "label " + summary.Height + " px, text " + needed + " px");
                    Assert.All(Descendants(form).OfType<Button>(), b => Assert.True(b.Bottom <= form.ClientSize.Height));
                }
            });
        }

        [Fact]
        public void Describe_ListsTheEnabledWriteTools()
        {
            using (new ConfigStateScope(TestSettings.Configured(writeTools: "outlook_flag_message,outlook_create_draft")))
            {
                Assert.Contains("Write tools: outlook_create_draft, outlook_flag_message", SettingsForm.DescribeSettings());
            }
        }

        [Fact]
        public void Describe_Unconfigured_ShowsProblem()
        {
            using (new ConfigStateScope(TestSettings.Unconfigured()))
            {
                var lines = SettingsForm.DescribeSettings();

                Assert.Contains("AI service: not set", lines);
                Assert.Contains("Model: not set", lines);
                Assert.Contains(lines, l => l.StartsWith("Not configured: ") && l.Contains("ApiBaseUrl"));
            }
        }

        // Nothing in the dialog can be edited: settings are an administrator's, in HKLM.
        [Fact]
        public void Dialog_IsReadOnly_WithTestAndCloseButtons()
        {
            Sta.Run(() =>
            {
                using (new ConfigStateScope(TestSettings.Configured()))
                using (var form = new SettingsForm(null))
                {
                    var controls = Descendants(form).ToList();

                    Assert.Equal("About Jameel Copilot", form.Text);
                    Assert.Empty(controls.OfType<TextBoxBase>());
                    Assert.Empty(controls.OfType<ComboBox>());
                    Assert.Empty(controls.OfType<CheckBox>());
                    var buttons = controls.OfType<Button>().ToList();
                    Assert.Equal(new[] { "Close", "Test connection" }, buttons.Select(b => b.Text).OrderBy(t => t));
                    // Windows Server 2025 renders default-styled buttons blank.
                    Assert.All(buttons, b => Assert.False(b.UseVisualStyleBackColor));
                }
            });
        }

        private static IEnumerable<Control> Descendants(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                yield return child;
                foreach (var grandchild in Descendants(child)) yield return grandchild;
            }
        }
    }
}
