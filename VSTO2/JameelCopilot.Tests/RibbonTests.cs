using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace JameelCopilot.Tests
{
    public class RibbonTests
    {
        private static readonly XNamespace Ui = "http://schemas.microsoft.com/office/2009/07/customui";

        // The XML Outlook is served, for the Inbox and for compose windows.
        private static XDocument Served(string ribbonId) => XDocument.Parse(new Ribbon().GetCustomUI(ribbonId));

        private static string RibbonIdFor(string tab) =>
            tab == "TabMail" ? "Microsoft.Outlook.Explorer" : "Microsoft.Outlook.Mail.Compose";

        private static XElement AiAssistantGroup(string tab) =>
            Served(RibbonIdFor(tab))
                .Descendants(Ui + "tab").Single(t => (string)t.Attribute("idMso") == tab)
                .Elements(Ui + "group").Single(g => (string)g.Attribute("label") == "Jameel Copilot");

        // The read-only About dialog (version, AI service, connection test) is
        // on the ribbon wherever the assistant is.
        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabNewMailMessage")]
        public void AboutButton_IsInTheAiAssistantGroup(string tab)
        {
            var about = AiAssistantGroup(tab).Elements(Ui + "button")
                .SingleOrDefault(b => (string)b.Attribute("onAction") == "OnSettingsClick");

            Assert.NotNull(about);
            Assert.Equal("About Jameel Copilot", (string)about.Attribute("screentip"));
            Assert.False(string.IsNullOrEmpty((string)about.Attribute("imageMso")));
        }

        // The owner wants the two buttons as icons alone: the mark, and the gear for About.
        // Office always shows a large button's label, and the simplified (single-line)
        // ribbon ignores showLabel, so the label is an empty text from GetNoLabel. The
        // name is kept as the tooltip's title (screentip). label and getLabel exclude each other.
        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabNewMailMessage")]
        public void Buttons_ShowNoLabel_AndKeepTheirNameAsTooltip(string tab)
        {
            foreach (var button in AiAssistantGroup(tab).Elements(Ui + "button"))
            {
                Assert.Equal("false", (string)button.Attribute("showLabel"));
                Assert.NotEqual("large", (string)button.Attribute("size"));
                Assert.Null(button.Attribute("label"));
                Assert.Equal("GetNoLabel", (string)button.Attribute("getLabel"));
                Assert.False(string.IsNullOrEmpty((string)button.Attribute("screentip")));
            }
        }

        [Fact]
        public void GetNoLabel_ReturnsEmptyText()
        {
            var callback = typeof(Ribbon).GetMethod("GetNoLabel", BindingFlags.Public | BindingFlags.Instance);
            Assert.NotNull(callback);
            Assert.Equal("", (string)callback.Invoke(new Ribbon(), new object[] { null }));
        }

        // The assistant's own button carries the Jameel Copilot mark; Reports and About keep Office's icons.
        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabNewMailMessage")]
        public void MainButton_ShowsTheJcMark(string tab)
        {
            var main = AiAssistantGroup(tab).Elements(Ui + "button")
                .Single(b => (string)b.Attribute("onAction") == "OnAIAssistantClick");

            Assert.Equal("GetJcImage", (string)main.Attribute("getImage"));
            Assert.Null(main.Attribute("imageMso"));
        }

        [Fact]
        public void GetJcImage_ReturnsAnImageAtLeast32Wide()
        {
            // Through reflection, as Office calls it: the parameter is an Office type the tests do not reference.
            var callback = typeof(Ribbon).GetMethod("GetJcImage", BindingFlags.Public | BindingFlags.Instance);
            using (var image = (System.Drawing.Bitmap)callback.Invoke(new Ribbon(), new object[] { null }))
            {
                Assert.NotNull(image);
                Assert.True(image.Width >= 32 && image.Width == image.Height, image.Width + "x" + image.Height);
            }
        }

        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabNewMailMessage")]
        public void NoQuickToolsMenu(string tab)
        {
            Assert.Empty(AiAssistantGroup(tab).Elements(Ui + "menu"));
        }

        [Fact]
        public void ReadMessageWindow_GetsNoRibbon()
        {
            Assert.Null(new Ribbon().GetCustomUI("Microsoft.Outlook.Mail.Read"));
        }

        [Fact]
        public void NoSupertipMentionsTheOldSignInOrSettings()
        {
            var supertips = Served("Microsoft.Outlook.Explorer").Descendants()
                .Select(e => (string)e.Attribute("supertip"))
                .Where(s => s != null);

            Assert.All(supertips, s => Assert.DoesNotContain("ChatGPT", s));
            Assert.All(supertips, s => Assert.DoesNotContain("password", s));
        }

        // Office drops the add-in's whole ribbon when a callback is missing or
        // two controls share an id.
        [Fact]
        public void EveryCallback_HasAPublicHandler()
        {
            var callbacks = Served("Microsoft.Outlook.Explorer").Descendants()
                .SelectMany(e => e.Attributes())
                .Where(a => a.Name.LocalName.StartsWith("on") || a.Name.LocalName.StartsWith("get"))
                .Select(a => a.Value)
                .Distinct()
                .ToList();

            Assert.Contains("OnSettingsClick", callbacks);
            foreach (var name in callbacks)
            {
                var handler = typeof(Ribbon).GetMethod(name, BindingFlags.Public | BindingFlags.Instance);
                Assert.True(handler != null && handler.GetParameters().Length == 1, "No handler for " + name);
            }
        }

        [Fact]
        public void ControlIds_AreUnique()
        {
            var ids = Served("Microsoft.Outlook.Explorer").Descendants()
                .Select(e => (string)e.Attribute("id"))
                .Where(id => id != null)
                .ToList();

            Assert.Equal(ids.Count, ids.Distinct().Count());
        }
    }
}
