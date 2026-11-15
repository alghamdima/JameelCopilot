using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using JameelCopilot.Services;

namespace JameelCopilot
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            // One XML serves the compose window and the Inbox: Office ignores the
            // tab that doesn't belong to the current window. Other windows (an
            // opened email, a meeting) get no Jameel Copilot ribbon.
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("JameelCopilot.Ribbon.xml");
            }
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAIAssistantClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }

        public void OnSettingsClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowSettings();
        }

        /// <summary>
        /// Jameel Copilot's buttons show their icon alone: the simplified ribbon ignores showLabel,
        /// so the label is empty. The name is in the XML as the tooltip (screentip).
        /// </summary>
        public string GetNoLabel(Office.IRibbonControl control)
        {
            return "";
        }

        /// <summary>The Jameel Copilot mark for the assistant's button, at the size the screen's scaling needs.</summary>
        public Bitmap GetJcImage(Office.IRibbonControl control)
        {
            try
            {
                float dpi;
                using (var screen = Graphics.FromHwnd(IntPtr.Zero))
                {
                    dpi = screen.DpiX;
                }
                return LogoImages.Sized((int)Math.Round(32 * dpi / 96));
            }
            catch (Exception)
            {
                // Office then shows the button without a picture; the ribbon still loads.
                return null;
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string name in resourceNames)
            {
                if (name.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase))
                {
                    using (StreamReader reader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }
    }
}