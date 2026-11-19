using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using JameelCopilot.Diagnostics;

namespace JameelCopilot.TaskPane.Chat
{
    /// <summary>
    /// One-shot bootstrap helper for the WebView2 control that hosts the
    /// Phase 2 chat surface. Responsibilities:
    /// <list type="bullet">
    ///   <item>Resolve / create the per-machine WebView2 data folder under
    ///         <c>%LOCALAPPDATA%\JameelCopilot\WebView2Data</c> so multiple Outlook
    ///         processes don't fight over the same singleton runtime instance.</item>
    ///   <item>Materialize the bundled <c>WebUI/</c> assets (HTML, CSS, JS,
    ///         vendored libs) into <c>%LOCALAPPDATA%\JameelCopilot\WebUI</c>.</item>
    ///   <item>Map the <c>jameelcopilot.local</c> virtual host to that folder so
    ///         relative <c>fetch()</c>s and <c>&lt;script src&gt;</c>s resolve
    ///         without disk-path leakage.</item>
    ///   <item>Surface <see cref="WebView2Initialized"/> /
    ///         <see cref="WebView2InitializationFailed"/> events the
    ///         <c>ChatController</c> awaits before pushing the first message.</item>
    /// </list>
    /// </summary>
    public static class WebView2Bootstrap
    {
        public const string VirtualHost = "jameelcopilot.local";

        private static readonly string PageVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// The page's address. It carries Jameel Copilot's version, and so do the script
        /// and style addresses inside the page: after an upgrade the browser
        /// asks for addresses it has never seen, and cannot run an old script
        /// from its cache against the new host.
        /// </summary>
        public static string PageUrl => "https://" + VirtualHost + "/index.html?v=" + PageVersion;

        internal static string VersionedIndex(string html, string version)
        {
            return (html ?? "")
                .Replace("href=\"styles.css\"", "href=\"styles.css?v=" + version + "\"")
                .Replace("src=\"markdown.js\"", "src=\"markdown.js?v=" + version + "\"")
                .Replace("src=\"chat.js\"", "src=\"chat.js?v=" + version + "\"");
        }

        private static string LocalAppDataRoot =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JameelCopilot");

        public static string WebUiFolder => Path.Combine(LocalAppDataRoot, "WebUI");
        public static string WebView2DataFolder => Path.Combine(LocalAppDataRoot, "WebView2Data");
        public static string PdfWebView2DataFolder => Path.Combine(LocalAppDataRoot, "WebView2PdfData");

        /// <summary>
        /// Initialize a WebView2 control, extract embedded WebUI resources,
        /// and wire up the virtual host mapping. Caller must add the control
        /// to its parent before calling this. Throws on failure - caller
        /// is expected to fall back to a friendly "install WebView2" panel.
        /// </summary>
        public static Task InitializeAsync(WebView2 host)
        {
            return InitializeAsync(host, WebUiFolder, WebView2DataFolder);
        }

        // Test seam: the same start-up, with the pages and the browser profile in other folders.
        internal static async Task InitializeAsync(WebView2 host, string webUiFolder, string dataFolder)
        {
            if (host == null) throw new ArgumentNullException(nameof(host));

            TraceLog.Write("Bootstrap.InitializeAsync entered", "WebView2Bootstrap");
            Directory.CreateDirectory(webUiFolder);
            Directory.CreateDirectory(dataFolder);
            TraceLog.Write("Folders ensured", "WebView2Bootstrap");
            ExtractEmbeddedWebUi(webUiFolder);
            TraceLog.Write("Embedded WebUI extracted", "WebView2Bootstrap");

            TraceLog.Write(">> CoreWebView2Environment.CreateAsync", "WebView2Bootstrap");
            var env = await CreateEnvironmentAsync(dataFolder).ConfigureAwait(true);
            TraceLog.Write("<< CreateAsync returned", "WebView2Bootstrap");

            TraceLog.Write(">> host.EnsureCoreWebView2Async", "WebView2Bootstrap");
            await host.EnsureCoreWebView2Async(env).ConfigureAwait(true);
            TraceLog.Write("<< EnsureCoreWebView2Async returned", "WebView2Bootstrap");

            LockDown(host.CoreWebView2);
            TraceLog.Write("Locked down", "WebView2Bootstrap");

            host.CoreWebView2.SetVirtualHostNameToFolderMapping(
                VirtualHost,
                webUiFolder,
                CoreWebView2HostResourceAccessKind.DenyCors);
            TraceLog.Write("Virtual host mapping set", "WebView2Bootstrap");
        }

        /// <summary>
        /// A WebView2 environment whose crash dumps stay on this PC instead of
        /// being sent to Microsoft.
        /// </summary>
        public static Task<CoreWebView2Environment> CreateEnvironmentAsync(string userDataFolder)
        {
            var options = new CoreWebView2EnvironmentOptions { IsCustomCrashReportingEnabled = true };
            return CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
        }

        /// <summary>
        /// Every WebView2 the add-in creates goes through here. The pages show
        /// model output, which a crafted email can influence, so the view may
        /// load nothing but the add-in's own files, may not navigate or open
        /// windows, and has the browser features it doesn't need switched off.
        /// </summary>
        /// <param name="allowHostHtml">
        /// True only for the PDF renderer, whose page is handed over with
        /// NavigateToString (a data: or about:blank address).
        /// </param>
        public static void LockDown(CoreWebView2 core, bool allowHostHtml = false)
        {
            if (core == null) throw new ArgumentNullException(nameof(core));

            var settings = core.Settings;
            settings.AreDevToolsEnabled = false;
            settings.AreDefaultContextMenusEnabled = false;
            settings.IsStatusBarEnabled = false;
            settings.AreBrowserAcceleratorKeysEnabled = false;
            settings.IsZoomControlEnabled = false;
            settings.AreHostObjectsAllowed = false;
            settings.AreDefaultScriptDialogsEnabled = false;
            settings.IsGeneralAutofillEnabled = false;
            settings.IsPasswordAutosaveEnabled = false;
            // The pages are local files; a reputation lookup would only tell Microsoft about them.
            settings.IsReputationCheckingRequired = false;

            core.NavigationStarting += (sender, e) =>
            {
                if (!IsAllowedNavigation(e.Uri, allowHostHtml)) e.Cancel = true;
            };
            core.FrameNavigationStarting += (sender, e) =>
            {
                if (!IsAllowedNavigation(e.Uri, allowHostHtml)) e.Cancel = true;
            };
            core.NewWindowRequested += (sender, e) => e.Handled = true;

            // Behind the pages' Content-Security-Policy: refuse any request that leaves the add-in.
            core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
            core.WebResourceRequested += (sender, e) =>
            {
                if (!IsAllowedResource(e.Request.Uri))
                {
                    e.Response = core.Environment.CreateWebResourceResponse(null, 403, "Blocked", "");
                }
            };
        }

        /// <summary>True only for the add-in's own pages: https on the virtual host.</summary>
        public static bool IsTrustedSource(string uri)
        {
            return Uri.TryCreate(uri, UriKind.Absolute, out var parsed)
                && parsed.Scheme == Uri.UriSchemeHttps
                && string.IsNullOrEmpty(parsed.UserInfo)
                && string.Equals(parsed.Host, VirtualHost, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsAllowedNavigation(string uri, bool allowHostHtml)
        {
            if (IsTrustedSource(uri)) return true;
            return allowHostHtml && uri != null
                && (uri == "about:blank" || uri.StartsWith("data:text/html", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsAllowedResource(string uri)
        {
            return IsTrustedSource(uri) || (uri != null
                && (uri == "about:blank" || uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase)));
        }

        /// <summary>Returns true if the Evergreen WebView2 runtime is installed on this machine.</summary>
        public static bool IsRuntimeInstalled()
        {
            try
            {
                var version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(version);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static readonly object _extractLock = new object();

        /// <summary>
        /// Copies every <c>WebUI\*</c> embedded resource from the JameelCopilot
        /// assembly into the given folder. Replaces any existing
        /// files so an upgrade always lands the latest UI. Resources are
        /// embedded with logical names like
        /// <c>JameelCopilot.WebUI.index.html</c>.
        /// </summary>
        internal static void ExtractEmbeddedWebUi(string webUiFolder)
        {
            lock (_extractLock)
            {
                Directory.CreateDirectory(webUiFolder);
                var asm = Assembly.GetExecutingAssembly();
                const string prefix = "JameelCopilot.WebUI.";
                foreach (var name in asm.GetManifestResourceNames())
                {
                    if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                    var relative = ResourceNameToRelativePath(name.Substring(prefix.Length));
                    var targetPath = Path.Combine(webUiFolder, relative);
                    Directory.CreateDirectory(Path.GetDirectoryName(targetPath));

                    using (var src = asm.GetManifestResourceStream(name))
                    {
                        if (relative == "index.html")
                        {
                            using (var reader = new StreamReader(src))
                            {
                                File.WriteAllText(targetPath, VersionedIndex(reader.ReadToEnd(), PageVersion));
                            }
                            continue;
                        }
                        using (var dst = File.Create(targetPath))
                        {
                            src.CopyTo(dst);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Convert a resource name like <c>styles.css</c> or
        /// <c>vendor.marked.min.js</c> into a relative file path. The last
        /// dot is treated as the extension separator; all earlier dots are
        /// folder separators. So <c>vendor.marked.min.js</c> becomes
        /// <c>vendor\marked.min.js</c>.
        /// </summary>
        public static string ResourceNameToRelativePath(string resourceTail)
        {
            if (string.IsNullOrEmpty(resourceTail)) return resourceTail;
            var lastDot = resourceTail.LastIndexOf('.');
            if (lastDot <= 0)
            {
                return resourceTail;
            }
            var stem = resourceTail.Substring(0, lastDot);
            var ext = resourceTail.Substring(lastDot); // includes the dot
            // Special case: if the stem has a hyphen-separated minified suffix
            // like ".min", we want to keep it on the filename, not on a
            // folder. So treat the LAST two dots as one logical extension.
            if (stem.EndsWith(".min", StringComparison.Ordinal))
            {
                var lastDot2 = stem.LastIndexOf('.');
                if (lastDot2 > 0)
                {
                    var stem2 = stem.Substring(0, lastDot2);
                    var minSegment = stem.Substring(lastDot2); // ".min"
                    return stem2.Replace('.', Path.DirectorySeparatorChar) + minSegment + ext;
                }
            }
            return stem.Replace('.', Path.DirectorySeparatorChar) + ext;
        }
    }
}
