using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// What the add-in project must reference so that the installed add-in has
    /// every file it needs. The tests carry their own copies of the packages,
    /// so a file missing from the add-in's own output does not fail them.
    /// </summary>
    public class ProjectReferenceTests
    {
        private static readonly XNamespace Ns = "http://schemas.microsoft.com/developer/msbuild/2003";

        private static readonly string ProjectDir =
            Path.GetDirectoryName(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));

        private static XDocument Project() => XDocument.Load(Path.Combine(ProjectDir, "JameelCopilot.csproj"));

        // The project names each NuGet assembly with an exact version and a
        // path. When the two disagree MSBuild drops the reference without a
        // warning and the DLL is missing from the installed add-in.
        [Fact]
        public void EveryPackageReference_MatchesTheAssemblyItPointsAt()
        {
            var problems = new List<string>();
            var checkedCount = 0;

            foreach (var reference in Project().Descendants(Ns + "Reference"))
            {
                var hintPath = (string)reference.Element(Ns + "HintPath");
                if (hintPath == null) continue;
                checkedCount++;

                var declared = new AssemblyName((string)reference.Attribute("Include"));
                var file = Path.GetFullPath(Path.Combine(ProjectDir, hintPath));
                if (!File.Exists(file))
                {
                    problems.Add(declared.Name + ": " + hintPath + " does not exist (restore the packages first)");
                    continue;
                }

                var actual = AssemblyName.GetAssemblyName(file);
                if (declared.Version != null && declared.Version != actual.Version)
                {
                    problems.Add(declared.Name + ": the project asks for " + declared.Version + " but " + hintPath + " is " + actual.Version);
                }
            }

            Assert.True(checkedCount >= 10, "Expected the project's package references; found " + checkedCount);
            Assert.True(problems.Count == 0, string.Join("\n", problems));
        }

        // The Office interop assemblies are not installed with the add-in: the
        // types it uses are compiled into it. A reference that is not embedded
        // builds, and then fails on the PC with "could not load file".
        [Fact]
        public void EveryOfficeInteropReference_IsEmbedded()
        {
            var office = Project().Descendants(Ns + "Reference")
                .Where(r => ((string)r.Attribute("Include")).StartsWith("Microsoft.Office.Interop.") ||
                            ((string)r.Attribute("Include")).StartsWith("Office,"))
                .ToList();

            Assert.Contains(office, r => ((string)r.Attribute("Include")).StartsWith("Microsoft.Office.Interop.Word,"));
            Assert.All(office, r => Assert.Equal("true", (string)r.Element(Ns + "EmbedInteropTypes")));
        }

        // Microsoft.Web.WebView2.Core.dll cannot start without the native
        // WebView2Loader.dll. Only the package's build rules copy it to the
        // output; a project that does not import them ships without it and
        // the chat pane never opens.
        [Fact]
        public void ImportsTheWebView2BuildRules_SoTheNativeLoaderShips()
        {
            var version = (string)XDocument.Load(Path.Combine(ProjectDir, "packages.config")).Root.Elements("package")
                .Single(p => (string)p.Attribute("id") == "Microsoft.Web.WebView2").Attribute("version");
            var targets = Path.Combine("..", "packages", "Microsoft.Web.WebView2." + version, "build", "Microsoft.Web.WebView2.targets");

            var imports = Project().Descendants(Ns + "Import").Select(i => (string)i.Attribute("Project")).ToList();

            Assert.Contains(targets, imports);
            Assert.True(File.Exists(Path.GetFullPath(Path.Combine(ProjectDir, targets))), targets + " does not exist (restore the packages first)");
        }

        // One packages folder for the whole solution: a second one makes a clean clone unbuildable.
        [Fact]
        public void EveryHintPath_UsesTheSolutionPackagesFolder()
        {
            var hintPaths = Project().Descendants(Ns + "HintPath").Select(h => h.Value).ToList();

            Assert.NotEmpty(hintPaths);
            Assert.All(hintPaths, h => Assert.StartsWith(Path.Combine("..", "packages") + Path.DirectorySeparatorChar, h));
        }
    }
}
