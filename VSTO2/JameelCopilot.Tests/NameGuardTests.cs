using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests
{
    /// <summary>
    /// The product is called Jameel Copilot. The name of the project it was forked from
    /// appears nowhere in the repo, and nothing a user sees says "AI Assistant".
    /// </summary>
    public class NameGuardTests
    {
        private static readonly string Root =
            Path.GetDirectoryName(RepoFiles.Find("VSTO2", "JameelCopilot", "JameelCopilot.csproj"));

        // Built in two pieces so that this file does not contain the name it forbids.
        private static readonly string OldName = "Outlook" + "AI";

        private static readonly string[] SkippedDirs = { ".git", "bin", "obj", "packages", "node_modules", "out", ".superpowers" };

        private static IEnumerable<string> TextFiles(string dir)
        {
            foreach (var file in Directory.GetFiles(dir))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".png" || ext == ".ico" || ext == ".dll" || ext == ".pfx" || ext == ".msi") continue;
                yield return file;
            }

            foreach (var sub in Directory.GetDirectories(dir))
            {
                if (SkippedDirs.Contains(Path.GetFileName(sub), StringComparer.OrdinalIgnoreCase)) continue;
                foreach (var file in TextFiles(sub)) yield return file;
            }
        }

        [Fact]
        public void NoOldNameAnywhereInTheRepo()
        {
            var repo = Path.GetDirectoryName(Path.GetDirectoryName(Root));
            var offenders = TextFiles(repo)
                .Where(f => f.Substring(repo.Length + 1).IndexOf(OldName, StringComparison.OrdinalIgnoreCase) >= 0
                         || File.ReadAllText(f).IndexOf(OldName, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(f => f.Substring(repo.Length + 1))
                .ToList();

            Assert.True(offenders.Count == 0, "The old name is still used: " + string.Join(", ", offenders));
        }

        // The product was called "H" + "B" before Jameel Copilot; built in pieces so this file passes.
        private static readonly System.Text.RegularExpressions.Regex PreviousName = new System.Text.RegularExpressions.Regex(
            @"\b" + "H" + @"B\b|" + "H" + "BOutlook|" + "hb" + "outlook|" + @"\bh" + @"b[_.-]",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        [Fact]
        public void NoPreviousNameAnywhereInTheRepo()
        {
            var repo = Path.GetDirectoryName(Path.GetDirectoryName(Root));
            var offenders = TextFiles(repo)
                .Where(f => PreviousName.IsMatch(f.Substring(repo.Length + 1)) || PreviousName.IsMatch(File.ReadAllText(f)))
                .Select(f => f.Substring(repo.Length + 1))
                .ToList();

            Assert.True(offenders.Count == 0, "The previous name is still used: " + string.Join(", ", offenders));
        }

        [Fact]
        public void NoOldLabelsInPagesOrRibbon()
        {
            var offenders = Directory.GetFiles(Root, "*.*", SearchOption.AllDirectories)
                .Where(f => new[] { ".html", ".xml", ".css" }.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                .Where(f => !f.Substring(Root.Length + 1).StartsWith("bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                         && !f.Substring(Root.Length + 1).StartsWith("obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                .Where(f => { var t = File.ReadAllText(f); return t.Contains("AI Assistant") || t.Contains("AI Writing"); })
                .Select(f => f.Substring(Root.Length + 1))
                .ToList();

            Assert.True(offenders.Count == 0, "The old labels are still shown: " + string.Join(", ", offenders));
        }

        [Fact]
        public void AssemblyAndAddInAreNamedJameelCopilot()
        {
            var project = File.ReadAllText(Path.Combine(Root, "JameelCopilot.csproj"));

            Assert.Contains("<AssemblyName>JameelCopilot</AssemblyName>", project);
            Assert.Contains("<ProductName>JameelCopilot</ProductName>", project);
            Assert.Contains("<FriendlyName>JameelCopilot</FriendlyName>", project);
            Assert.Contains("<RootNamespace>JameelCopilot</RootNamespace>", project);
        }

        // The build stamps both from one number (1.0.0 unless the MSI build passes another).
        [Fact]
        public void FileVersionIsTheAssemblyVersion()
        {
            var assembly = typeof(Config).Assembly;

            Assert.Equal("JameelCopilot", assembly.GetName().Name);
            Assert.Equal(
                assembly.GetName().Version.ToString(),
                System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
        }
    }
}
