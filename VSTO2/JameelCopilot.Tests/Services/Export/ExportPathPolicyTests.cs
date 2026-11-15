using System;
using System.IO;
using JameelCopilot.Services.Export;
using Xunit;

namespace JameelCopilot.Tests.Services.Export
{
    public class ExportPathPolicyTests
    {
        private static (ExportPathPolicy policy, string baseDir) CreateInSandbox()
        {
            var baseDir = Path.Combine(Path.GetTempPath(), "JameelCopilotTest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(baseDir);
            return (new ExportPathPolicy(new ExportPathResolver(baseDirOverride: baseDir)), baseDir);
        }

        // "Open" hands the file to its default program, so only the two kinds
        // of file the add-in itself writes may be opened.
        [Theory]
        [InlineData("report.exe")]
        [InlineData("report.lnk")]
        [InlineData("report.xlsm")]
        [InlineData("report.pdf.bat")]
        [InlineData("report.xlsx.exe")]
        [InlineData("report.html")]
        [InlineData("report")]
        public void RejectsFilesThatAreNotReports(string name)
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(Path.Combine(baseDir, name)));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Theory]
        [InlineData("Report.XLSX")]
        [InlineData("Report.Pdf")]
        public void AcceptsReportExtensionsInAnyCase(string name)
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                policy.RequireInsideReportsDir(Path.Combine(baseDir, name));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void AcceptsCanonicalPathInsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var ok = Path.Combine(baseDir, "Report.xlsx");
                File.WriteAllText(ok, "");

                policy.RequireInsideReportsDir(ok);
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void AcceptsPathCaseInsensitive()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var weirdCase = baseDir.ToUpperInvariant() + Path.DirectorySeparatorChar + "Report.xlsx";
                File.WriteAllText(Path.Combine(baseDir, "Report.xlsx"), "");

                policy.RequireInsideReportsDir(weirdCase);
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsTraversalAttempt()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var evil = Path.Combine(baseDir, "..", "..", "Windows", "System32", "cmd.exe");

                Assert.Throws<UnauthorizedExportPathException>(() => policy.RequireInsideReportsDir(evil));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsPathOutsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"C:\Windows\System32\cmd.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsUncPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"\\server\share\evil.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsNullPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(null));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsEmptyPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(""));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsRelativePath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir("evil.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }
    }
}
