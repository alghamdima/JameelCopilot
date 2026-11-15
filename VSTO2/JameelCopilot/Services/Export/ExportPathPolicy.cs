using System;
using System.IO;

namespace JameelCopilot.Services.Export
{
    public sealed class ExportPathPolicy : IExportPathPolicy
    {
        private readonly ExportPathResolver _resolver;

        public ExportPathPolicy(ExportPathResolver resolver)
        {
            _resolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
        }

        public void RequireInsideReportsDir(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new UnauthorizedExportPathException("Path is null or empty.");
            }

            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                throw new UnauthorizedExportPathException("UNC paths are not permitted.");
            }

            if (!Path.IsPathRooted(path))
            {
                throw new UnauthorizedExportPathException("Path is not absolute.");
            }

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(path);
            }
            catch (Exception ex)
            {
                throw new UnauthorizedExportPathException("Path could not be normalized: " + ex.Message);
            }

            var baseFull = Path.GetFullPath(_resolver.ResolveBaseDir());
            if (!baseFull.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                baseFull += Path.DirectorySeparatorChar;
            }

            if (!fullPath.StartsWith(baseFull, StringComparison.OrdinalIgnoreCase))
            {
                throw new UnauthorizedExportPathException(
                    "Path '" + fullPath + "' is not inside the Reports directory '" + baseFull + "'.");
            }

            // The caller hands the file to its default program, so only the two
            // kinds of file the add-in itself writes are let through.
            var extension = Path.GetExtension(fullPath);
            if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                throw new UnauthorizedExportPathException("Only .xlsx and .pdf reports can be opened.");
            }
        }
    }
}
