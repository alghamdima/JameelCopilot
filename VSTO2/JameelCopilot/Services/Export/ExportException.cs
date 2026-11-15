using System;

namespace JameelCopilot.Services.Export
{
    public sealed class ExportException : Exception
    {
        public ExportException(string code, string detail, Exception inner = null)
            : base(detail, inner)
        {
            Code = code;
        }

        public string Code { get; }
    }
}
