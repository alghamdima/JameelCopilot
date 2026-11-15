namespace JameelCopilot.Services.Export
{
    public interface IExportPathPolicy
    {
        void RequireInsideReportsDir(string path);
    }
}
