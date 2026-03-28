using System.IO;
using System.Text;

namespace AdvancedRdp.Services;

public sealed class DebugLogService
{
    private readonly object _syncRoot = new();

    public string LogFilePath { get; }

    public DebugLogService(string scope)
    {
        var appDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AdvancedRdp",
            "logs");
        Directory.CreateDirectory(appDir);

        var safeScope = string.Concat(scope.Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch));
        LogFilePath = Path.Combine(appDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{safeScope}.log");
    }

    public void Write(string source, string message)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{source}] {message}{Environment.NewLine}";
        lock (_syncRoot)
        {
            File.AppendAllText(LogFilePath, line, Encoding.UTF8);
        }
    }
}
