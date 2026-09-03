using System;
using System.Globalization;
using System.IO;
using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.Diagnostics;

internal sealed class LocalDiagnosticLog : IDiagnosticLog
{
    private readonly SettingsManager _settings;
    private readonly string _path;
    private readonly object _gate = new();

    internal LocalDiagnosticLog(SettingsManager settings)
    {
        _settings = settings;
        _path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelShiftScroll",
            "diagnostics.jsonl");
    }

    public void Write(string eventName, Exception? exception = null)
    {
        if (!_settings.Current.DiagnosticsEnabled)
        {
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(_path);
            if (directory is null)
            {
                return;
            }

            Directory.CreateDirectory(directory);
            var line = string.Format(
                CultureInfo.InvariantCulture,
                "{{\"timeUtc\":\"{0:O}\",\"event\":\"{1}\",\"exceptionType\":\"{2}\"}}{3}",
                DateTime.UtcNow,
                Escape(eventName),
                Escape(exception?.GetType().FullName ?? string.Empty),
                Environment.NewLine);
            lock (_gate)
            {
                File.AppendAllText(_path, line);
            }
        }
        catch
        {
            // Diagnostics must never affect Excel or input handling.
        }
    }

    private static string Escape(string value) => value
        .Replace("\\", "\\\\")
        .Replace("\"", "\\\"")
        .Replace("\r", "\\r")
        .Replace("\n", "\\n");
}
