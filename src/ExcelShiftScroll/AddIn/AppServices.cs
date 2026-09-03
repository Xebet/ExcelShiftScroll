using System;
using ExcelShiftScroll.Diagnostics;
using ExcelShiftScroll.Excel;
using ExcelShiftScroll.Input;
using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.AddIn;

internal static class AppServices
{
    private static readonly object Gate = new();
    private static ExcelMouseHook? _hook;
    private static ScrollDispatcher? _dispatcher;
    private static SettingsManager? _settings;
    private static IDiagnosticLog _log = NullDiagnosticLog.Instance;
    private static string _status = "Not initialized";

    internal static SettingsManager Settings
    {
        get
        {
            lock (Gate)
            {
                return _settings ??= new SettingsManager(
                    new JsonSettingsStore(JsonSettingsStore.DefaultPath));
            }
        }
    }

    internal static string Status
    {
        get
        {
            lock (Gate)
            {
                return _status;
            }
        }
    }

    internal static void Initialize()
    {
        lock (Gate)
        {
            if (_hook is not null)
            {
                return;
            }

            var settings = Settings;
            _log = new LocalDiagnosticLog(settings);
            var dispatcher = new ScrollDispatcher(
                new ExcelMacroQueue(),
                new ExcelWindowScroller(),
                _log);
            var hook = new ExcelMouseHook(
                new InputDecisionEngine(),
                settings,
                new WorksheetRegionDetector(),
                dispatcher);

            try
            {
                hook.Install();
                _dispatcher = dispatcher;
                _hook = hook;
                _status = "Mouse hook active";
                _log.Write("add_in_started");
            }
            catch (Exception exception)
            {
                hook.Dispose();
                dispatcher.Dispose();
                _status = "Mouse hook unavailable";
                _log.Write("hook_install_failed", exception);
            }
        }
    }

    internal static void Shutdown()
    {
        lock (Gate)
        {
            _hook?.Dispose();
            _hook = null;
            _dispatcher?.Dispose();
            _dispatcher = null;
            _status = "Stopped";
            _log.Write("add_in_stopped");
            _log = NullDiagnosticLog.Instance;
        }
    }
}
