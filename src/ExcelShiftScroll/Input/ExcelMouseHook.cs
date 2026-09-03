using System;
using System.Runtime.InteropServices;
using ExcelShiftScroll.Excel;
using ExcelShiftScroll.Interop;
using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.Input;

internal sealed class ExcelMouseHook : IDisposable
{
    private readonly InputDecisionEngine _decisionEngine;
    private readonly SettingsManager _settings;
    private readonly IWorksheetRegionDetector _regionDetector;
    private readonly ScrollDispatcher _dispatcher;
    private readonly NativeMethods.HookProc _callback;
    private readonly HookLifetime _lifetime;

    internal ExcelMouseHook(
        InputDecisionEngine decisionEngine,
        SettingsManager settings,
        IWorksheetRegionDetector regionDetector,
        ScrollDispatcher dispatcher)
    {
        _decisionEngine = decisionEngine ?? throw new ArgumentNullException(nameof(decisionEngine));
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _regionDetector = regionDetector ?? throw new ArgumentNullException(nameof(regionDetector));
        _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        _callback = HookCallback; // Strong reference prevents delegate collection.
        _lifetime = new HookLifetime(InstallNativeHook, NativeMethods.UnhookWindowsHookEx);
    }

    internal void Install() => _lifetime.Install();

    public void Dispose() => _lifetime.Dispose();

    private IntPtr InstallNativeHook() => NativeMethods.SetWindowsHookEx(
        NativeMethods.WhMouse,
        _callback,
        IntPtr.Zero,
        NativeMethods.GetCurrentThreadId());

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code < 0)
        {
            return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        try
        {
            var message = unchecked((int)wParam.ToInt64());
            if (message != NativeMethods.WmMouseWheel)
            {
                // Native WM_MOUSEHWHEEL and every unrelated mouse message pass through.
                return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
            }

            var data = Marshal.PtrToStructure<NativeMethods.MouseHookStructEx>(lParam);
            var delta = unchecked((short)(data.MouseData >> 16));
            var snapshot = new InputSnapshot(
                isVerticalWheel: true,
                isHorizontalWheel: false,
                wheelDelta: delta,
                modifiers: ReadModifiers(),
                isExcelForeground: _regionDetector.IsCurrentExcelForeground(),
                isWorksheetArea: _regionDetector.IsWorksheetArea(data.WindowHandle, data.Point));

            var decision = _decisionEngine.Decide(snapshot, _settings.Current);
            if (!decision.Handled)
            {
                return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
            }

            if (decision.ColumnDelta != 0 && !_dispatcher.Enqueue(decision.ColumnDelta))
            {
                return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
            }

            return new IntPtr(1);
        }
        catch
        {
            // Never allow an input callback failure to escape into Excel.
            // Deliberately avoid file logging inside the hook callback.
            return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }
    }

    private static ModifierKeys ReadModifiers()
    {
        var modifiers = ModifierKeys.None;
        if (IsKeyDown(NativeMethods.VkShift))
        {
            modifiers |= ModifierKeys.Shift;
        }

        if (IsKeyDown(NativeMethods.VkControl))
        {
            modifiers |= ModifierKeys.Control;
        }

        if (IsKeyDown(NativeMethods.VkMenu))
        {
            modifiers |= ModifierKeys.Alt;
        }

        if (IsKeyDown(NativeMethods.VkLWin) || IsKeyDown(NativeMethods.VkRWin))
        {
            modifiers |= ModifierKeys.Windows;
        }

        return modifiers;
    }

    private static bool IsKeyDown(int virtualKey) =>
        (NativeMethods.GetKeyState(virtualKey) & unchecked((short)0x8000)) != 0;
}
