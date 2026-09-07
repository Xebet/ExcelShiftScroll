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
    private readonly NativeHorizontalWheelDispatcher _nativeDispatcher;
    private readonly ScrollDispatcher _dispatcher;
    private readonly NativeMethods.HookProc _callback;
    private readonly HookLifetime _lifetime;
    private IntPtr _lastTarget;

    internal ExcelMouseHook(
        InputDecisionEngine decisionEngine,
        SettingsManager settings,
        IWorksheetRegionDetector regionDetector,
        NativeHorizontalWheelDispatcher nativeDispatcher,
        ScrollDispatcher dispatcher)
    {
        _decisionEngine = decisionEngine ?? throw new ArgumentNullException(nameof(decisionEngine));
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _regionDetector = regionDetector ?? throw new ArgumentNullException(nameof(regionDetector));
        _nativeDispatcher = nativeDispatcher ?? throw new ArgumentNullException(nameof(nativeDispatcher));
        _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        _callback = HookCallback; // Strong reference prevents delegate collection.
        _lifetime = new HookLifetime(InstallNativeHook, NativeMethods.UnhookWindowsHookEx);
        _settings.Changed += OnSettingsChanged;
    }

    internal void Install() => _lifetime.Install();

    public void Dispose()
    {
        _settings.Changed -= OnSettingsChanged;
        _lifetime.Dispose();
    }

    private void OnSettingsChanged(object? sender, EventArgs args)
    {
        _decisionEngine.Reset();
        _nativeDispatcher.Reset();
        _dispatcher.CancelPending();
    }

    internal static bool ShouldProcess(int code, IntPtr message) =>
        code == 0 && message == new IntPtr(NativeMethods.WmMouseWheel);

    private IntPtr InstallNativeHook() => NativeMethods.SetWindowsHookEx(
        NativeMethods.WhMouse,
        _callback,
        IntPtr.Zero,
        NativeMethods.GetCurrentThreadId());

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        // HC_NOREMOVE is a peek, not a consumed input event. Never replay it.
        if (!ShouldProcess(code, wParam))
        {
            return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
        }

        try
        {
            var data = Marshal.PtrToStructure<NativeMethods.MouseHookStructEx>(lParam);
            var delta = unchecked((short)(data.MouseData >> 16));
            var snapshot = new InputSnapshot(
                isVerticalWheel: true,
                isHorizontalWheel: false,
                wheelDelta: delta,
                modifiers: ReadModifiers(),
                isExcelForeground: _regionDetector.IsCurrentExcelForeground(),
                isWorksheetArea: _regionDetector.IsWorksheetArea(data.WindowHandle, data.Point));

            var targetWindow = NativeMethods.WindowFromPoint(data.Point);
            if (targetWindow == IntPtr.Zero) { targetWindow = data.WindowHandle; }
            if (_lastTarget != targetWindow)
            {
                _decisionEngine.Reset();
                _dispatcher.CancelPending();
                _lastTarget = targetWindow;
            }

            var settings = _settings.Current;
            var decision = _decisionEngine.Decide(snapshot, settings);
            if (!decision.Handled)
            {
                _nativeDispatcher.Reset();
                _dispatcher.CancelPending();
                return NativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
            }

            var nativePosted = _nativeDispatcher.TryPost(
                targetWindow,
                data.Point,
                decision.HorizontalWheelDelta,
                settings.ColumnsPerDetent);
            if (nativePosted)
            {
                // The native path already owns this delta, including fractions.
                _decisionEngine.Reset();
                _dispatcher.CancelPending();
            }
            if (!nativePosted &&
                decision.ColumnDelta != 0 &&
                !_dispatcher.Enqueue(decision.ColumnDelta,
                    NativeMethods.GetAncestor(targetWindow, NativeMethods.GaRoot)))
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
