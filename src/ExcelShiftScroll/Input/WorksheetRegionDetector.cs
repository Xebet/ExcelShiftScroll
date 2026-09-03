using System;
using System.Diagnostics;
using System.Text;
using ExcelShiftScroll.Interop;

namespace ExcelShiftScroll.Input;

internal sealed class WorksheetRegionDetector : IWorksheetRegionDetector
{
    private const string WorksheetClassName = "EXCEL7";
    private readonly uint _processId;

    internal WorksheetRegionDetector()
    {
        using var process = Process.GetCurrentProcess();
        _processId = unchecked((uint)process.Id);
    }

    public bool IsCurrentExcelForeground()
    {
        var foreground = NativeMethods.GetForegroundWindow();
        return foreground != IntPtr.Zero && IsCurrentProcessWindow(foreground);
    }

    public bool IsWorksheetArea(IntPtr hookWindow, NativeMethods.Point screenPoint)
    {
        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero || !IsCurrentProcessWindow(foreground))
        {
            return false;
        }

        var hitWindow = NativeMethods.WindowFromPoint(screenPoint);
        if (hitWindow == IntPtr.Zero)
        {
            hitWindow = hookWindow;
        }

        if (hitWindow == IntPtr.Zero || !IsCurrentProcessWindow(hitWindow))
        {
            return false;
        }

        var hitRoot = NativeMethods.GetAncestor(hitWindow, NativeMethods.GaRoot);
        var foregroundRoot = NativeMethods.GetAncestor(foreground, NativeMethods.GaRoot);
        if (hitRoot == IntPtr.Zero || foregroundRoot == IntPtr.Zero || hitRoot != foregroundRoot)
        {
            return false;
        }

        // Fail closed: the worksheet rendering window is currently EXCEL7.
        // We verify the full ancestry rather than trusting only the hook target,
        // and reject unknown future classes instead of intercepting other UI.
        for (var current = hitWindow; current != IntPtr.Zero; current = NativeMethods.GetParent(current))
        {
            if (string.Equals(GetClassName(current), WorksheetClassName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (current == hitRoot)
            {
                break;
            }
        }

        return false;
    }

    private bool IsCurrentProcessWindow(IntPtr window)
    {
        NativeMethods.GetWindowThreadProcessId(window, out var processId);
        return processId == _processId;
    }

    private static string GetClassName(IntPtr window)
    {
        var buffer = new StringBuilder(256);
        return NativeMethods.GetClassName(window, buffer, buffer.Capacity) > 0
            ? buffer.ToString()
            : string.Empty;
    }
}
