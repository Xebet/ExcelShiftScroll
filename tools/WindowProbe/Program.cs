using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelShiftScroll.WindowProbe;

internal static class Program
{
    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    private static readonly HashSet<uint> ExcelProcessIds = Process
        .GetProcessesByName("EXCEL")
        .Select(process => unchecked((uint)process.Id))
        .ToHashSet();

    private static int Main()
    {
        Console.WriteLine("ExcelShiftScroll class-only window probe");
        Console.WriteLine("No titles, workbook data, input, or coordinates are read.");

        if (ExcelProcessIds.Count == 0)
        {
            Console.WriteLine("No running Excel process was found.");
            return 2;
        }

        EnumWindows(VisitTopLevel, IntPtr.Zero);
        return 0;
    }

    private static bool VisitTopLevel(IntPtr window, IntPtr parameter)
    {
        var threadId = GetWindowThreadProcessId(window, out var processId);
        if (!ExcelProcessIds.Contains(processId))
        {
            return true;
        }

        PrintWindow(window, processId, threadId, depth: 0);
        EnumChildWindows(window, VisitChild, IntPtr.Zero);
        return true;
    }

    private static bool VisitChild(IntPtr window, IntPtr parameter)
    {
        var threadId = GetWindowThreadProcessId(window, out var processId);
        PrintWindow(window, processId, threadId, depth: 1);
        return true;
    }

    private static void PrintWindow(IntPtr window, uint processId, uint threadId, int depth)
    {
        var className = new StringBuilder(256);
        GetClassName(window, className, className.Capacity);
        Console.WriteLine(
            $"{new string(' ', depth * 2)}hwnd=0x{window.ToInt64():X} pid={processId} tid={threadId} class={className}");
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int maximumCount);
}
