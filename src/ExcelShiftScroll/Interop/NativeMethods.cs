using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelShiftScroll.Interop;

internal static class NativeMethods
{
    internal const int WhMouse = 7;
    internal const int WmMouseWheel = 0x020A;
    internal const int WmMouseHWheel = 0x020E;
    internal const uint GaRoot = 2;
    internal const uint SpiGetWheelScrollChars = 0x006C;
    internal const uint WheelPageScroll = uint.MaxValue;

    internal const int VkShift = 0x10;
    internal const int VkControl = 0x11;
    internal const int VkMenu = 0x12;
    internal const int VkLWin = 0x5B;
    internal const int VkRWin = 0x5C;

    internal delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct Point
    {
        internal Point(int x, int y)
        {
            X = x;
            Y = y;
        }

        internal int X { get; }
        internal int Y { get; }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MouseHookStructEx
    {
        internal Point Point;
        internal IntPtr WindowHandle;
        internal uint HitTestCode;
        internal UIntPtr ExtraInfo;
        internal uint MouseData;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern IntPtr SetWindowsHookEx(
        int hookType,
        HookProc callback,
        IntPtr module,
        uint threadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    internal static extern IntPtr CallNextHookEx(
        IntPtr hook,
        int code,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport("kernel32.dll")]
    internal static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    internal static extern short GetKeyState(int virtualKey);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    internal static extern IntPtr WindowFromPoint(Point point);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetAncestor(IntPtr window, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int GetClassName(IntPtr window, StringBuilder className, int maximumCount);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetParent(IntPtr window);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PostMessage(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SystemParametersInfo(
        uint action,
        uint parameter,
        out uint value,
        uint update);
}
