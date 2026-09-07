using System;
using ExcelShiftScroll.Interop;

namespace ExcelShiftScroll.Excel;

/// <summary>
/// Converts eligible vertical-wheel input into native horizontal-wheel messages.
/// Posting rather than sending avoids re-entering Excel from the mouse hook.
/// </summary>
internal sealed class NativeHorizontalWheelDispatcher
{
    private readonly Func<uint?> _getNativeColumnsPerDetent;
    private readonly Func<IntPtr, uint, IntPtr, IntPtr, bool> _postMessage;
    private readonly object _gate = new();
    private long _scaleRemainder;
    private int _lastRequestedColumns;
    private uint _lastNativeColumns;
    private IntPtr _lastTarget;

    internal void Reset()
    {
        lock (_gate) { _scaleRemainder = 0; }
    }

    internal NativeHorizontalWheelDispatcher()
        : this(GetNativeColumnsPerDetent, NativeMethods.PostMessage)
    {
    }

    internal NativeHorizontalWheelDispatcher(
        Func<uint?> getNativeColumnsPerDetent,
        Func<IntPtr, uint, IntPtr, IntPtr, bool> postMessage)
    {
        _getNativeColumnsPerDetent = getNativeColumnsPerDetent ??
            throw new ArgumentNullException(nameof(getNativeColumnsPerDetent));
        _postMessage = postMessage ?? throw new ArgumentNullException(nameof(postMessage));
    }

    internal bool TryPost(
        IntPtr targetWindow,
        NativeMethods.Point screenPoint,
        int horizontalWheelDelta,
        int requestedColumnsPerDetent)
    {
        if (targetWindow == IntPtr.Zero || horizontalWheelDelta == 0)
        {
            Reset();
            return false;
        }

        var nativeColumns = _getNativeColumnsPerDetent();
        if (!nativeColumns.HasValue ||
            nativeColumns.Value == 0 ||
            nativeColumns.Value == NativeMethods.WheelPageScroll ||
            requestedColumnsPerDetent <= 0)
        {
            Reset();
            return false;
        }

        if (_lastTarget != targetWindow)
        {
            Reset();
            _lastTarget = targetWindow;
        }

        var scaledDelta = ScaleDelta(
            horizontalWheelDelta,
            requestedColumnsPerDetent,
            nativeColumns.Value);
        if (scaledDelta == 0)
        {
            // A sub-unit high-resolution delta is retained for the next event.
            return true;
        }

        var pointParameter = PackPoint(screenPoint);
        var postedAny = false;
        while (scaledDelta != 0)
        {
            short chunk;
            if (scaledDelta > short.MaxValue)
            {
                chunk = short.MaxValue;
            }
            else if (scaledDelta < short.MinValue)
            {
                chunk = short.MinValue;
            }
            else
            {
                chunk = (short)scaledDelta;
            }
            var wheelParameter = PackWheelParameter(chunk);
            if (!_postMessage(
                    targetWindow,
                    NativeMethods.WmMouseHWheel,
                    wheelParameter,
                    pointParameter))
            {
                Reset();
                return postedAny;
            }

            postedAny = true;
            scaledDelta -= chunk;
        }

        return true;
    }

    private int ScaleDelta(int delta, int requestedColumns, uint nativeColumns)
    {
        lock (_gate)
        {
            if (_lastRequestedColumns != requestedColumns || _lastNativeColumns != nativeColumns)
            {
                _scaleRemainder = 0;
                _lastRequestedColumns = requestedColumns;
                _lastNativeColumns = nativeColumns;
            }

            var numerator = _scaleRemainder + ((long)delta * requestedColumns);
            var scaled = numerator / nativeColumns;
            _scaleRemainder = numerator % nativeColumns;
            return scaled > int.MaxValue
                ? int.MaxValue
                : scaled < int.MinValue
                    ? int.MinValue
                    : (int)scaled;
        }
    }

    private static uint? GetNativeColumnsPerDetent() =>
        NativeMethods.SystemParametersInfo(
            NativeMethods.SpiGetWheelScrollChars,
            0,
            out var value,
            0)
            ? value
            : null;

    private static IntPtr PackWheelParameter(short delta)
    {
        // Shift identifies the source gesture, but it must not be forwarded in
        // the horizontal message. Excel gives Shift + WM_MOUSEHWHEEL a special
        // behavior instead of using its ordinary native horizontal scroll path.
        var packed = (uint)(ushort)delta << 16;
        return new IntPtr(unchecked((int)packed));
    }

    private static IntPtr PackPoint(NativeMethods.Point point)
    {
        var packed = (uint)(ushort)point.X | ((uint)(ushort)point.Y << 16);
        return new IntPtr(unchecked((int)packed));
    }
}
