using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.Input;

public sealed class InputDecisionEngine
{
    internal const int WheelDeltaPerDetent = 120;
    private readonly object _gate = new();
    private int _accumulatedDelta;

    public void Reset()
    {
        lock (_gate) { _accumulatedDelta = 0; }
    }

    public ScrollDecision Decide(InputSnapshot input, ScrollSettings settings)
    {
        if (!settings.Enabled ||
            !input.IsVerticalWheel ||
            input.IsHorizontalWheel ||
            input.Modifiers != ModifierKeys.Shift ||
            !input.IsExcelForeground ||
            !input.IsWorksheetArea ||
            input.WheelDelta == 0)
        {
            lock (_gate)
            {
                _accumulatedDelta = 0;
            }
            return ScrollDecision.PassThrough;
        }

        lock (_gate)
        {
            _accumulatedDelta += input.WheelDelta;
            var detents = _accumulatedDelta / WheelDeltaPerDetent;
            _accumulatedDelta %= WheelDeltaPerDetent;

            // Vertical wheel-up is positive. Horizontal wheel-left is negative,
            // so both the exact-column fallback and the native smooth path use
            // the opposite sign by default. The native value deliberately keeps
            // partial high-resolution deltas instead of snapping them to 120.
            var columnDelta = -detents * settings.ColumnsPerDetent;
            var horizontalWheelDelta = -input.WheelDelta;
            if (settings.ReverseDirection)
            {
                columnDelta = -columnDelta;
                horizontalWheelDelta = -horizontalWheelDelta;
            }

            return ScrollDecision.Consume(columnDelta, horizontalWheelDelta);
        }
    }
}
