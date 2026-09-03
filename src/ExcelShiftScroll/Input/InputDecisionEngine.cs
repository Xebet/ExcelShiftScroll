using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.Input;

public sealed class InputDecisionEngine
{
    internal const int WheelDeltaPerDetent = 120;
    private readonly object _gate = new();
    private int _accumulatedDelta;

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

            // Vertical wheel-up is positive. The default mapping is left, hence
            // the negative column delta. A partial high-resolution delta is
            // consumed and retained until it reaches one complete detent.
            var columnDelta = -detents * settings.ColumnsPerDetent;
            if (settings.ReverseDirection)
            {
                columnDelta = -columnDelta;
            }

            return ScrollDecision.Consume(columnDelta);
        }
    }
}
