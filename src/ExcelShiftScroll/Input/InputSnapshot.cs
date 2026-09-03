namespace ExcelShiftScroll.Input;

public readonly struct InputSnapshot
{
    public InputSnapshot(
        bool isVerticalWheel,
        bool isHorizontalWheel,
        int wheelDelta,
        ModifierKeys modifiers,
        bool isExcelForeground,
        bool isWorksheetArea)
    {
        IsVerticalWheel = isVerticalWheel;
        IsHorizontalWheel = isHorizontalWheel;
        WheelDelta = wheelDelta;
        Modifiers = modifiers;
        IsExcelForeground = isExcelForeground;
        IsWorksheetArea = isWorksheetArea;
    }

    public bool IsVerticalWheel { get; }
    public bool IsHorizontalWheel { get; }
    public int WheelDelta { get; }
    public ModifierKeys Modifiers { get; }
    public bool IsExcelForeground { get; }
    public bool IsWorksheetArea { get; }
}
