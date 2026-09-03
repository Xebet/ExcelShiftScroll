using System;

namespace ExcelShiftScroll.Input;

[Flags]
public enum ModifierKeys
{
    None = 0,
    Shift = 1,
    Control = 2,
    Alt = 4,
    Windows = 8,
}
