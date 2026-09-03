using System;
using ExcelShiftScroll.Interop;

namespace ExcelShiftScroll.Input;

internal interface IWorksheetRegionDetector
{
    bool IsCurrentExcelForeground();
    bool IsWorksheetArea(IntPtr hookWindow, NativeMethods.Point screenPoint);
}
