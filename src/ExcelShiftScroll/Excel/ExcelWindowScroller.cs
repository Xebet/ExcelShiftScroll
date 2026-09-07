using System;
using ExcelDna.Integration;
using ExcelShiftScroll.Interop;

namespace ExcelShiftScroll.Excel;

internal sealed class ExcelWindowScroller : IExcelScroller
{
    public void ScrollColumns(int columnDelta, IntPtr targetWindow)
    {
        if (columnDelta == 0 || targetWindow == IntPtr.Zero ||
            NativeMethods.GetAncestor(NativeMethods.GetForegroundWindow(), NativeMethods.GaRoot) != targetWindow)
        {
            return;
        }

        dynamic application = ExcelDnaUtil.Application;
        dynamic window = application.ActiveWindow;
        if (window is null)
        {
            return;
        }

        var toRight = columnDelta > 0 ? (object)columnDelta : Type.Missing;
        // Never redirect delayed input into whichever workbook became active.
        var activeHandle = new IntPtr((int)window.Hwnd);
        if (NativeMethods.GetAncestor(activeHandle, NativeMethods.GaRoot) != targetWindow)
        {
            return;
        }
        var toLeft = columnDelta < 0 ? (object)-columnDelta : Type.Missing;
        window.SmallScroll(Type.Missing, Type.Missing, toRight, toLeft);
    }
}
