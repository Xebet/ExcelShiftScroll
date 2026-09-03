using System;
using ExcelDna.Integration;

namespace ExcelShiftScroll.Excel;

internal sealed class ExcelWindowScroller : IExcelScroller
{
    public void ScrollColumns(int columnDelta)
    {
        if (columnDelta == 0)
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
        var toLeft = columnDelta < 0 ? (object)-columnDelta : Type.Missing;
        window.SmallScroll(Type.Missing, Type.Missing, toRight, toLeft);
    }
}
