using System;
using ExcelDna.Integration;

namespace ExcelShiftScroll.Excel;

internal sealed class ExcelMacroQueue : IExcelMacroQueue
{
    public void Queue(Action action) => ExcelAsyncUtil.QueueAsMacro(() => action());
}
