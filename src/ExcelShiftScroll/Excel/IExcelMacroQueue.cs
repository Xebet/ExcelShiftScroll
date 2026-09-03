using System;

namespace ExcelShiftScroll.Excel;

public interface IExcelMacroQueue
{
    void Queue(Action action);
}
