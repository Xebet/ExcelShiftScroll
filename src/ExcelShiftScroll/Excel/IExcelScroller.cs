namespace ExcelShiftScroll.Excel;

public interface IExcelScroller
{
    void ScrollColumns(int columnDelta, System.IntPtr targetWindow);
}
