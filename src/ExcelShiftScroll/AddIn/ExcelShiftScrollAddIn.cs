using ExcelDna.Integration;

namespace ExcelShiftScroll.AddIn;

public sealed class ExcelShiftScrollAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        AppServices.Initialize();
    }

    public void AutoClose() => AppServices.Shutdown();
}
