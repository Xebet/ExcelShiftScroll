using System.Linq;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelShiftScroll.AddIn;

public static class VersionFunction
{
    [ExcelFunction(
        Name = "ExcelShiftScroll.Version",
        Description = "Returns the loaded ExcelShiftScroll add-in version.",
        IsThreadSafe = true)]
    public static string Version()
    {
        var informationalVersion = typeof(VersionFunction).Assembly
            .GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), inherit: false)
            .OfType<AssemblyInformationalVersionAttribute>()
            .Select(attribute => attribute.InformationalVersion)
            .FirstOrDefault();
        return informationalVersion?.Split('+')[0] ?? "unknown";
    }

    [ExcelFunction(
        Name = "ExcelShiftScroll.Status",
        Description = "Returns the in-process ExcelShiftScroll hook status.",
        IsThreadSafe = true)]
    public static string Status() => AppServices.Status;
}
