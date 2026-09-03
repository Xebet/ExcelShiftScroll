using System;

namespace ExcelShiftScroll.Diagnostics;

public interface IDiagnosticLog
{
    void Write(string eventName, Exception? exception = null);
}
