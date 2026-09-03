using System;

namespace ExcelShiftScroll.Diagnostics;

public sealed class NullDiagnosticLog : IDiagnosticLog
{
    public static NullDiagnosticLog Instance { get; } = new();
    private NullDiagnosticLog() { }
    public void Write(string eventName, Exception? exception = null) { }
}
