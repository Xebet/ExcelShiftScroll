[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$XllPath
)

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelShiftScrollNativeTest
{
    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(
        IntPtr parent,
        EnumWindowsProc callback,
        IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(
        IntPtr window,
        StringBuilder className,
        int maximumCount);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr window, out Rect rect);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out Point point);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern void keybd_event(
        byte virtualKey,
        byte scanCode,
        uint flags,
        UIntPtr extraInfo);

    [DllImport("user32.dll")]
    private static extern void mouse_event(
        uint flags,
        uint x,
        uint y,
        uint data,
        UIntPtr extraInfo);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    public static IntPtr FindWorksheet(IntPtr root)
    {
        IntPtr result = IntPtr.Zero;
        long largestArea = 0;
        EnumChildWindows(root, (window, parameter) =>
        {
            var className = new StringBuilder(256);
            if (GetClassName(window, className, className.Capacity) > 0 &&
                string.Equals(className.ToString(), "EXCEL7", StringComparison.OrdinalIgnoreCase) &&
                IsWindowVisible(window) &&
                GetWindowRect(window, out var rect))
            {
                var area = (long)(rect.Right - rect.Left) * (rect.Bottom - rect.Top);
                if (area > largestArea)
                {
                    largestArea = area;
                    result = window;
                }
            }

            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static bool PostHorizontalWheel(IntPtr window, short delta, bool includeShift)
    {
        const uint WmMouseHWheel = 0x020E;
        const uint MkShift = 0x0004;
        if (!GetWindowRect(window, out var rect))
        {
            return false;
        }

        var x = (short)(rect.Left + ((rect.Right - rect.Left) / 2));
        var y = (short)(rect.Top + ((rect.Bottom - rect.Top) / 2));
        var wheel = ((uint)(ushort)delta << 16) | (includeShift ? MkShift : 0);
        var point = (uint)(ushort)x | ((uint)(ushort)y << 16);
        return PostMessage(
            window,
            WmMouseHWheel,
            new IntPtr(unchecked((int)wheel)),
            new IntPtr(unchecked((int)point)));
    }

    public static bool PreparePhysicalInput(
        IntPtr excelRoot,
        IntPtr worksheet,
        out int previousX,
        out int previousY)
    {
        previousX = 0;
        previousY = 0;
        if (!GetCursorPos(out var previous) ||
            !GetWindowRect(worksheet, out var rect))
        {
            return false;
        }

        previousX = previous.X;
        previousY = previous.Y;
        var x = rect.Left + ((rect.Right - rect.Left) / 2);
        var y = rect.Top + ((rect.Bottom - rect.Top) / 2);
        return SetCursorPos(x, y) && SetForegroundWindow(excelRoot);
    }

    public static void SetShift(bool down)
    {
        const byte VkShift = 0x10;
        const uint KeyEventFKeyUp = 0x0002;
        keybd_event(VkShift, 0, down ? 0 : KeyEventFKeyUp, UIntPtr.Zero);
    }

    public static void SendVerticalWheel(int delta)
    {
        const uint MouseEventFWheel = 0x0800;
        mouse_event(MouseEventFWheel, 0, 0, unchecked((uint)delta), UIntPtr.Zero);
    }

    public static void RestoreCursor(int x, int y)
    {
        SetCursorPos(x, y);
    }
}
'@

$resolvedXll = (Resolve-Path -LiteralPath $XllPath).Path
$existingExcelIds = @(
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Id }
)
$excel = $null
$excelProcess = $null
$ownsExcelProcess = $false
$workbooks = $null
$workbook = $null
$window = $null
try {
    $excel = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 250
    $createdProcesses = @(
        Get-Process -Name 'EXCEL' -ErrorAction Stop |
            Where-Object { $existingExcelIds -notcontains $_.Id }
    )
    if ($createdProcesses.Count -ne 1) {
        throw "Excel COM automation did not create exactly one isolated process; refusing to alter an existing session. New process count: $($createdProcesses.Count)."
    }
    $excelProcess = $createdProcesses[0]
    $ownsExcelProcess = $true

    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    $window = $excel.ActiveWindow

    if (-not $excel.RegisterXLL($resolvedXll)) {
        throw "Excel RegisterXLL returned false for $resolvedXll"
    }

    $worksheetWindow = [ExcelShiftScrollNativeTest]::FindWorksheet([IntPtr]$excel.Hwnd)
    if ($worksheetWindow -eq [IntPtr]::Zero) {
        throw 'No EXCEL7 worksheet window was found.'
    }

    $window.ScrollColumn = 10
    Start-Sleep -Milliseconds 750
    $before = [int]$window.ScrollColumn
    if (-not [ExcelShiftScrollNativeTest]::PostHorizontalWheel($worksheetWindow, 120, $false)) {
        throw 'PostMessage failed for rightward WM_MOUSEHWHEEL.'
    }
    Start-Sleep -Milliseconds 750
    $afterRight = [int]$window.ScrollColumn

    $window.ScrollColumn = 10
    Start-Sleep -Milliseconds 750
    if (-not [ExcelShiftScrollNativeTest]::PostHorizontalWheel($worksheetWindow, -120, $false)) {
        throw 'PostMessage failed for leftward WM_MOUSEHWHEEL.'
    }
    Start-Sleep -Milliseconds 750
    $afterLeft = [int]$window.ScrollColumn

    if ($afterRight -le $before) {
        throw "Native rightward message did not advance ScrollColumn: before=$before after=$afterRight"
    }
    if ($afterLeft -ge 10) {
        throw "Native leftward message did not reduce ScrollColumn: before=10 after=$afterLeft"
    }

    Write-Output "Native WM_MOUSEHWHEEL verified: delta=+120 $before->$afterRight; delta=-120 10->$afterLeft; version=$($excel.Evaluate('ExcelShiftScroll.Version()'))"

    $previousX = 0
    $previousY = 0
    if (-not [ExcelShiftScrollNativeTest]::PreparePhysicalInput(
            [IntPtr]$excel.Hwnd,
            $worksheetWindow,
            [ref]$previousX,
            [ref]$previousY)) {
        throw 'Could not focus the isolated Excel worksheet for end-to-end input testing.'
    }

    try {
        $window.ScrollColumn = 10
        Start-Sleep -Milliseconds 750
        [ExcelShiftScrollNativeTest]::SetShift($true)
        [ExcelShiftScrollNativeTest]::SendVerticalWheel(120)
        Start-Sleep -Milliseconds 250
        [ExcelShiftScrollNativeTest]::SetShift($false)
        Start-Sleep -Milliseconds 750
        $afterWheelUp = [int]$window.ScrollColumn

        $window.ScrollColumn = 10
        Start-Sleep -Milliseconds 750
        [ExcelShiftScrollNativeTest]::SetShift($true)
        [ExcelShiftScrollNativeTest]::SendVerticalWheel(-120)
        Start-Sleep -Milliseconds 250
        [ExcelShiftScrollNativeTest]::SetShift($false)
        Start-Sleep -Milliseconds 750
        $afterWheelDown = [int]$window.ScrollColumn
    }
    finally {
        [ExcelShiftScrollNativeTest]::SetShift($false)
        [ExcelShiftScrollNativeTest]::RestoreCursor($previousX, $previousY)
    }

    if ($afterWheelUp -ge 10) {
        throw "End-to-end Shift + wheel-up did not scroll left: before=10 after=$afterWheelUp"
    }
    if ($afterWheelDown -le 10) {
        throw "End-to-end Shift + wheel-down did not scroll right: before=10 after=$afterWheelDown"
    }

    Write-Output "End-to-end gesture verified: Shift+wheel-up 10->$afterWheelUp; Shift+wheel-down 10->$afterWheelDown"
}
finally {
    if ($ownsExcelProcess -and $null -ne $excel) {
        try {
            [void]$excel.UnregisterXLL($resolvedXll)
        }
        catch {
            # Continue teardown even if Excel is already closing.
        }
    }
    if ($null -ne $window) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window)
    }
    if ($ownsExcelProcess -and $null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $workbook) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $workbooks) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks)
    }
    if ($null -ne $excel) {
        if ($ownsExcelProcess) {
            $excel.Quit()
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

if ($null -ne $excelProcess -and -not $excelProcess.WaitForExit(30000)) {
    throw "Excel process $($excelProcess.Id) remained after Quit."
}

$remainingNewProcesses = @(
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Where-Object { $existingExcelIds -notcontains $_.Id }
)
if ($remainingNewProcesses.Count -ne 0) {
    throw 'A newly created Excel process remained after the native horizontal-wheel test.'
}

Write-Output 'The isolated Excel process exited cleanly; pre-existing Excel processes were not modified or closed.'
