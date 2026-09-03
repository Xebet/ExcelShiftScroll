[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$XllPath,

    [Parameter(Mandatory = $false)]
    [string]$ExpectedVersion = '0.2.0'
)

$ErrorActionPreference = 'Stop'
$existingExcelIds = @(
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Id }
)

$resolvedXll = (Resolve-Path -LiteralPath $XllPath).Path
$excel = $null
$excelProcess = $null
$ownsExcelProcess = $false
try {
    $excel = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 250
    $createdProcesses = @(
        Get-Process -Name 'EXCEL' -ErrorAction Stop |
            Where-Object { $existingExcelIds -notcontains $_.Id }
    )
    if ($createdProcesses.Count -ne 1) {
        throw "Excel COM automation did not create exactly one isolated process; refusing to modify or close an existing session. New process count: $($createdProcesses.Count)."
    }
    $excelProcess = $createdProcesses[0]
    $ownsExcelProcess = $true

    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $registered = $excel.RegisterXLL($resolvedXll)
    if (-not $registered) {
        throw "Excel RegisterXLL returned false for $resolvedXll"
    }

    $version = [string]$excel.Evaluate('ExcelShiftScroll.Version()')
    $status = [string]$excel.Evaluate('ExcelShiftScroll.Status()')
    if ($version -ne $ExpectedVersion) {
        throw "Version UDF returned '$version', expected '$ExpectedVersion'."
    }
    if ($status -ne 'Mouse hook active') {
        throw "Status UDF returned '$status', expected 'Mouse hook active'."
    }

    Write-Output "RegisterXLL=true version=$version status=$status"
}
finally {
    if ($null -ne $excel) {
        try {
            if ($ownsExcelProcess) {
                $excel.Quit()
            }
        }
        finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null
        }
    }
}

if ($null -ne $excelProcess -and -not $excelProcess.WaitForExit(30000)) {
    throw "Excel process $($excelProcess.Id) remained after Quit."
}

$remainingNewProcesses = @(
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Where-Object { $existingExcelIds -notcontains $_.Id }
)
if ($remainingNewProcesses.Count -ne 0) {
    throw 'A newly created Excel process remained after the isolated XLL smoke test.'
}

Write-Output 'The isolated Excel process exited cleanly; pre-existing Excel processes were not modified or closed.'
