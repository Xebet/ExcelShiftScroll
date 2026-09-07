[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$assembly = Join-Path $repositoryRoot 'src\ExcelShiftScroll\bin\Release\net48\ExcelShiftScroll.dll'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ExcelShiftScroll.SettingsTests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$settingsPath = Join-Path $testRoot 'settings.json'
$workers = @()
try {
    # Framework assembly runs in Windows PowerShell, not PowerShell 7's CLR.
    $command = @'
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.Serialization
[void][Reflection.Assembly]::LoadFrom('__ASSEMBLY__')
$store = New-Object ExcelShiftScroll.Settings.JsonSettingsStore('__SETTINGS__')
for ($i = 0; $i -lt 60; $i++) {
    $settings = New-Object ExcelShiftScroll.Settings.ScrollSettings
    $settings.Enabled = $false
    $settings.ColumnsPerDetent = 5
    $store.Save($settings)
    if ($store.Load().Enabled) { throw 'Invalid/default configuration observed during concurrent writes.' }
}
'@
    $command = $command.Replace('__ASSEMBLY__', $assembly.Replace("'", "''")).Replace('__SETTINGS__', $settingsPath.Replace("'", "''"))
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    foreach ($number in 1..3) {
        $workers += Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList '-NoProfile', '-EncodedCommand', $encoded -WindowStyle Hidden -PassThru
    }
    foreach ($worker in $workers) {
        if (-not $worker.WaitForExit(30000)) { throw 'Settings test worker timed out.' }
        if ($worker.ExitCode -ne 0) { throw "Settings test worker failed: $($worker.ExitCode)" }
    }
    $actual = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
    if ($actual.enabled -ne $false -or $actual.columnsPerDetent -ne 5) { throw 'Final JSON is incorrect.' }
    if (@(Get-ChildItem -LiteralPath $testRoot -Filter '*.tmp').Count -ne 0) { throw 'Staging files remain.' }
    Write-Output 'PASS 3 independent processes, 180 settings saves/reads, valid final JSON, no temporary files.'
} finally {
    foreach ($worker in $workers) {
        if (-not $worker.HasExited) { $worker.Kill(); $worker.WaitForExit() }
        $worker.Dispose()
    }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $prefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\ExcelShiftScroll.SettingsTests-'
    if (-not $resolved.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
