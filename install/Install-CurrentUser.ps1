[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) 'ExcelShiftScroll\AddIn')
)

$ErrorActionPreference = 'Stop'

if (@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Excel is running. Close every Excel window and try again. / Excel 正在运行，请关闭所有 Excel 窗口后重试。'
}

$sourceFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'ExcelShiftScroll*.xll' -File)
if ($sourceFiles.Count -ne 1) {
    throw "Expected exactly one ExcelShiftScroll XLL beside this script; found $($sourceFiles.Count)."
}

$source = $sourceFiles[0]
if ($source.Name -notin @('ExcelShiftScroll64.xll', 'ExcelShiftScroll32.xll')) {
    throw 'Unrecognized XLL filename.'
}
$checksumPath = Join-Path $PSScriptRoot 'SHA256SUMS.txt'
$checksum = @(Get-Content -LiteralPath $checksumPath | Where-Object { $_.Trim().Length -gt 0 })
if ($checksum.Count -ne 1 -or $checksum[0] -notmatch "^([0-9a-fA-F]{64}) \*$([regex]::Escape($source.Name))$") {
    throw 'The package checksum file is missing or invalid. Download the complete release ZIP.'
}
$expectedHash = $Matches[1]
if ((Get-FileHash -LiteralPath $source.FullName -Algorithm SHA256).Hash -ne $expectedHash) {
    throw 'Package checksum mismatch. Nothing was installed.'
}
$installDirectory = [IO.Path]::GetFullPath($InstallDirectory)
$destination = Join-Path $installDirectory $source.Name

New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
$installLock = $null
$stage = Join-Path $installDirectory ('.install-' + [guid]::NewGuid().ToString('N') + '.tmp')
$backup = $destination + '.' + (Get-Date -Format 'yyyyMMddHHmmss') + '.' + [guid]::NewGuid().ToString('N') + '.bak'
$committed = $false
$hadPrevious = $false
try {
    # Held across staging, replacement and verification; released even on failure.
    $installLock = [IO.File]::Open((Join-Path $installDirectory '.install.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    Copy-Item -LiteralPath $source.FullName -Destination $stage
    if ((Get-FileHash -LiteralPath $stage -Algorithm SHA256).Hash -ne $expectedHash) {
        throw 'Staged file checksum mismatch.'
    }
    Unblock-File -LiteralPath $stage
    if ($null -ne (Get-Item -LiteralPath $stage -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue)) {
        throw 'Windows still reports Mark-of-the-Web. Installation stopped before replacement.'
    }
    if (@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -ne 0) {
        throw 'Excel started during installation. Close Excel and try again.'
    }
    $hadPrevious = Test-Path -LiteralPath $destination -PathType Leaf
    if ($hadPrevious) {
        [IO.File]::Replace($stage, $destination, $backup)
    } else {
        [IO.File]::Move($stage, $destination)
    }
    $committed = $true
    if ((Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash -ne $expectedHash -or
        $null -ne (Get-Item -LiteralPath $destination -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue)) {
        throw 'Installed file verification failed.'
    }
    if ($hadPrevious) { Write-Output "Previous XLL backup: $backup" }
} catch {
    $installError = $_
    if ($committed) {
        try {
            if ($hadPrevious) { [IO.File]::Replace($backup, $destination, [NullString]::Value) }
            else { Remove-Item -LiteralPath $destination -Force }
            Write-Warning 'Installation failed; the previous state was restored.'
        } catch {
            throw "Installation failed and rollback could not complete. Do not load the XLL. Previous file, if any: $backup. $($_.Exception.Message)"
        }
    }
    throw $installError
} finally {
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue }
    if ($null -ne $installLock) { $installLock.Dispose() }
}

Write-Output "Installed and unblocked: $destination"
Write-Output 'No Trust Center, registry, or Excel add-in setting was changed.'
Write-Output 'Next: Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, then select the installed XLL.'
Write-Output '下一步：Excel > 文件 > 选项 > 加载项 > 管理：Excel 加载项 > 转到 > 浏览，然后选择上面的 XLL。'
