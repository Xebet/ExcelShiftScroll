[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

if (@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Excel is running. Close every Excel window and try again. / Excel 正在运行，请关闭所有 Excel 窗口后重试。'
}

$sourceFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'ExcelShiftScroll*.xll' -File)
if ($sourceFiles.Count -ne 1) {
    throw "Expected exactly one ExcelShiftScroll XLL beside this script; found $($sourceFiles.Count)."
}

$source = $sourceFiles[0]
$localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
$installDirectory = Join-Path $localAppData 'ExcelShiftScroll\AddIn'
$destination = Join-Path $installDirectory $source.Name

New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
Copy-Item -LiteralPath $source.FullName -Destination $destination -Force
Unblock-File -LiteralPath $destination

$zoneStream = Get-Item -LiteralPath $destination -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
if ($null -ne $zoneStream) {
    throw "Windows still reports Mark-of-the-Web on $destination. Right-click the XLL, choose Properties, select Unblock, and retry."
}

Write-Output "Installed and unblocked: $destination"
Write-Output 'No Trust Center, registry, or Excel add-in setting was changed.'
Write-Output 'Next: Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, then select the installed XLL.'
Write-Output '下一步：Excel > 文件 > 选项 > 加载项 > 管理：Excel 加载项 > 转到 > 浏览，然后选择上面的 XLL。'
