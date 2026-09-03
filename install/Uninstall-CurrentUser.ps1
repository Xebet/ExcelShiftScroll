[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

if (@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Excel is running. Close every Excel window and try again. / Excel 正在运行，请关闭所有 Excel 窗口后重试。'
}

$localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
$installDirectory = Join-Path $localAppData 'ExcelShiftScroll\AddIn'
$expectedFiles = @(
    (Join-Path $installDirectory 'ExcelShiftScroll64.xll'),
    (Join-Path $installDirectory 'ExcelShiftScroll32.xll')
)

foreach ($file in $expectedFiles) {
    if (Test-Path -LiteralPath $file -PathType Leaf) {
        Remove-Item -LiteralPath $file -Force
        Write-Output "Removed: $file"
    }
}

Write-Output 'Settings were preserved under %LocalAppData%\ExcelShiftScroll\settings.json.'
Write-Output 'Remove the ExcelShiftScroll check mark in Excel Add-ins before running this script.'
Write-Output '设置文件已保留；运行本脚本前应先在 Excel 加载项中取消勾选 ExcelShiftScroll。'
