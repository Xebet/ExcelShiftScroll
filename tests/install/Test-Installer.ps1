[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
# Load before defining mocks; Windows PowerShell's module auto-import would
# otherwise replace the Get-FileHash test function on its first real hash call.
Import-Module Microsoft.PowerShell.Utility
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ExcelShiftScroll.InstallTests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$package = Join-Path $testRoot 'package'
$destinationDirectory = Join-Path $testRoot 'installed'
New-Item -ItemType Directory -Path $package | Out-Null
$installer = Join-Path $package 'Install-CurrentUser.ps1'
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'install\Install-CurrentUser.ps1') -Destination $installer
$source = Join-Path $package 'ExcelShiftScroll64.xll'
$destination = Join-Path $destinationDirectory 'ExcelShiftScroll64.xll'
$checksum = Join-Path $package 'SHA256SUMS.txt'
$testState = @{ FailVerification = $false; ExcelRunning = $false; ProcessChecks = 0; StartExcelOnSecondCheck = $false; Passed = 0 }

# These mocks affect only this test process and the copied installer. All file
# operations use the isolated temporary tree, never the user's installed add-in.
function Get-Process {
    param($Name, $ErrorAction)
    $testState.ProcessChecks++
    if ($testState.ExcelRunning -or ($testState.StartExcelOnSecondCheck -and $testState.ProcessChecks -ge 2)) {
        [pscustomobject]@{ Id = -1 }
    }
}
function Get-FileHash {
    param($LiteralPath, $Algorithm)
    if ($testState.FailVerification -and $LiteralPath -eq $destination) {
        [pscustomobject]@{ Hash = ('0' * 64) }
    } else { Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $LiteralPath -Algorithm $Algorithm }
}
function Set-Package([string]$Contents) {
    [IO.File]::WriteAllText($source, $Contents)
    $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    [IO.File]::WriteAllText($checksum, "$hash *ExcelShiftScroll64.xll")
}
function Assert-True($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Expect-Failure([scriptblock]$Action) {
    $failed = $false
    try { & $Action | Out-Null } catch { $failed = $true; Write-Output "Expected failure: $($_.Exception.Message)" }
    Assert-True $failed 'Expected installation to fail.'
}
function Passed([string]$Name) { $testState.Passed++; Write-Output "PASS $Name" }
try {
    Set-Package 'version-one'
    & $installer -InstallDirectory $destinationDirectory | Out-Null
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-one') 'Fresh install mismatch.'
    Passed 'fresh installation'

    Set-Package 'version-two'
    & $installer -InstallDirectory $destinationDirectory | Out-Null
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-two') 'Upgrade mismatch.'
    $backups = @(Get-ChildItem -LiteralPath $destinationDirectory -Filter '*.bak')
    Assert-True ($backups.Count -eq 1 -and [IO.File]::ReadAllText($backups[0].FullName) -eq 'version-one') 'Backup mismatch.'
    Passed 'upgrade retains previous XLL'

    [IO.File]::WriteAllText($source, 'corrupt')
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-two') 'Corrupt package changed install.'
    Passed 'checksum mismatch leaves installed file unchanged'

    Remove-Item -LiteralPath $checksum
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    Passed 'missing checksum rejected'
    Set-Package 'version-three'

    $lock = [IO.File]::Open($destination, 'Open', 'Read', 'Read')
    try { Expect-Failure { & $installer -InstallDirectory $destinationDirectory } }
    finally { $lock.Dispose() }
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-two') 'Locked file changed.'
    Passed 'locked XLL preserved'

    $lock = [IO.File]::Open((Join-Path $destinationDirectory '.install.lock'), 'Open', 'ReadWrite', 'None')
    try { Expect-Failure { & $installer -InstallDirectory $destinationDirectory } }
    finally { $lock.Dispose() }
    Passed 'concurrent installer rejected'

    $testState.ExcelRunning = $true
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    $testState.ExcelRunning = $false
    Passed 'running Excel rejected'

    $testState.ProcessChecks = 0
    $testState.StartExcelOnSecondCheck = $true
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    $testState.StartExcelOnSecondCheck = $false
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-two') 'Excel recheck changed install.'
    Passed 'Excel start during staging rejected'

    $testState.FailVerification = $true
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    $testState.FailVerification = $false
    Assert-True ([IO.File]::ReadAllText($destination) -eq 'version-two') 'Rollback failed.'
    Passed 'post-commit verification failure restores previous XLL'

    Remove-Item -LiteralPath $destination
    $testState.FailVerification = $true
    Expect-Failure { & $installer -InstallDirectory $destinationDirectory }
    $testState.FailVerification = $false
    Assert-True (-not (Test-Path -LiteralPath $destination)) 'Failed fresh install was not rolled back.'
    Assert-True (@(Get-ChildItem -LiteralPath $destinationDirectory -Filter '*.tmp').Count -eq 0) 'Temporary files remain.'
    Passed 'failed fresh installation restores absent state and cleans staging'
    Write-Output "$($testState.Passed) installer tests passed."
} finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\ExcelShiftScroll.InstallTests-'
    if (-not $resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
