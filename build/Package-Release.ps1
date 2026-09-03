[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$')]
    [string]$Version = '0.2.0'
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$publishDirectory = Join-Path $repositoryRoot 'src\ExcelShiftScroll\bin\Release\net48\publish'
$releaseDirectory = Join-Path $repositoryRoot 'artifacts\release'

if (-not (Test-Path -LiteralPath $publishDirectory -PathType Container)) {
    throw "Packed Excel-DNA output was not found: $publishDirectory"
}

if (Test-Path -LiteralPath $releaseDirectory) {
    $resolvedRelease = [System.IO.Path]::GetFullPath($releaseDirectory)
    $expectedPrefix = [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot 'artifacts'))
    if (-not $resolvedRelease.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean an output directory outside artifacts: $resolvedRelease"
    }

    Remove-Item -LiteralPath $resolvedRelease -Recurse -Force
}

New-Item -ItemType Directory -Path $releaseDirectory | Out-Null

$packages = @(
    @{
        Architecture = 'x64'
        Source = 'ExcelShiftScroll-AddIn64-packed.xll'
        Destination = 'ExcelShiftScroll64.xll'
    },
    @{
        Architecture = 'x86'
        Source = 'ExcelShiftScroll-AddIn-packed.xll'
        Destination = 'ExcelShiftScroll32.xll'
    }
)

foreach ($package in $packages) {
    $source = Join-Path $publishDirectory $package.Source
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Expected $($package.Architecture) add-in was not found: $source"
    }

    $baseName = "ExcelShiftScroll-v$Version-$($package.Architecture)"
    $stageDirectory = Join-Path $releaseDirectory $baseName
    New-Item -ItemType Directory -Path $stageDirectory | Out-Null

    $destination = Join-Path $stageDirectory $package.Destination
    Copy-Item -LiteralPath $source -Destination $destination
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'README.md') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'README.zh-CN.md') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'INSTALL.txt') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'LICENSE') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'THIRD_PARTY_NOTICES.md') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'install\Install-CurrentUser.ps1') -Destination $stageDirectory
    Copy-Item -LiteralPath (Join-Path $repositoryRoot 'install\Uninstall-CurrentUser.ps1') -Destination $stageDirectory

    $xllHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
    "$xllHash *$($package.Destination)" | Set-Content -LiteralPath (Join-Path $stageDirectory 'SHA256SUMS.txt') -Encoding ascii

    $zipPath = Join-Path $releaseDirectory "$baseName.zip"
    Compress-Archive -Path (Join-Path $stageDirectory '*') -DestinationPath $zipPath -CompressionLevel Optimal
    Remove-Item -LiteralPath $stageDirectory -Recurse -Force
}

$zipChecksums = Get-ChildItem -LiteralPath $releaseDirectory -Filter '*.zip' |
    Sort-Object Name |
    ForEach-Object {
        $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        "$hash *$($_.Name)"
    }
$zipChecksums | Set-Content -LiteralPath (Join-Path $releaseDirectory 'SHA256SUMS.txt') -Encoding ascii

Get-ChildItem -LiteralPath $releaseDirectory | Select-Object Name, Length
