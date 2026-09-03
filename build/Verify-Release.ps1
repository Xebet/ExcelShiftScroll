[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ReleaseDirectory = (Join-Path $PSScriptRoot '..\artifacts\release')
)

$ErrorActionPreference = 'Stop'
$releasePath = [System.IO.Path]::GetFullPath($ReleaseDirectory)
$checksumPath = Join-Path $releasePath 'SHA256SUMS.txt'
if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
    throw "Global checksum file is missing: $checksumPath"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$globalLines = Get-Content -LiteralPath $checksumPath | Where-Object { $_.Trim().Length -gt 0 }
foreach ($line in $globalLines) {
    if ($line -notmatch '^([0-9a-fA-F]{64}) \*(.+\.zip)$') {
        throw "Invalid global checksum line: $line"
    }

    $expectedHash = $Matches[1].ToLowerInvariant()
    $zipName = $Matches[2]
    $zipPath = Join-Path $releasePath $zipName
    $actualHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "Checksum mismatch for $zipName"
    }

    $architecture = if ($zipName -match '-x64\.zip$') { 'x64' } elseif ($zipName -match '-x86\.zip$') { 'x86' } else { throw "Unknown package architecture: $zipName" }
    $expectedXll = if ($architecture -eq 'x64') { 'ExcelShiftScroll64.xll' } else { 'ExcelShiftScroll32.xll' }
    $requiredEntries = @(
        $expectedXll,
        'README.md',
        'README.zh-CN.md',
        'INSTALL.txt',
        'Install-CurrentUser.ps1',
        'Uninstall-CurrentUser.ps1',
        'LICENSE',
        'THIRD_PARTY_NOTICES.md',
        'SHA256SUMS.txt'
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
        foreach ($required in $requiredEntries) {
            if ($entryNames -notcontains $required) {
                throw "$zipName is missing required entry $required"
            }
        }

        $internalChecksumEntry = $archive.GetEntry('SHA256SUMS.txt')
        $reader = New-Object System.IO.StreamReader($internalChecksumEntry.Open())
        try {
            $internalLine = $reader.ReadToEnd().Trim()
        }
        finally {
            $reader.Dispose()
        }

        if ($internalLine -notmatch "^([0-9a-fA-F]{64}) \*$([regex]::Escape($expectedXll))$") {
            throw "Invalid internal checksum in $zipName"
        }

        $expectedXllHash = $Matches[1].ToLowerInvariant()
        $xllEntry = $archive.GetEntry($expectedXll)
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $stream = $xllEntry.Open()
        try {
            $actualXllHash = ([System.BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
        }
        finally {
            $stream.Dispose()
            $sha.Dispose()
        }

        if ($actualXllHash -ne $expectedXllHash) {
            throw "Internal XLL checksum mismatch in $zipName"
        }
    }
    finally {
        $archive.Dispose()
    }

    Write-Output "Verified $zipName ($architecture)"
}

if ($globalLines.Count -ne 2) {
    throw "Expected exactly two release ZIP checksums, found $($globalLines.Count)."
}
