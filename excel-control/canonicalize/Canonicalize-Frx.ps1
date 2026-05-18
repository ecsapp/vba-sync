# Canonicalize-Frx.ps1
#
# Zeroes the alignment-padding bytes inside a UserForm .frx so back-to-back
# exports from Excel produce identical files. Excel's MS-OFORMS writer leaks
# uninitialised heap into padding slots; the bytes are skipped by Excel's
# importer (proven via round-trip with 0xFF probe) so zeroing them is safe
# and the only practical way to suppress .frx churn in git history.
#
# Usage:
#   .\Canonicalize-Frx.ps1 -Path 'form.frx'          # writes <Path>.canonical
#   .\Canonicalize-Frx.ps1 -Path 'form.frx' -InPlace # overwrites form.frx
#
# Exit codes:
#   0  success
#   2  parse error (file looks like a .frx but MS-OFORMS walk failed)
#   3  I/O error (file not found, write failed, etc.)
#
# The parser is a port of oletools/oleform.py with two structural fixes that
# oleform never made:
#   1. Variable-length strings in SiteExtraDataBlock / MorphDataExtraDataBlock
#      / LabelExtraDataBlock / CommandButtonExtraDataBlock are padded to 4-byte
#      boundaries before the next field. oleform notes "Sometimes it looks
#      like 2 extra null bytes go here" but does not implement the alignment.
#   2. CommandButton DataBlock + ExtraDataBlock is decoded instead of skipped,
#      so caption-trailing padding is identified instead of being lumped into
#      an opaque cb-jump tail.
#
# In addition to OFORMS padding, the canonicalizer zeroes:
#   - CFBF directory entry CreationTime + ModifiedTime FILETIMEs for ALL
#     entries (Excel updates these on every save; they aren't form content).

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [switch]$InPlace
)

$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

. (Join-Path $PSScriptRoot 'Cfbf.ps1')
. (Join-Path $PSScriptRoot 'Oforms.ps1')

function Set-RangeZero([byte[]]$bytes, [int64]$offset, [int]$length) {
    if ($length -le 0) { return 0 }
    $end = [Math]::Min([int64]$bytes.Length, $offset + $length)
    $written = 0
    for ($i = $offset; $i -lt $end; $i++) {
        if ($bytes[$i] -ne 0) {
            $bytes[$i] = 0
            $written++
        } else {
            # Count zero-overwrite as a hit too -- the byte was inside a
            # padding slot, just happened to already be zero. Callers want
            # total padding-bytes-zeroed, not just bytes changed.
            $written++
        }
    }
    return $written
}

# Map a stream-relative offset range to file-relative byte ranges via the
# CFBF sector chain. A single stream-range may cross sector boundaries
# (especially in the mini-stream where each entry is only 64 bytes), so
# this returns multiple file ranges.
function Get-FileRanges($stream, [int64]$streamOffset, [int]$length) {
    $ranges = New-Object System.Collections.ArrayList
    $cur = $streamOffset
    $remaining = $length
    foreach ($r in $stream.Sectors) {
        if ($remaining -le 0) { break }
        if ($cur -lt $r.Length) {
            $available = [int]([Math]::Min([int64]$remaining, [int64]($r.Length - $cur)))
            [void]$ranges.Add([pscustomobject]@{
                FileOffset = ([int64]$r.FileOffset + $cur)
                Length     = $available
            })
            $remaining -= $available
            $cur = 0
        } else {
            $cur -= $r.Length
        }
    }
    if ($remaining -gt 0) {
        throw "Stream range $streamOffset+$length past stream size"
    }
    return $ranges
}

try {
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "File not found: $Path"
        exit 3
    }

    $absPath = (Resolve-Path -LiteralPath $Path).Path

    try {
        $ole = Read-CfbfStreams $absPath
    } catch {
        Write-Error "CFBF parse failed: $($_.Exception.Message)"
        exit 2
    }

    $bytes = $ole.Bytes
    $totalZeroed = 0

    $fStream = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'f' } | Select-Object -First 1
    $oStream = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'o' } | Select-Object -First 1
    if (-not $fStream) { Write-Error "No 'f' stream in CFBF -- not a .frx?"; exit 2 }
    if (-not $oStream) { Write-Error "No 'o' stream in CFBF -- not a .frx?"; exit 2 }

    $fBytes = [byte[]](ConvertTo-StreamBytes $bytes $fStream)
    $oBytes = [byte[]](ConvertTo-StreamBytes $bytes $oStream)

    try {
        $pads = Get-OformsPaddingRanges -FStreamBytes $fBytes -OStreamBytes $oBytes
    } catch {
        Write-Error "OFORMS walk failed: $($_.Exception.Message)"
        exit 2
    }

    foreach ($p in $pads.FPad) {
        foreach ($fr in (Get-FileRanges $fStream $p.Offset $p.Length)) {
            $totalZeroed += Set-RangeZero $bytes $fr.FileOffset $fr.Length
        }
    }
    foreach ($p in $pads.OPad) {
        foreach ($fr in (Get-FileRanges $oStream $p.Offset $p.Length)) {
            $totalZeroed += Set-RangeZero $bytes $fr.FileOffset $fr.Length
        }
    }

    # Zero CFBF directory entry timestamps (Creation + Modified FILETIME).
    # The directory chain is at sector firstDirSector and we already located
    # its file ranges; walk all 128-byte entries and zero offsets 0x64..0x73
    # within each entry (16 bytes total). Excel updates the root entry's
    # ModifiedTime on every save -- not form content, pure churn.
    $sectorSize = $ole.SectorSize
    $cfbfBase = $ole.CfbfBase
    $firstDirSector = Read-UInt32LE $bytes ($cfbfBase + 0x30)
    $fatEntriesPerSector = $sectorSize / 4
    # Re-derive the dir chain using the same FAT we built in Read-CfbfStreams.
    # Simpler: walk dirSectors by scanning all known stream entries' parent
    # storage relationships -- but for canonicalization we just need to find
    # every 128-byte entry. Easiest: rebuild the FAT here.
    $difat = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 109; $i++) {
        $v = Read-UInt32LE $bytes ($cfbfBase + 0x4C + ($i * 4))
        if ($v -le [uint32]4294967290) { [void]$difat.Add([uint32]$v) }
    }
    $firstDifatSec = Read-UInt32LE $bytes ($cfbfBase + 0x44)
    $curDifat = $firstDifatSec
    while ($curDifat -le [uint32]4294967290) {
        $sectorFileOff = $cfbfBase + $sectorSize + ($curDifat * $sectorSize)
        for ($i = 0; $i -lt (($sectorSize / 4) - 1); $i++) {
            $v = Read-UInt32LE $bytes ($sectorFileOff + ($i * 4))
            if ($v -le [uint32]4294967290) { [void]$difat.Add([uint32]$v) }
        }
        $curDifat = Read-UInt32LE $bytes ($sectorFileOff + $sectorSize - 4)
    }
    $fatList = New-Object 'System.Collections.Generic.List[uint32]'
    foreach ($fs in $difat) {
        $soff = $cfbfBase + $sectorSize + ([int64]$fs * $sectorSize)
        for ($i = 0; $i -lt $fatEntriesPerSector; $i++) {
            $fatList.Add([uint32](Read-UInt32LE $bytes ($soff + ($i * 4))))
        }
    }
    $fatArr = $fatList.ToArray()
    $dirChain = Get-FatChain $fatArr $firstDirSector
    $entriesPerSector = $sectorSize / 128
    foreach ($ds in $dirChain) {
        $sectorFileOff = [int64]$cfbfBase + $sectorSize + ([int64]$ds * $sectorSize)
        for ($i = 0; $i -lt $entriesPerSector; $i++) {
            $entOff = $sectorFileOff + ($i * 128)
            # CreationTime at +0x64 (8 bytes), ModifiedTime at +0x6C (8 bytes).
            $totalZeroed += Set-RangeZero $bytes ($entOff + 0x64) 16
        }
    }

    $outPath = if ($InPlace) { $absPath } else { "$absPath.canonical" }

    try {
        [System.IO.File]::WriteAllBytes($outPath, $bytes)
    } catch {
        Write-Error "Write failed: $($_.Exception.Message)"
        exit 3
    }

    Write-Host "canonicalized $totalZeroed padding bytes"
    exit 0
} catch {
    Write-Error $_
    exit 3
}
