# Compound File Binary Format (CFBF) reader -- minimum needed to find
# the byte offset and length of named streams ('f', 'o', 'CompObj') inside
# a .frx, so the MS-OFORMS walker can map each consumed stream byte back
# to its position in the original .frx file.
#
# CFBF spec: MS-CFB, https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
# Layout: header (512 bytes) at start of CFBF region; FAT chains; mini-FAT chains for
# streams < 4096 bytes; directory chain holds named-stream metadata.
#
# We return, per stream:
#   {
#     Name      = 'f' | 'o' | 'CompObj' | ...
#     Size      = uncompressed byte length
#     Sectors   = ordered list of (fileOffset, length) tuples — concatenating those
#                 byte ranges in order yields the stream content. Each tuple maps
#                 directly into the original file, so mutating bytes within these
#                 ranges modifies the stream in-place without repacking.
#     IsMini    = $true if stored in the mini-stream
#   }
#
# Caller composes the stream content by concatenating Sectors slices; when it
# needs to zero a padding byte at stream offset P, it walks Sectors looking for
# the slice that contains P and writes back to (fileOffset + (P - sliceStart)).

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# CFBF special FAT sector values. Cast literals to [uint32] explicitly --
# PowerShell parses 0xFFFFFFFA as [int] (= -6) which trips strict comparisons.
$Script:CFB_MAXREGSECT  = [uint32]4294967290   # 0xFFFFFFFA
$Script:CFB_DIFSECT     = [uint32]4294967292   # 0xFFFFFFFC
$Script:CFB_FATSECT     = [uint32]4294967293   # 0xFFFFFFFD
$Script:CFB_ENDOFCHAIN  = [uint32]4294967294   # 0xFFFFFFFE
$Script:CFB_FREESECT    = [uint32]4294967295   # 0xFFFFFFFF
$Script:CFB_NOSTREAM    = [uint32]4294967295   # 0xFFFFFFFF

function Read-UInt16LE([byte[]]$buf, [int]$off) {
    return [uint16]([uint16]$buf[$off] -bor ([uint16]$buf[$off+1] -shl 8))
}
function Read-UInt32LE([byte[]]$buf, [int]$off) {
    return ([uint32]$buf[$off]) -bor (([uint32]$buf[$off+1]) -shl 8) -bor `
           (([uint32]$buf[$off+2]) -shl 16) -bor (([uint32]$buf[$off+3]) -shl 24)
}

# Scan a buffer for the CFBF magic ("D0 CF 11 E0 A1 B1 1A E1"). .frx has a
# 24-byte OLE wrapper before the CFBF header.
function Find-CfbfStart([byte[]]$bytes) {
    for ($i = 0; $i -lt [Math]::Min($bytes.Length - 8, 256); $i++) {
        if ($bytes[$i]   -eq 0xD0 -and $bytes[$i+1] -eq 0xCF -and
            $bytes[$i+2] -eq 0x11 -and $bytes[$i+3] -eq 0xE0 -and
            $bytes[$i+4] -eq 0xA1 -and $bytes[$i+5] -eq 0xB1 -and
            $bytes[$i+6] -eq 0x1A -and $bytes[$i+7] -eq 0xE1) {
            return $i
        }
    }
    throw "CFBF magic not found in first 256 bytes"
}

# Read the FAT chain starting at sector index `first`. Returns ordered
# array of sector indices.
function Get-FatChain([uint32[]]$fat, [uint32]$first) {
    $chain = New-Object System.Collections.ArrayList
    $cur = $first
    while ($cur -le $Script:CFB_MAXREGSECT) {
        [void]$chain.Add([uint32]$cur)
        if ($cur -ge [uint32]$fat.Length) {
            throw ("FAT chain sector out of range: $cur >= $($fat.Length)")
        }
        $cur = [uint32]$fat[$cur]
    }
    return ,$chain.ToArray([uint32])
}

# Build sectors -> (fileOffset, length) tuples for a regular-stream chain.
# `cfbfBase` is the byte offset of CFBF magic within the .frx; sectors are
# 1-indexed past the header (sector 0 = first sector after the 512-byte header).
function Get-RegularSectorRanges([uint32[]]$chain, [int]$sectorSize, [int]$cfbfBase, [long]$size) {
    $ranges = New-Object System.Collections.ArrayList
    $remaining = [int64]$size
    foreach ($s in $chain) {
        if ($remaining -le 0) { break }
        $take = [int][Math]::Min($remaining, [int64]$sectorSize)
        $fileOffset = [int64]$cfbfBase + [int64]($sectorSize) + [int64]($s) * [int64]$sectorSize
        [void]$ranges.Add([pscustomobject]@{ FileOffset = $fileOffset; Length = $take })
        $remaining -= $take
    }
    return ,$ranges.ToArray()
}

# Build sectors -> (fileOffset, length) for a mini-stream entry.
# Mini-sectors are 64 bytes inside the mini-stream (which itself is a regular
# chain rooted on the root entry's StartingSectorLocation).
function Get-MiniSectorRanges([uint32[]]$miniChain, [int]$miniStart, [pscustomobject[]]$miniContainerRanges, [int]$miniSectorSize, [long]$size) {
    # miniContainerRanges = the concatenated regular-sector ranges that hold
    # the mini-stream payload. Walk the mini chain, then translate each mini
    # sector index into a (fileOffset, length) inside the container ranges.
    $ranges = New-Object System.Collections.ArrayList
    $remaining = [int64]$size
    foreach ($mini in $miniChain) {
        if ($remaining -le 0) { break }
        $take = [int][Math]::Min($remaining, [int64]$miniSectorSize)
        $miniByteOffset = [int64]$mini * [int64]$miniSectorSize
        # find which container range holds this offset.
        $local = $miniByteOffset
        $found = $false
        foreach ($r in $miniContainerRanges) {
            if ($local -lt $r.Length) {
                [void]$ranges.Add([pscustomobject]@{ FileOffset = ([int64]$r.FileOffset + $local); Length = $take })
                $found = $true
                break
            }
            $local -= $r.Length
        }
        if (-not $found) { throw "Mini sector $mini past mini-stream container" }
        $remaining -= $take
    }
    return ,$ranges.ToArray()
}

function Read-CfbfStreams([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $cfbfBase = Find-CfbfStart $bytes

    # Header at cfbfBase, 512 bytes.
    $sectorShift     = Read-UInt16LE $bytes ($cfbfBase + 0x1E)
    $miniSectorShift = Read-UInt16LE $bytes ($cfbfBase + 0x20)
    $sectorSize      = 1 -shl $sectorShift
    $miniSectorSize  = 1 -shl $miniSectorShift
    $numDirSectors   = Read-UInt32LE $bytes ($cfbfBase + 0x28)  # only used for v4
    $numFatSectors   = Read-UInt32LE $bytes ($cfbfBase + 0x2C)
    $firstDirSector  = Read-UInt32LE $bytes ($cfbfBase + 0x30)
    $miniStreamCutoff= Read-UInt32LE $bytes ($cfbfBase + 0x38)
    $firstMiniFatSec = Read-UInt32LE $bytes ($cfbfBase + 0x3C)
    $numMiniFatSec   = Read-UInt32LE $bytes ($cfbfBase + 0x40)
    $firstDifatSec   = Read-UInt32LE $bytes ($cfbfBase + 0x44)
    $numDifatSec     = Read-UInt32LE $bytes ($cfbfBase + 0x48)

    if ($sectorSize -ne 512 -and $sectorSize -ne 4096) {
        throw "Unsupported sector size: $sectorSize"
    }
    if ($miniSectorSize -ne 64) {
        throw "Unsupported mini sector size: $miniSectorSize"
    }

    # --- DIFAT (sectors holding FAT-sector indices) ---
    $difat = New-Object System.Collections.ArrayList
    # First 109 DIFAT entries live in the header at +0x4C.
    for ($i = 0; $i -lt 109; $i++) {
        $v = Read-UInt32LE $bytes ($cfbfBase + 0x4C + ($i * 4))
        if ($v -le $Script:CFB_MAXREGSECT) { [void]$difat.Add([uint32]$v) }
    }
    # Walk extended DIFAT chain if present.
    $curDifat = $firstDifatSec
    while ($curDifat -le $Script:CFB_MAXREGSECT) {
        $sectorFileOff = $cfbfBase + $sectorSize + ($curDifat * $sectorSize)
        for ($i = 0; $i -lt (($sectorSize / 4) - 1); $i++) {
            $v = Read-UInt32LE $bytes ($sectorFileOff + ($i * 4))
            if ($v -le $Script:CFB_MAXREGSECT) { [void]$difat.Add([uint32]$v) }
        }
        $curDifat = Read-UInt32LE $bytes ($sectorFileOff + $sectorSize - 4)
    }

    # --- FAT (sector chain table) ---
    $fatEntriesPerSector = $sectorSize / 4
    $fat = New-Object 'System.Collections.Generic.List[uint32]'
    foreach ($fatSec in $difat) {
        $sectorFileOff = $cfbfBase + $sectorSize + ([int64]$fatSec * $sectorSize)
        for ($i = 0; $i -lt $fatEntriesPerSector; $i++) {
            $fat.Add([uint32](Read-UInt32LE $bytes ($sectorFileOff + ($i * 4))))
        }
    }
    $fatArr = $fat.ToArray()

    # --- Directory chain ---
    $dirChain = Get-FatChain $fatArr $firstDirSector
    $dirRanges = Get-RegularSectorRanges $dirChain $sectorSize $cfbfBase ([int64]($dirChain.Length * $sectorSize))

    # Read directory entries (128 bytes each).
    $dirEntries = New-Object System.Collections.ArrayList
    foreach ($r in $dirRanges) {
        $entriesInSector = $r.Length / 128
        for ($i = 0; $i -lt $entriesInSector; $i++) {
            $entOff = [int64]$r.FileOffset + ($i * 128)
            # Name (UTF-16LE), up to 64 bytes (32 wchars), length in bytes (incl. trailing null) at +0x40.
            $nameLen = Read-UInt16LE $bytes $entOff   # name uses offset 0..0x3F
            # actually name length is at +0x40
            $nameLen = Read-UInt16LE $bytes ($entOff + 0x40)
            $name = ''
            if ($nameLen -gt 0 -and $nameLen -le 64) {
                # exclude trailing null
                $nameBytesCount = $nameLen - 2
                if ($nameBytesCount -gt 0) {
                    $name = [System.Text.Encoding]::Unicode.GetString($bytes, $entOff, $nameBytesCount)
                }
            }
            $type      = $bytes[$entOff + 0x42]
            $startSec  = Read-UInt32LE $bytes ($entOff + 0x74)
            $sizeLow   = Read-UInt32LE $bytes ($entOff + 0x78)
            $sizeHigh  = Read-UInt32LE $bytes ($entOff + 0x7C)
            $sizeTotal = [int64]$sizeLow -bor ([int64]$sizeHigh -shl 32)
            [void]$dirEntries.Add([pscustomobject]@{
                Name       = $name
                Type       = $type
                StartSector= $startSec
                Size       = $sizeTotal
                EntryOffset= $entOff
            })
        }
    }

    # --- Root entry (index 0) controls the mini-stream container. ---
    $root = $dirEntries[0]
    $miniContainerChain  = Get-FatChain $fatArr $root.StartSector
    $miniContainerRanges = Get-RegularSectorRanges $miniContainerChain $sectorSize $cfbfBase $root.Size

    # --- Mini-FAT chain ---
    $miniFat = New-Object 'System.Collections.Generic.List[uint32]'
    if ($firstMiniFatSec -le $Script:CFB_MAXREGSECT) {
        $miniFatChain = Get-FatChain $fatArr $firstMiniFatSec
        foreach ($mfs in $miniFatChain) {
            $sectorFileOff = $cfbfBase + $sectorSize + ([int64]$mfs * $sectorSize)
            for ($i = 0; $i -lt $fatEntriesPerSector; $i++) {
                $miniFat.Add([uint32](Read-UInt32LE $bytes ($sectorFileOff + ($i * 4))))
            }
        }
    }
    $miniFatArr = $miniFat.ToArray()

    # --- Build stream descriptors for non-storage non-empty entries. ---
    $streams = New-Object System.Collections.ArrayList
    for ($idx = 0; $idx -lt $dirEntries.Count; $idx++) {
        $e = $dirEntries[$idx]
        if ($e.Type -ne 2) { continue }    # 2 = stream
        if ($e.Size -le 0) { continue }
        $isMini = ($e.Size -lt $miniStreamCutoff -and $idx -ne 0)
        if ($isMini) {
            $chain = Get-FatChain $miniFatArr $e.StartSector
            $ranges = Get-MiniSectorRanges $chain $e.StartSector $miniContainerRanges $miniSectorSize $e.Size
        } else {
            $chain = Get-FatChain $fatArr $e.StartSector
            $ranges = Get-RegularSectorRanges $chain $sectorSize $cfbfBase $e.Size
        }
        [void]$streams.Add([pscustomobject]@{
            Name    = $e.Name
            Size    = $e.Size
            Sectors = $ranges
            IsMini  = $isMini
        })
    }

    return [pscustomobject]@{
        Bytes      = $bytes
        CfbfBase   = $cfbfBase
        SectorSize = $sectorSize
        Streams    = $streams
    }
}

# Concatenate sector ranges into a single byte[] (the stream content).
function ConvertTo-StreamBytes([byte[]]$fileBytes, [object]$stream) {
    $out = New-Object byte[] $stream.Size
    $pos = 0
    foreach ($r in $stream.Sectors) {
        [Array]::Copy($fileBytes, [int]$r.FileOffset, $out, $pos, [int]$r.Length)
        $pos += $r.Length
    }
    return ,$out
}

# Given a stream-relative byte offset, return the file-relative byte offset
# by walking the stream's Sectors.
function ConvertTo-FileOffset([object]$stream, [int64]$streamOffset) {
    $cur = $streamOffset
    foreach ($r in $stream.Sectors) {
        if ($cur -lt $r.Length) {
            return [int64]$r.FileOffset + $cur
        }
        $cur -= $r.Length
    }
    throw "Stream offset $streamOffset past stream size $($stream.Size)"
}
