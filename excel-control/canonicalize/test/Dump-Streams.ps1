# Extract f and o mini-streams from a raw .frx for offline analysis.
# Writes <prefix>.f.bin and <prefix>.o.bin alongside the input.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$FrxPath
)
$bytes = [System.IO.File]::ReadAllBytes($FrxPath)
$cfbf = 24
$secSize = 512
function R32($off) { $bytes[$off] + $bytes[$off+1]*256 + $bytes[$off+2]*65536 + [int64]$bytes[$off+3]*16777216 }
function R16($off) { $bytes[$off] + $bytes[$off+1]*256 }
function SectorOff($sec) { $cfbf + 512 + $sec*$secSize }

# Mini-FAT
$firstMiniFat = R32 ($cfbf+60)
$miniFat = @()
$mfOff = SectorOff $firstMiniFat
for ($i = 0; $i -lt $secSize/4; $i++) {
    $v = R32 ($mfOff + $i*4)
    if ($v -ge 0x80000000) { $miniFat += -1 } else { $miniFat += $v }
}

# Walk directory: dir sector R32(cfbf+48)
$firstDir = R32 ($cfbf+48)
$dirOff = SectorOff $firstDir
$entries = @{}
for ($i = 0; $i -lt 4; $i++) {
    $b = $dirOff + $i*128
    $nameLen = R16 ($b+64)
    if ($nameLen -lt 2) { continue }
    $name = ''
    for ($j = 0; $j -lt ($nameLen-2)/2; $j++) {
        $ch = $bytes[$b + $j*2] + $bytes[$b + $j*2 + 1]*256
        if ($ch -ne 0) { $name += [char]$ch }
    }
    $type = $bytes[$b+66]
    $startSec = R32 ($b+116)
    $size = R32 ($b+120)
    if ($type -eq 5) {
        $entries['ROOT_START'] = $startSec
        $entries['ROOT_SIZE'] = $size
    } else {
        $entries[$name] = @{ start = $startSec; size = $size }
    }
}

$miniBase = SectorOff $entries['ROOT_START']
$dir = Split-Path -Parent $FrxPath
$prefix = [System.IO.Path]::GetFileNameWithoutExtension($FrxPath)

foreach ($name in @('f', 'o')) {
    if (-not $entries.ContainsKey($name)) { continue }
    $e = $entries[$name]
    $size = $e.size
    $sec = $e.start
    $buf = New-Object 'System.Byte[]' $size
    $idx = 0
    while ($sec -ge 0 -and $idx -lt $size) {
        $src = $miniBase + $sec * 64
        $copy = if ($idx + 64 -gt $size) { $size - $idx } else { 64 }
        [Array]::Copy($bytes, $src, $buf, $idx, $copy)
        $idx += $copy
        if ($sec -ge $miniFat.Count) { break }
        $sec = $miniFat[$sec]
        if ($sec -lt 0) { break }
    }
    $outPath = Join-Path $dir ("$prefix.$name.bin")
    [System.IO.File]::WriteAllBytes($outPath, $buf)
    Write-Host "$outPath  ($($buf.Length) bytes)"
}
