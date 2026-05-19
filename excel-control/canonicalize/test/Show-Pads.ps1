# Run the OFORMS walker over the f and o streams of a .frx and print the
# stream-relative padding ranges it identified.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Path)
. (Join-Path $PSScriptRoot '..\Cfbf.ps1')
. (Join-Path $PSScriptRoot '..\Oforms.ps1')

$ole = Read-CfbfStreams $Path
$f = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'f' } | Select-Object -First 1
$o = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'o' } | Select-Object -First 1
if (-not $f) { throw "No 'f' stream" }
if (-not $o) { throw "No 'o' stream" }

$fBytes = [byte[]](ConvertTo-StreamBytes $ole.Bytes $f)
$oBytes = [byte[]](ConvertTo-StreamBytes $ole.Bytes $o)

$pads = Get-OformsPaddingRanges -FStreamBytes $fBytes -OStreamBytes $oBytes

Write-Host "f stream padding (size=$($fBytes.Length)):"
foreach ($p in $pads.FPad) { "  off=0x{0:X} len={1}" -f $p.Offset, $p.Length | Write-Host }
Write-Host "o stream padding (size=$($oBytes.Length)):"
foreach ($p in $pads.OPad) { "  off=0x{0:X} len={1}" -f $p.Offset, $p.Length | Write-Host }
