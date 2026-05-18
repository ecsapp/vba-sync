# Probe: load Cfbf.ps1 and list streams in a .frx.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Path)
. (Join-Path $PSScriptRoot '..\Cfbf.ps1')
$ole = Read-CfbfStreams $Path
"CFBF base: $($ole.CfbfBase) (0x{0:X})" -f $ole.CfbfBase | Write-Host
"Sector size: $($ole.SectorSize)" | Write-Host
"File size: $($ole.Bytes.Length)" | Write-Host
"Streams:" | Write-Host
foreach ($s in $ole.Streams) {
    $r0 = if ($s.Sectors.Count -gt 0) { '0x{0:X}+{1}' -f ($s.Sectors[0].FileOffset), $s.Sectors[0].Length } else { '(empty)' }
    "  '{0,-12}' size={1,6} mini={2,-5} firstRange={3}" -f $s.Name, $s.Size, $s.IsMini, $r0 | Write-Host
}
