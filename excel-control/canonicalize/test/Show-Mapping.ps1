# For each byte that differs between A and B, identify which CFBF
# stream (or 'header/dir/fat') it falls inside.
[CmdletBinding()]
param(
    [string]$A = (Join-Path $PSScriptRoot 'A\frmLogin.frx'),
    [string]$B = (Join-Path $PSScriptRoot 'B\frmLogin.frx')
)
. (Join-Path $PSScriptRoot '..\Cfbf.ps1')

$ole = Read-CfbfStreams $A
$bytesA = [System.IO.File]::ReadAllBytes($A)
$bytesB = [System.IO.File]::ReadAllBytes($B)
$n = $bytesA.Length

function Get-Region([int64]$off, [object]$ole) {
    foreach ($s in $ole.Streams) {
        $local = 0
        foreach ($r in $s.Sectors) {
            if ($off -ge $r.FileOffset -and $off -lt ($r.FileOffset + $r.Length)) {
                return ("stream '$($s.Name.Trim())' local-off=$($local + ($off - $r.FileOffset))")
            }
            $local += $r.Length
        }
    }
    return 'CFBF-header/dir/fat (not in any stream)'
}

for ($i = 0; $i -lt $n; $i++) {
    if ($bytesA[$i] -ne $bytesB[$i]) {
        $reg = Get-Region $i $ole
        '0x{0:X4}  A={1:X2}  B={2:X2}  {3}' -f $i, $bytesA[$i], $bytesB[$i], $reg | Write-Host
    }
}
