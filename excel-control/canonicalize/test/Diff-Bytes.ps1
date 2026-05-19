# Helper: print byte-level diffs between two files (positions + values).
[CmdletBinding()]
param([Parameter(Mandatory)][string]$A, [Parameter(Mandatory)][string]$B)

$bytesA = [System.IO.File]::ReadAllBytes($A)
$bytesB = [System.IO.File]::ReadAllBytes($B)
if ($bytesA.Length -ne $bytesB.Length) {
    Write-Host "Sizes differ: A=$($bytesA.Length) B=$($bytesB.Length)"; exit
}
$n = $bytesA.Length
$diffs = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $n; $i++) {
    if ($bytesA[$i] -ne $bytesB[$i]) {
        [void]$diffs.Add([pscustomobject]@{Offset=$i; OffsetHex=("0x{0:X}" -f $i); A=("{0:X2}" -f $bytesA[$i]); B=("{0:X2}" -f $bytesB[$i])})
    }
}
"$($diffs.Count) byte differences"
$diffs | Format-Table -AutoSize | Out-String | Write-Host
