# Hex dump a slice of a file. Default chunk = 64 bytes around offset.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [int]$Offset = 0,
    [int]$Length = 64
)
$bytes = [System.IO.File]::ReadAllBytes($Path)
$end = [Math]::Min($bytes.Length, $Offset + $Length)
$start = [Math]::Max(0, $Offset)
$line = New-Object System.Text.StringBuilder
for ($i = $start; $i -lt $end; $i += 16) {
    $hex = ''
    $asc = ''
    for ($j = 0; $j -lt 16 -and ($i + $j) -lt $end; $j++) {
        $b = $bytes[$i + $j]
        $hex += ('{0:X2} ' -f $b)
        $asc += ($(if ($b -ge 32 -and $b -lt 127) { [char]$b } else { '.' }))
    }
    '{0,8:X} {1,-48} |{2}|' -f $i, $hex, $asc | Write-Host
}
