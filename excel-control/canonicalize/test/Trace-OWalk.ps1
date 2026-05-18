# Inline trace of the o-stream walk for the test form. We mark each
# Read-TextProps call site with its pos so we can see what TextProps
# payloads to canonicalise.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Path)
. (Join-Path $PSScriptRoot '..\Cfbf.ps1')
. (Join-Path $PSScriptRoot '..\Oforms.ps1')

$ole = Read-CfbfStreams $Path
$f = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'f' } | Select-Object -First 1
$o = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'o' } | Select-Object -First 1
$fBytes = [byte[]](ConvertTo-StreamBytes $ole.Bytes $f)
$oBytes = [byte[]](ConvertTo-StreamBytes $ole.Bytes $o)

# Walk f to enumerate sites.
$fs = [PadStream]::new($fBytes, 'f')
$sites = Read-FormControl $fs
$os = [PadStream]::new($oBytes, 'o')
foreach ($v in $sites) {
    $before = $os.Pos
    "site cls=$($v.ClsidCacheIndex) oPos before=0x$('{0:X}' -f $before)"
    switch ($v.ClsidCacheIndex) {
        21 { Read-LabelControl $os }
        23 { Read-MorphDataControl $os }
        26 { Read-MorphDataControl $os }
        17 { Read-CommandButtonControl $os }
        default { throw "Unknown idx $($v.ClsidCacheIndex)" }
    }
    "  oPos after=0x$('{0:X}' -f $os.Pos) (consumed $($os.Pos - $before) bytes)"
}
"final oPos=0x$('{0:X}' -f $os.Pos), oLen=$($oBytes.Length)"
"pad ranges in o:"
foreach ($p in $os.PadRanges) { "  off=0x{0:X} len={1}" -f $p.Offset, $p.Length | Write-Host }
