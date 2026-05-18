# N-export convergence test: export the same workbook N times via separate
# Excel sessions, canonicalize each output, and verify every canonical form
# is byte-identical to the first.
#
# Tighter than Export-Samples + Diff-Bytes because it exercises multiple
# fresh Excel processes -- the worst case for non-determinism (every save
# starts with a freshly initialised heap layout in EXCEL.EXE).

[CmdletBinding()]
param(
    [string]$Workbook = (Join-Path $PSScriptRoot 'userform.xlsm'),
    [string]$WorkDir  = (Join-Path $PSScriptRoot 'convergence'),
    [int]$Iterations  = 5
)

$ErrorActionPreference = 'Stop'
$VBEXT_CT_MSFORM = 3

function New-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    return $xl
}
function Get-FormComp($wb) { foreach ($c in $wb.VBProject.VBComponents) { if ($c.Type -eq $VBEXT_CT_MSFORM) { return $c } } }

if (Test-Path $WorkDir) { Remove-Item -Recurse -Force $WorkDir }
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$canonicalSets = @()
for ($i = 1; $i -le $Iterations; $i++) {
    $iterDir = Join-Path $WorkDir "iter$i"
    New-Item -ItemType Directory -Path $iterDir -Force | Out-Null
    $xl = New-Excel
    $pid1 = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime -Desc | Select-Object -First 1).Id
    try {
        $wb = $xl.Workbooks.Open($Workbook)
        $wb.Save()
        $form = Get-FormComp $wb
        $form.Export((Join-Path $iterDir "$($form.Name).frm"))
        $wb.Close($false)
    } finally {
        $xl.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        [GC]::Collect()
        if ($pid1) { Start-Sleep -ms 300; try { Stop-Process -Id $pid1 -Force -EA SilentlyContinue } catch {} }
    }
    $frx = Get-ChildItem $iterDir -Filter '*.frx' | Select-Object -First 1
    pwsh -NoProfile -File (Join-Path $PSScriptRoot '..\Canonicalize-Frx.ps1') -Path $frx.FullName -InPlace | Out-Null
    $canonicalSets += ,([System.IO.File]::ReadAllBytes($frx.FullName))
    Write-Host ("  iter $i -> {0} bytes" -f $canonicalSets[-1].Length)
}

$reference = $canonicalSets[0]
$allEqual = $true
for ($i = 1; $i -lt $canonicalSets.Length; $i++) {
    $other = $canonicalSets[$i]
    if ($reference.Length -ne $other.Length) { $allEqual = $false; Write-Host "iter $($i+1): size differs" -ForegroundColor Red; continue }
    $d = 0
    for ($j = 0; $j -lt $reference.Length; $j++) {
        if ($reference[$j] -ne $other[$j]) { $d++ }
    }
    if ($d -ne 0) { $allEqual = $false; Write-Host "iter $($i+1): $d byte diffs vs iter 1" -ForegroundColor Red }
}
if ($allEqual) {
    Write-Host "PASS: all $Iterations canonical exports byte-identical ($($reference.Length) bytes)" -ForegroundColor Green
    exit 0
}
Write-Host "FAIL: canonical forms diverge across iterations" -ForegroundColor Red
exit 1
