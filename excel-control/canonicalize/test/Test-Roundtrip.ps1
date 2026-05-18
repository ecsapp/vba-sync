# Round-trip test: import a canonicalized .frx into a fresh workbook,
# save, re-export, canonicalize again, and verify byte-for-byte equality
# with the original canonical form.
#
# This proves:
#   1. Excel still accepts the canonicalized .frx (no parse errors, all
#      controls preserved -- their property values survive padding zeros).
#   2. The canonical form is a fixed point of the export-canonicalize
#      pipeline (idempotent).
#
# Usage: pwsh .\Test-Roundtrip.ps1

[CmdletBinding()]
param(
    [string]$OrigWorkbook = (Join-Path $PSScriptRoot 'userform.xlsm'),
    [string]$WorkDir      = (Join-Path $PSScriptRoot 'roundtrip')
)

$ErrorActionPreference = 'Stop'
$VBEXT_CT_MSFORM = 3
$XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52

function New-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    return $xl
}

function Get-FormComp($wb) {
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq $VBEXT_CT_MSFORM) { return $c }
    }
    return $null
}

# Export the existing form from $OrigWorkbook into $WorkDir/original/.
if (Test-Path $WorkDir) { Remove-Item -Recurse -Force $WorkDir }
New-Item -ItemType Directory -Path "$WorkDir/original" -Force | Out-Null
New-Item -ItemType Directory -Path "$WorkDir/rebuilt"  -Force | Out-Null

$xl = New-Excel
$pid1 = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime -Desc | Select-Object -First 1).Id
try {
    $wb = $xl.Workbooks.Open($OrigWorkbook)
    $form = Get-FormComp $wb
    if (-not $form) { throw "no UserForm in $OrigWorkbook" }
    $formName = $form.Name
    $form.Export((Join-Path "$WorkDir/original" "$formName.frm"))
    $wb.Close($false)
} finally {
    $xl.Quit() | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    if ($pid1) { Start-Sleep -ms 300; try { Stop-Process -Id $pid1 -Force -EA SilentlyContinue } catch {} }
}

$origFrx = Get-ChildItem "$WorkDir/original" -Filter '*.frx' | Select-Object -First 1
$origFrm = Get-ChildItem "$WorkDir/original" -Filter '*.frm' | Select-Object -First 1

# Canonicalize the original (in place).
pwsh -NoProfile -File (Join-Path $PSScriptRoot '..\Canonicalize-Frx.ps1') -Path $origFrx.FullName -InPlace | Out-Null
$canonOrigBytes = [System.IO.File]::ReadAllBytes($origFrx.FullName)

# Build a fresh workbook, import the canonicalized .frm/.frx, save, re-export.
$xl2 = New-Excel
$pid2 = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime -Desc | Select-Object -First 1).Id
try {
    $wb = $xl2.Workbooks.Add()
    # Save as macro-enabled first so VBProject is writable.
    $savePath = Join-Path $WorkDir 'rebuilt.xlsm'
    if (Test-Path $savePath) { Remove-Item -Force $savePath }
    $wb.SaveAs($savePath, $XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
    $imported = $wb.VBProject.VBComponents.Import($origFrm.FullName)
    $importedName = $imported.Name
    $wb.Save()
    $imported.Export((Join-Path "$WorkDir/rebuilt" "$importedName.frm"))
    $wb.Close($false)
} finally {
    $xl2.Quit() | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl2) | Out-Null
    [GC]::Collect()
    if ($pid2) { Start-Sleep -ms 300; try { Stop-Process -Id $pid2 -Force -EA SilentlyContinue } catch {} }
}

$rebuiltFrx = Get-ChildItem "$WorkDir/rebuilt" -Filter '*.frx' | Select-Object -First 1
if (-not $rebuiltFrx) { throw "Rebuilt workbook did not export a .frx -- import may have failed" }
pwsh -NoProfile -File (Join-Path $PSScriptRoot '..\Canonicalize-Frx.ps1') -Path $rebuiltFrx.FullName -InPlace | Out-Null
$canonRebuiltBytes = [System.IO.File]::ReadAllBytes($rebuiltFrx.FullName)

if ($canonOrigBytes.Length -ne $canonRebuiltBytes.Length) {
    Write-Host "FAIL: sizes differ (orig=$($canonOrigBytes.Length) rebuilt=$($canonRebuiltBytes.Length))" -ForegroundColor Red
    exit 1
}
$diffs = 0
for ($i = 0; $i -lt $canonOrigBytes.Length; $i++) {
    if ($canonOrigBytes[$i] -ne $canonRebuiltBytes[$i]) { $diffs++ }
}
if ($diffs -eq 0) {
    Write-Host "PASS: round-trip canonical form is a fixed point ($($canonOrigBytes.Length) bytes)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: round-trip produced $diffs byte differences" -ForegroundColor Red
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Diff-Bytes.ps1') -A $origFrx.FullName -B $rebuiltFrx.FullName
    exit 1
}
