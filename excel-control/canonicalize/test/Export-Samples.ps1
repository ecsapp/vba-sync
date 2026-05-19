# Open userform.xlsm via Excel COM, export the UserForm twice into
# distinct destination folders (A and B). Excel's non-deterministic .frx
# writer should produce byte-different .frx files between A and B even
# though the form content is identical -- that mismatch is what the
# canonicalizer must collapse.
#
# Usage: pwsh .\Export-Samples.ps1

[CmdletBinding()]
param(
    [string]$Workbook = (Join-Path $PSScriptRoot 'userform.xlsm'),
    [string]$OutA     = (Join-Path $PSScriptRoot 'A'),
    [string]$OutB     = (Join-Path $PSScriptRoot 'B')
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

function Get-FormComp($wb) {
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq $VBEXT_CT_MSFORM) { return $c }
    }
    return $null
}

function Export-Form([string]$wbPath, [string]$outDir) {
    if (Test-Path $outDir) { Remove-Item -Recurse -Force $outDir }
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null

    $xl = New-Excel
    $excelPid = (Get-Process EXCEL -ErrorAction SilentlyContinue |
        Sort-Object StartTime -Descending | Select-Object -First 1).Id
    try {
        $wb = $xl.Workbooks.Open($wbPath)
        try {
            # Touch + save the workbook so Excel re-serialises form streams.
            $wb.Save()

            $form = Get-FormComp $wb
            if (-not $form) { throw "No UserForm in $wbPath" }

            $frmPath = Join-Path $outDir ($form.Name + '.frm')
            $form.Export($frmPath)
            # Excel writes the paired .frx automatically alongside .frm.
        } finally {
            $wb.Close($false)
        }
    } finally {
        $xl.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        [System.GC]::Collect()
        if ($excelPid) {
            Start-Sleep -Milliseconds 300
            try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
}

Write-Host "Exporting sample A..." -ForegroundColor Cyan
Export-Form $Workbook $OutA
Write-Host "Exporting sample B..." -ForegroundColor Cyan
Export-Form $Workbook $OutB

$frxA = Get-ChildItem -LiteralPath $OutA -Filter '*.frx' | Select-Object -First 1
$frxB = Get-ChildItem -LiteralPath $OutB -Filter '*.frx' | Select-Object -First 1

if (-not $frxA -or -not $frxB) { throw "Missing exported .frx in A or B" }

$bytesA = [System.IO.File]::ReadAllBytes($frxA.FullName)
$bytesB = [System.IO.File]::ReadAllBytes($frxB.FullName)

Write-Host ""
Write-Host "A: $($frxA.FullName) [$($bytesA.Length) bytes]"
Write-Host "B: $($frxB.FullName) [$($bytesB.Length) bytes]"

if ($bytesA.Length -ne $bytesB.Length) {
    Write-Host "Sizes differ" -ForegroundColor Yellow
} else {
    $diffs = 0
    for ($i = 0; $i -lt $bytesA.Length; $i++) {
        if ($bytesA[$i] -ne $bytesB[$i]) { $diffs++ }
    }
    if ($diffs -eq 0) {
        Write-Host "A == B (no byte diff -- nothing to canonicalize)" -ForegroundColor Yellow
    } else {
        Write-Host "A vs B: $diffs byte differences" -ForegroundColor Green
    }
}
