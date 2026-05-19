# test-canonicalize-vba.ps1
#
# Convergence test for modFrxCanonicalize.bas:
#   1. Run Export-Samples.ps1 to produce non-deterministic A/frmLogin.frx
#      and B/frmLogin.frx (these differ -- 14 bytes on the reference fixture).
#   2. Spawn a headless Excel COM, add a temporary workbook, import
#      modFrxCanonicalize.bas into its VBProject.
#   3. Call CanonicalizeFrx on A copy and B copy.
#   4. Compare the two canonicalized files byte-for-byte.  PASS iff equal.
#   5. Run a second pass on a re-canonicalized A: idempotence check.
#   6. PID-targeted Excel cleanup (NEVER blanket-kill EXCEL.EXE -- the
#      user may have their own session open).

[CmdletBinding()]
param(
    [switch]$SkipExport,   # reuse existing A/B
    [switch]$KeepExcel     # for debugging
)

$ErrorActionPreference = 'Stop'

$here     = $PSScriptRoot
$testDir  = Join-Path $here 'test'
# .bas lives in the xlam source tree (it ships inside VBA Sync.xlam now).
$basPath  = Resolve-Path (Join-Path $here '..\..\VBA Sync\Modules\modFrxCanonicalize.bas') |
    Select-Object -ExpandProperty Path
$sampleA  = Join-Path $testDir 'A\frmLogin.frx'
$sampleB  = Join-Path $testDir 'B\frmLogin.frx'

if (-not (Test-Path $basPath)) { throw ".bas not found at $basPath" }

# Trust access to the VBA project model (required to call .Import).
# This sets a registry value that allows OM access; if it's already set, no-op.
function Enable-VbaTrust {
    $verKey = (Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object -Property { [version]$_.PSChildName } -Descending |
        Select-Object -First 1)
    if (-not $verKey) { return }
    $secPath = Join-Path $verKey.PSPath 'Excel\Security'
    if (-not (Test-Path $secPath)) {
        New-Item -Path $secPath -Force | Out-Null
    }
    Set-ItemProperty -Path $secPath -Name 'AccessVBOM' -Value 1 -Type DWord -Force
}

function New-ExcelCom {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    return $xl
}

function Get-LastExcelPid {
    (Get-Process EXCEL -ErrorAction SilentlyContinue |
        Sort-Object StartTime -Descending | Select-Object -First 1).Id
}

function Compare-Bytes($pathA, $pathB) {
    $a = [System.IO.File]::ReadAllBytes($pathA)
    $b = [System.IO.File]::ReadAllBytes($pathB)
    if ($a.Length -ne $b.Length) {
        return @{ Equal=$false; Diffs=-1; LenA=$a.Length; LenB=$b.Length }
    }
    $diffs = 0
    for ($i = 0; $i -lt $a.Length; $i++) {
        if ($a[$i] -ne $b[$i]) { $diffs++ }
    }
    return @{ Equal = ($diffs -eq 0); Diffs = $diffs; LenA = $a.Length; LenB = $b.Length }
}

# 1. Make sure we have A/B samples
if (-not $SkipExport) {
    Write-Host "[1/5] Re-exporting samples A/B..." -ForegroundColor Cyan
    & (Join-Path $testDir 'Export-Samples.ps1') | Out-Host
} else {
    Write-Host "[1/5] Skipping export (using existing A/B)" -ForegroundColor DarkGray
}
if (-not (Test-Path $sampleA)) { throw "Missing sample A at $sampleA" }
if (-not (Test-Path $sampleB)) { throw "Missing sample B at $sampleB" }

$preCmp = Compare-Bytes $sampleA $sampleB
Write-Host ("    A vs B before canonicalize: {0} bytes, {1} diffs" -f $preCmp.LenA, $preCmp.Diffs) -ForegroundColor Yellow
if ($preCmp.Diffs -eq 0) {
    Write-Warning "A and B are already identical -- the test fixture is not exercising non-determinism."
}

# 2. Make working copies (we'll canonicalize in-place)
$canonA = Join-Path $testDir 'A\frmLogin.canon.frx'
$canonB = Join-Path $testDir 'B\frmLogin.canon.frx'
Copy-Item -LiteralPath $sampleA -Destination $canonA -Force
Copy-Item -LiteralPath $sampleB -Destination $canonB -Force

# 3. Spawn Excel, import .bas, call CanonicalizeFrx
Enable-VbaTrust

Write-Host "[2/5] Spawning Excel COM..." -ForegroundColor Cyan
$xl = New-ExcelCom
$excelPid = Get-LastExcelPid
Write-Host "    Excel PID: $excelPid"

$zeroedA = -2
$zeroedB = -2
$importedErr = $null
try {
    $wb = $xl.Workbooks.Add()
    try {
        Write-Host "[3/5] Importing $basPath..." -ForegroundColor Cyan
        $proj = $wb.VBProject
        $comp = $proj.VBComponents.Import($basPath)
        # In case Excel renamed it on import, prefer the actual name.
        $modName = $comp.Name
        Write-Host "    Module name: $modName"

        Write-Host "[4/5] Calling CanonicalizeFrx on A..." -ForegroundColor Cyan
        $zeroedA = $xl.Run("$modName.CanonicalizeFrx", $canonA)
        Write-Host "    Returned: $zeroedA bytes zeroed"
        if ($zeroedA -lt 0) {
            $errA = $xl.Run("$modName.CanonicalizeFrxLastError")
            Write-Host "    LastError: $errA" -ForegroundColor Red
        }

        Write-Host "[4/5] Calling CanonicalizeFrx on B..." -ForegroundColor Cyan
        $zeroedB = $xl.Run("$modName.CanonicalizeFrx", $canonB)
        Write-Host "    Returned: $zeroedB bytes zeroed"
        if ($zeroedB -lt 0) {
            $errB = $xl.Run("$modName.CanonicalizeFrxLastError")
            Write-Host "    LastError: $errB" -ForegroundColor Red
        }

        # Idempotence: a second run on canon(A) should zero ZERO bytes.
        Write-Host "[4/5] Calling CanonicalizeFrx on A again (idempotence)..." -ForegroundColor Cyan
        $zeroedA2 = $xl.Run("$modName.CanonicalizeFrx", $canonA)
        Write-Host "    Returned: $zeroedA2 bytes zeroed (must be 0 for idempotence)"
        if ($zeroedA2 -ne 0) {
            Write-Host "    FAIL: canonicalize not idempotent" -ForegroundColor Red
        }
    } catch {
        $importedErr = $_
    } finally {
        $wb.Close($false)
    }
} finally {
    try { $xl.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    [System.GC]::Collect()
    if ($excelPid -and -not $KeepExcel) {
        Start-Sleep -Milliseconds 300
        try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}

if ($importedErr) {
    Write-Host "ERROR during import/run: $importedErr" -ForegroundColor Red
    throw $importedErr
}

if ($zeroedA -lt 0) {
    Write-Host "FAIL: CanonicalizeFrx(A) returned $zeroedA (parse error)" -ForegroundColor Red
    exit 1
}
if ($zeroedB -lt 0) {
    Write-Host "FAIL: CanonicalizeFrx(B) returned $zeroedB (parse error)" -ForegroundColor Red
    exit 1
}

# 4. Compare canonicalized A vs canonicalized B
Write-Host "[5/5] Comparing canonicalized A vs B..." -ForegroundColor Cyan
$postCmp = Compare-Bytes $canonA $canonB
Write-Host ("    canon(A) vs canon(B): {0} bytes vs {1} bytes, {2} diffs" -f $postCmp.LenA, $postCmp.LenB, $postCmp.Diffs) -ForegroundColor Yellow

if ($postCmp.Equal) {
    Write-Host ""
    Write-Host "PASS: canon(A) == canon(B)  (zeroed A=$zeroedA, B=$zeroedB)" -ForegroundColor Green

    # Round-trip check: import canonicalized .frx into a new Excel workbook,
    # re-export, canonicalize, and compare to the original canonicalized .frx.
    Write-Host ""
    Write-Host "[bonus] Round-trip test: canonicalize -> Excel import -> re-export -> canonicalize" -ForegroundColor Cyan
    $roundtripDir = Join-Path $testDir 'RT'
    if (Test-Path $roundtripDir) { Remove-Item -Recurse -Force $roundtripDir }
    New-Item -ItemType Directory -Path $roundtripDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $testDir 'A\frmLogin.frm') -Destination (Join-Path $roundtripDir 'frmLogin.frm') -Force
    Copy-Item -LiteralPath $canonA -Destination (Join-Path $roundtripDir 'frmLogin.frx') -Force

    $rtXl = New-ExcelCom
    $rtPid = Get-LastExcelPid
    $rtFrxOut = Join-Path $roundtripDir 'frmLogin.reexport.frx'
    $rtOk = $false
    try {
        $rtWb = $rtXl.Workbooks.Add()
        try {
            $rtComp = $rtWb.VBProject.VBComponents.Import((Join-Path $roundtripDir 'frmLogin.frm'))
            $reFrmPath = Join-Path $roundtripDir 'reexport.frm'
            $rtComp.Export($reFrmPath)
            $rtFrxRaw = [System.IO.Path]::ChangeExtension($reFrmPath, '.frx')
            Copy-Item -LiteralPath $rtFrxRaw -Destination $rtFrxOut -Force
            $rtOk = $true
        } finally {
            $rtWb.Close($false)
        }
    } finally {
        try { $rtXl.Quit() | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rtXl) | Out-Null } catch {}
        [System.GC]::Collect()
        if ($rtPid) {
            Start-Sleep -Milliseconds 300
            try { Stop-Process -Id $rtPid -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    if (-not $rtOk) {
        Write-Host "    Round-trip import/export failed" -ForegroundColor Yellow
        exit 0
    }

    # Canonicalize the re-exported .frx and compare to the original canon(A).
    $rt2Xl = New-ExcelCom
    $rt2Pid = Get-LastExcelPid
    try {
        $rt2Wb = $rt2Xl.Workbooks.Add()
        try {
            $rt2Comp = $rt2Wb.VBProject.VBComponents.Import($basPath)
            $null = $rt2Xl.Run("$($rt2Comp.Name).CanonicalizeFrx", $rtFrxOut)
        } finally {
            $rt2Wb.Close($false)
        }
    } finally {
        try { $rt2Xl.Quit() | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rt2Xl) | Out-Null } catch {}
        [System.GC]::Collect()
        if ($rt2Pid) {
            Start-Sleep -Milliseconds 300
            try { Stop-Process -Id $rt2Pid -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    $rtCmp = Compare-Bytes $canonA $rtFrxOut
    if ($rtCmp.Equal) {
        Write-Host "    PASS: round-trip canon == original canon" -ForegroundColor Green
    } else {
        Write-Host "    FAIL: round-trip canon differs ($($rtCmp.Diffs) bytes)" -ForegroundColor Red
        $aBytes = [System.IO.File]::ReadAllBytes($canonA)
        $bBytes = [System.IO.File]::ReadAllBytes($rtFrxOut)
        $shown = 0
        for ($i = 0; $i -lt [Math]::Min($aBytes.Length, $bBytes.Length) -and $shown -lt 20; $i++) {
            if ($aBytes[$i] -ne $bBytes[$i]) {
                Write-Host ("      diff @ 0x{0:X8} ({0,5}): canon=0x{1:X2} rt=0x{2:X2}" -f $i, $aBytes[$i], $bBytes[$i])
                $shown++
            }
        }
    }
    exit 0
} else {
    Write-Host ""
    Write-Host "FAIL: canon(A) != canon(B)  ($($postCmp.Diffs) byte diffs remain)" -ForegroundColor Red
    # Show first few diffs
    $a = [System.IO.File]::ReadAllBytes($canonA)
    $b = [System.IO.File]::ReadAllBytes($canonB)
    $shown = 0
    for ($i = 0; $i -lt $a.Length -and $shown -lt 20; $i++) {
        if ($a[$i] -ne $b[$i]) {
            Write-Host ("    diff @ 0x{0:X8} ({0,5}): A=0x{1:X2} B=0x{2:X2}" -f $i, $a[$i], $b[$i])
            $shown++
        }
    }
    exit 1
}
