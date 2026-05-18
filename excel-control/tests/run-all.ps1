# Run every test-*.ps1 in this directory serially. Exit code = fail count.

[CmdletBinding()]
param()

$here = $PSScriptRoot
$tests = Get-ChildItem -LiteralPath $here -Filter 'test-*.ps1' | Sort-Object Name

. (Join-Path $here '_helpers.ps1')
# Clean up orphans from any prior crashed test runs (PID-targeted —
# never touches the user's interactive Excel)
Clear-OrphanSessionExcels (Join-Path (Split-Path $here -Parent) 'sessions')
Start-Sleep -Milliseconds 200

$results = @()
foreach ($t in $tests) {
    Write-Host ""
    Write-Host "=== $($t.Name) ===" -ForegroundColor Cyan
    & $t.FullName | Out-Host
    $status = if ($LASTEXITCODE -eq 0) { 'PASS' } else { 'FAIL' }
    $results += [pscustomobject]@{ Test = $t.Name; Status = $status; Exit = $LASTEXITCODE }
    # Inter-test cooldown — each test's own finally block already killed
    # ITS spawned EXCEL.EXE (PID-targeted via Stop-SessionProcesses); we
    # just give COM a moment to settle between back-to-back sessions.
    Start-Sleep -Milliseconds 800
}

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize
$failed = @($results | Where-Object { $_.Status -ne 'PASS' }).Count
$passed = $results.Count - $failed
Write-Host ("Passed: $passed / $($results.Count)") -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })

exit $failed
