# Run every test-*.ps1 in this directory serially. Exit code = fail count.

[CmdletBinding()]
param()

$here = $PSScriptRoot
$tests = Get-ChildItem -LiteralPath $here -Filter 'test-*.ps1' | Sort-Object Name

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 200

$results = @()
foreach ($t in $tests) {
    Write-Host ""
    Write-Host "=== $($t.Name) ===" -ForegroundColor Cyan
    & $t.FullName | Out-Host
    $status = if ($LASTEXITCODE -eq 0) { 'PASS' } else { 'FAIL' }
    $results += [pscustomobject]@{ Test = $t.Name; Status = $status; Exit = $LASTEXITCODE }
    # Cooldown — Excel COM cleanup needs a moment between back-to-back
    # sessions or subsequent ones get cleanup hangs.
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
}

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize
$failed = @($results | Where-Object { $_.Status -ne 'PASS' }).Count
$passed = $results.Count - $failed
Write-Host ("Passed: $passed / $($results.Count)") -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })

exit $failed
