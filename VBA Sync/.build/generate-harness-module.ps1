# Generates `VBA Sync/Modules/modHarnessFiles.bas` by reading the
# excel-control harness source files and emitting them as VBA string
# functions. The .xlam imports the generated module so all harness
# files ship inside the .xlam binary — no separate dev tree required
# at user install time.
#
# Why each file becomes a Function (not a Const):
#   - VBA string literals are capped at 1023 chars per line
#   - A `Public Const X As String = "...long..."` would overflow on
#     anything bigger than a few short paragraphs
#   - A function builds the string line-by-line via repeated `s = s & ...`
#     which has no length cap and is what generated VBA tooling
#     conventionally produces for bundled text
#
# Run this whenever any source under excel-control/ changes:
#   pwsh .\VBA Sync\.build\generate-harness-module.ps1
# Then rebuild VBA Sync.xlam by importing the updated .bas files.

[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'

# Resolve repo root from this script's location, regardless of where
# the dev tree lives (Arnaud has it inside %APPDATA%\Microsoft\AddIns\,
# unusual but supported).
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..') |
    Select-Object -ExpandProperty Path
$srcRoot  = Join-Path $repoRoot 'excel-control'
$outPath  = Join-Path $repoRoot 'VBA Sync\Modules\modHarnessFiles.bas'

# (source-relative-path → VBA function name) mapping. Function names
# follow approved verbs (Get-*) so no PSScriptAnalyzer / VBA verb
# warnings.
$files = @(
    @{ Rel = 'start-session.ps1';                       Fn = 'GetStartSessionPs1' },
    @{ Rel = 'session-dialog-watcher.ps1';              Fn = 'GetSessionDialogWatcherPs1' },
    @{ Rel = 'capture.ps1';                             Fn = 'GetCapturePs1' },
    @{ Rel = 'tools\INTERFACE.md';                      Fn = 'GetInterfaceMd' },
    @{ Rel = 'tools\clsAssert.cls';                     Fn = 'GetClsAssertCls' },
    @{ Rel = 'tools\.claude\settings.json';             Fn = 'GetToolsClaudeSettingsJson' },
    @{ Rel = 'skill\SKILL.md';                          Fn = 'GetSkillMd' }
    # Note: the MS-OFORMS .frx canonicalizer is NOT embedded here.
    # It runs in pure VBA inside the xlam (modFrxCanonicalize.bas),
    # invoked directly by modSync.ExportComponent — no PS shell-out
    # at runtime. The PS port in excel-control/canonicalize/ stays in
    # the source tree as a reference / second implementation but is
    # not bundled into end-user workbooks.
)

# Escape a single chunk of source for use inside a VBA double-quoted
# string literal. Only `"` needs escaping (doubled). VBA strings can
# contain anything else as-is including tabs and unicode.
function Format-VbaStringLiteral([string]$Chunk) {
    return '"' + ($Chunk -replace '"', '""') + '"'
}

# VBA caps each PHYSICAL LINE at 1023 chars total. The line we emit is
# `    s = s & "<literal>" & vbLf` (~22 chars of overhead). Capping the
# literal at 800 chars per emit leaves ample headroom even after escaping
# embedded quotes (which double in length).
$VBA_MAX_LITERAL_CHARS = 800

# Emit one source line as one or more `    s = s & "..."` statements.
# - If the line fits in one literal: emit `s = s & "<lit>" & vbLf`.
# - If it overflows: emit several `s = s & "<chunkN>"` lines, with the
#   final chunk getting the `& vbLf` suffix. Each statement is its own
#   physical line so VBA's 1023-char limit applies per statement.
function Emit-LineStatements {
    param(
        [System.Text.StringBuilder]$Sb,
        [string]$Line,
        [bool]$AppendNewline
    )
    if ($Line.Length -le $VBA_MAX_LITERAL_CHARS) {
        $lit = Format-VbaStringLiteral $Line
        $tail = if ($AppendNewline) { ' & vbLf' } else { '' }
        [void]$Sb.AppendLine("    s = s & $lit$tail")
        return
    }
    # Long line: split into chunks. Concat across multiple statements.
    $pos = 0
    while ($pos -lt $Line.Length) {
        $take = [Math]::Min($VBA_MAX_LITERAL_CHARS, $Line.Length - $pos)
        $chunk = $Line.Substring($pos, $take)
        $pos += $take
        $isLastChunk = ($pos -ge $Line.Length)
        $lit = Format-VbaStringLiteral $chunk
        if ($isLastChunk -and $AppendNewline) {
            [void]$Sb.AppendLine("    s = s & $lit & vbLf")
        } else {
            [void]$Sb.AppendLine("    s = s & $lit")
        }
    }
}

# Build one Function body for a single file.
function Build-Function {
    param([string]$Rel, [string]$Fn, [string]$Content)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("' Source: excel-control\$Rel")
    [void]$sb.AppendLine("Public Function $Fn() As String")
    [void]$sb.AppendLine('    Dim s As String')
    # Normalize line endings then split on LF. `-split "`r?`n", -1`
    # was the original here but PowerShell treats the trailing `-1` as
    # an options enum, not a max-substring count, so the split silently
    # returned the entire content as a single element. Normalize then
    # plain .Split() instead -- no surprises.
    $lines = $Content.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n")
    $lastIdx = $lines.Count - 1
    for ($i = 0; $i -le $lastIdx; $i++) {
        $line = $lines[$i]
        if ($i -lt $lastIdx) {
            Emit-LineStatements -Sb $sb -Line $line -AppendNewline $true
        } else {
            # Last line: trailing newline only if source ended with one.
            # Our split would have produced an empty final element in that
            # case; skip emitting anything for the synthetic empty.
            if ($line.Length -gt 0) {
                Emit-LineStatements -Sb $sb -Line $line -AppendNewline $false
            }
        }
    }
    [void]$sb.AppendLine("    $Fn = s")
    [void]$sb.AppendLine('End Function')
    return $sb.ToString()
}

# Header
$out = New-Object System.Text.StringBuilder
[void]$out.AppendLine('Attribute VB_Name = "modHarnessFiles"')
[void]$out.AppendLine('Option Explicit')
[void]$out.AppendLine('')
[void]$out.AppendLine("' AUTO-GENERATED by VBA Sync\.build\generate-harness-module.ps1")
[void]$out.AppendLine("' DO NOT EDIT BY HAND. Regenerate after any change to excel-control/.")
[void]$out.AppendLine("'")
[void]$out.AppendLine("' Each function returns the verbatim content of an excel-control source")
[void]$out.AppendLine("' file. modSync.WriteHarness reads them and writes the files into the")
[void]$out.AppendLine("' user's exported workbook. The .xlam is thus self-contained: an end user")
[void]$out.AppendLine("' who downloads only VBA Sync.xlam gets the full agent harness on Export.")
[void]$out.AppendLine('')

$missing = New-Object System.Collections.ArrayList
foreach ($f in $files) {
    $src = Join-Path $srcRoot $f.Rel
    if (-not (Test-Path -LiteralPath $src)) {
        [void]$missing.Add($f.Rel)
        continue
    }
    $content = [System.IO.File]::ReadAllText($src)
    Write-Host "  embedding $($f.Rel) ($($content.Length) chars) -> $($f.Fn)()" -ForegroundColor DarkGray
    $body = Build-Function -Rel $f.Rel -Fn $f.Fn -Content $content
    [void]$out.AppendLine($body)
}

if ($missing.Count -gt 0) {
    Write-Host "MISSING SOURCES (skipped):" -ForegroundColor Yellow
    foreach ($m in $missing) { Write-Host "  $m" -ForegroundColor Yellow }
    throw "One or more harness source files missing — aborting before producing an incomplete module"
}

[System.IO.File]::WriteAllText($outPath, $out.ToString())
$size = (Get-Item -LiteralPath $outPath).Length
Write-Host "" -ForegroundColor Green
Write-Host "Wrote $outPath ($size bytes)" -ForegroundColor Green
Write-Host "Next: rebuild VBA Sync.xlam to bake in the updated modHarnessFiles." -ForegroundColor Cyan
