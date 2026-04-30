<#
.SYNOPSIS
    Normalises Excel OOXML for clean Git diffs. Called by VBA Sync's export step.

.DESCRIPTION
    Reads xl/workbook.xml, xl/_rels/workbook.xml.rels, xl/worksheets/*.xml and
    xl/tables/*.xml from a temp-extracted .xlsx, and writes:

      MANIFEST.json                   - workbook structure as JSON (sheets, defined
                                        names, lambdas, calc settings, code names)
      lambdas/<Name>.lambda           - one file per unique LAMBDA defined name,
                                        deduped across local-sheet copies
      worksheets/<NN> - <Name>.xml    - per-sheet, sheetId-prefixed for stable sort
                                        and rename-survivable, multi-line indented,
                                        volatile metadata stripped
      tables/<TableName>.xml          - per-table, named after the table not its rId
      .normalize.log                  - per-run log; truncated each run; VBA reads
                                        this to surface warnings/errors

    All output is UTF-8 without BOM. JSON uses LF endings; XML uses CRLF.
    Idempotent: same input always yields byte-identical output.

.PARAMETER Source
    Path to a folder produced by Expand-Archive on an .xlsx/.xlsm. The script
    reads the 'xl' subtree under this path.

.PARAMETER Destination
    Path to write the normalised output. Created if missing. The script does
    NOT delete pre-existing files here; that is the caller's responsibility
    (VBA Sync's PruneStaleFiles handles it).

.EXAMPLE
    powershell -File Normalize-ExcelXml.ps1 -Source C:\tmp\extract -Destination C:\tmp\out

.NOTES
    Targets Windows PowerShell 5.1 (no PS7 features, no module installs).
    Exit codes: 0 = success, 1 = unrecoverable startup error, 2 = some files failed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Destination
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Xml.Linq

# ---- constants -------------------------------------------------------------

$script:LogPath = $null
$script:FailedCount = 0
$script:RedactedSheetProtection = New-Object System.Collections.Generic.List[string]
$script:RedactedWorkbookProtection = $false
$script:RedactedFileSharing = $false

# Namespaces we care about by URI (so we are not coupled to Excel's prefix choice)
$NS = @{
    main      = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    r         = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel       = 'http://schemas.openxmlformats.org/package/2006/relationships'
    mc        = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
}

# OOXML extension namespaces that are *always* safe to keep when not in mc:Ignorable.
# These can carry meaningful content (extended conditional formats, drawings, slicers)
# so we never strip them wholesale even if Excel happens to list them in mc:Ignorable.
$NS_KEEP = @(
    'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',           # x14
    'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',              # x14ac (kept; absPath stripped by name)
    'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',           # x15
    'http://schemas.microsoft.com/office/drawing/2014/main',                    # xdr14
    'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'       # xdr
)

# Attributes to strip by qualified name regardless of where they appear.
# Format: '{namespaceUri}localName'  or  'localName' (no namespace)
#
# CAREFUL: anything in this list gets stripped from EVERY element. If an
# attribute is sometimes meaningful (sqref on <dataValidation>, etc), don't
# put it here -- handle the volatile case by stripping the parent element
# wholesale via VOLATILE_ELEMENTS instead.
$VOLATILE_ATTRS = @(
    # workbook.xml
    'lastEdited', 'lowestEdited', 'rupBuild',
    'codeName',  # only on <fileVersion> -- we re-add stable codeNames from <workbookPr> + <sheetPr>
    'calcId',
    'defaultThemeVersion',
    # x15ac:absPath leaks the local filesystem path
    '{http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac}absPath',
    '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}absPath',
    # workbookView window state
    'xWindow', 'yWindow', 'windowWidth', 'windowHeight', 'tabRatio',
    'firstSheet', 'activeTab',
    # sheetView UI state
    'tabSelected', 'zoomScale', 'zoomScaleNormal', 'workbookViewId',
    # <pane> scroll position (xSplit/ySplit/state for frozen panes are KEPT)
    'topLeftCell'
    # NOTE: sqref/activeCell/pane/activePane removed -- those only appeared
    # inside <selection>, which is wholesale-stripped by VOLATILE_ELEMENTS.
    # Keeping them here would corrupt <dataValidation sqref="..."> and
    # <conditionalFormatting sqref="...">.
)

# Elements to strip wholesale by qualified name
$VOLATILE_ELEMENTS = @(
    "{$($NS.main)}fileSharing",            # password hash + username -- redacted, logged
    # workbookProtection and sheetProtection are handled specially (log + strip attrs, keep element)
    # selection is stripped wherever it appears
    "{$($NS.main)}selection"
)

# Elements whose attributes we keep but where we strip the protection-secret attrs
$PROTECTION_STRIP_ATTRS = @(
    'algorithmName', 'hashValue', 'saltValue', 'spinCount',
    'workbookAlgorithmName', 'workbookHashValue', 'workbookSaltValue', 'workbookSpinCount',
    'revisionsAlgorithmName', 'revisionsHashValue', 'revisionsSaltValue', 'revisionsSpinCount'
)

# ---- IO helpers ------------------------------------------------------------

function Write-Utf8NoBom {
    param([string]$Path, [string]$Text, [string]$LineEnding = "`r`n")
    $normalized = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    if ($LineEnding -ne "`n") { $normalized = $normalized -replace "`n", $LineEnding }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $normalized, $enc)
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -Force
    }
}

function Write-NormalizeLog {
    param(
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level,
        [string]$Message
    )
    $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $line = "$ts [$Level] $Message"
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
    if ($Level -ne 'INFO') {
        [Console]::Error.WriteLine($line)
    }
}

function Get-SafeFileName {
    # Percent-encode the 9 illegal Windows filename chars (and % itself, as the
    # escape). Reversible. Sheet names have no other constraints in Excel.
    param([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return '_' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Name.ToCharArray()) {
        $needsEscape = $c -eq '%' -or $c -eq '/' -or $c -eq '\' -or $c -eq ':' -or `
                       $c -eq '*' -or $c -eq '?' -or $c -eq '"' -or $c -eq '<' -or `
                       $c -eq '>' -or $c -eq '|'
        if ($needsEscape) {
            [void]$sb.Append('%')
            [void]$sb.Append([Convert]::ToString([int]$c, 16).ToUpperInvariant().PadLeft(2, '0'))
        } else {
            [void]$sb.Append($c)
        }
    }
    # Trim trailing dots/spaces (illegal on Windows)
    $result = $sb.ToString().TrimEnd(' ', '.')
    if ($result.Length -gt 100) { $result = $result.Substring(0, 100) }
    if ([string]::IsNullOrEmpty($result)) { return '_' }
    return $result
}

# ---- XML pipeline ----------------------------------------------------------

function Load-Xml {
    param([string]$Path)
    return [System.Xml.Linq.XDocument]::Load($Path, [System.Xml.Linq.LoadOptions]::PreserveWhitespace)
}

# ---- Cell reference helpers -----------------------------------------------
# Parse "B7" -> @{Col=2; Row=7}. Returns $null on malformed input.
function ConvertFrom-A1Ref {
    param([string]$Ref)
    if ([string]::IsNullOrEmpty($Ref)) { return $null }
    $m = [regex]::Match($Ref, '^\$?([A-Za-z]+)\$?(\d+)$')
    if (-not $m.Success) { return $null }
    $colLetters = $m.Groups[1].Value.ToUpperInvariant()
    $rowNum = [int]$m.Groups[2].Value
    $colNum = 0
    foreach ($c in $colLetters.ToCharArray()) {
        $colNum = $colNum * 26 + ([int]$c - 64)
    }
    return @{ Col = $colNum; Row = $rowNum }
}

# Convert column number (1-based) to letters: 1->A, 26->Z, 27->AA, 702->ZZ.
function ConvertTo-ColLetters {
    param([int]$Col)
    $s = ''
    while ($Col -gt 0) {
        $r = ($Col - 1) % 26
        $s = [char]([int]'A'[0] + $r) + $s
        $Col = [Math]::Floor(($Col - 1) / 26)
    }
    return $s
}

# Build A1 ref from (col, row) — "5,2" -> "E2"
function ConvertTo-A1Ref {
    param([int]$Col, [int]$Row)
    return (ConvertTo-ColLetters -Col $Col) + [string]$Row
}

# Parse a range "A1:C5" into start/end (col,row). Returns $null on malformed.
function ConvertFrom-A1Range {
    param([string]$Range)
    if (-not $Range) { return $null }
    $parts = $Range.Split(':')
    if ($parts.Count -eq 1) {
        $a = ConvertFrom-A1Ref -Ref $parts[0]
        if (-not $a) { return $null }
        return @{ StartCol = $a.Col; StartRow = $a.Row; EndCol = $a.Col; EndRow = $a.Row }
    }
    if ($parts.Count -ne 2) { return $null }
    $a = ConvertFrom-A1Ref -Ref $parts[0]
    $b = ConvertFrom-A1Ref -Ref $parts[1]
    if (-not $a -or -not $b) { return $null }
    return @{
        StartCol = [Math]::Min($a.Col, $b.Col)
        StartRow = [Math]::Min($a.Row, $b.Row)
        EndCol   = [Math]::Max($a.Col, $b.Col)
        EndRow   = [Math]::Max($a.Row, $b.Row)
    }
}

# ---- SharedStrings + styles loaders ---------------------------------------
# Both indexed-by-position lookups. Loaded once per workbook.

function Load-SharedStrings {
    # Returns a string array where index = sharedString id, value = string content.
    # Returns @() if sharedStrings.xml doesn't exist.
    param([string]$SourceXlPath)
    $sstPath = Join-Path $SourceXlPath 'sharedStrings.xml'
    if (-not (Test-Path -LiteralPath $sstPath)) { return @() }
    $doc = Load-Xml -Path $sstPath
    $strings = New-Object System.Collections.Generic.List[string]
    foreach ($si in $doc.Descendants([System.Xml.Linq.XName]::Get('si', $NS.main))) {
        # <si> can contain a single <t> OR multiple <r><t>...</t></r> runs (rich text).
        # We want the concatenated text content either way.
        $tElements = @($si.Descendants([System.Xml.Linq.XName]::Get('t', $NS.main)))
        $text = ($tElements | ForEach-Object { $_.Value }) -join ''
        $strings.Add($text)
    }
    return ,$strings.ToArray()
}

function Load-Styles {
    # Returns a hashtable with:
    #   CellXfs       -> array of cellXf records (style index -> resolved attrs)
    #   Dxfs          -> array of dxf records (dxfId -> resolved attrs)
    #   NumberFormats -> hashtable numFmtId -> formatCode
    #   Fonts         -> array of font records
    #   Fills         -> array of fill records (pattern + colour)
    #   Borders       -> array of border records
    # Returns $null if styles.xml missing.
    param([string]$SourceXlPath)
    $stylesPath = Join-Path $SourceXlPath 'styles.xml'
    if (-not (Test-Path -LiteralPath $stylesPath)) { return $null }
    $doc = Load-Xml -Path $stylesPath

    # numFmts: built-in IDs (0-49) are predefined; custom IDs (>= 164) are in the file
    $numFmts = @{}
    $builtIns = @{
        0='General'; 1='0'; 2='0.00'; 3='#,##0'; 4='#,##0.00';
        9='0%'; 10='0.00%'; 11='0.00E+00'; 12='# ?/?'; 13='# ??/??';
        14='m/d/yyyy'; 15='d-mmm-yy'; 16='d-mmm'; 17='mmm-yy'; 18='h:mm AM/PM';
        19='h:mm:ss AM/PM'; 20='h:mm'; 21='h:mm:ss'; 22='m/d/yyyy h:mm';
        37='#,##0 ;(#,##0)'; 38='#,##0 ;[Red](#,##0)'; 39='#,##0.00;(#,##0.00)';
        40='#,##0.00;[Red](#,##0.00)'; 45='mm:ss'; 46='[h]:mm:ss'; 47='mmss.0';
        48='##0.0E+0'; 49='@'
    }
    foreach ($k in $builtIns.Keys) { $numFmts[$k] = $builtIns[$k] }
    foreach ($nf in $doc.Descendants([System.Xml.Linq.XName]::Get('numFmt', $NS.main))) {
        $id = [int]$nf.Attribute('numFmtId').Value
        $numFmts[$id] = $nf.Attribute('formatCode').Value
    }

    # Fonts, fills, borders — capture as resolved hashtables
    $fonts = @()
    $fontsParent = $doc.Descendants([System.Xml.Linq.XName]::Get('fonts', $NS.main)) | Select-Object -First 1
    if ($fontsParent) {
        foreach ($f in $fontsParent.Elements([System.Xml.Linq.XName]::Get('font', $NS.main))) {
            $fonts += ConvertFrom-FontElement -Element $f
        }
    }

    $fills = @()
    $fillsParent = $doc.Descendants([System.Xml.Linq.XName]::Get('fills', $NS.main)) | Select-Object -First 1
    if ($fillsParent) {
        foreach ($fi in $fillsParent.Elements([System.Xml.Linq.XName]::Get('fill', $NS.main))) {
            $fills += ConvertFrom-FillElement -Element $fi
        }
    }

    $borders = @()
    $bordersParent = $doc.Descendants([System.Xml.Linq.XName]::Get('borders', $NS.main)) | Select-Object -First 1
    if ($bordersParent) {
        foreach ($bo in $bordersParent.Elements([System.Xml.Linq.XName]::Get('border', $NS.main))) {
            $borders += ConvertFrom-BorderElement -Element $bo
        }
    }

    # cellXfs (the s="N" attribute on <c> indexes into this)
    $cellXfs = @()
    $cellXfsParent = $doc.Descendants([System.Xml.Linq.XName]::Get('cellXfs', $NS.main)) | Select-Object -First 1
    if ($cellXfsParent) {
        foreach ($xf in $cellXfsParent.Elements([System.Xml.Linq.XName]::Get('xf', $NS.main))) {
            $cellXfs += @{
                NumFmtId = if ($xf.Attribute('numFmtId')) { [int]$xf.Attribute('numFmtId').Value } else { 0 }
                FontId   = if ($xf.Attribute('fontId'))   { [int]$xf.Attribute('fontId').Value }   else { 0 }
                FillId   = if ($xf.Attribute('fillId'))   { [int]$xf.Attribute('fillId').Value }   else { 0 }
                BorderId = if ($xf.Attribute('borderId')) { [int]$xf.Attribute('borderId').Value } else { 0 }
                ApplyNumberFormat = (($xf.Attribute('applyNumberFormat')) -and ($xf.Attribute('applyNumberFormat').Value -eq '1'))
                ApplyFont   = (($xf.Attribute('applyFont'))   -and ($xf.Attribute('applyFont').Value -eq '1'))
                ApplyFill   = (($xf.Attribute('applyFill'))   -and ($xf.Attribute('applyFill').Value -eq '1'))
                ApplyBorder = (($xf.Attribute('applyBorder')) -and ($xf.Attribute('applyBorder').Value -eq '1'))
            }
        }
    }

    # dxfs (table dataDxfId etc reference these — typically partial style overrides)
    $dxfs = @()
    $dxfsParent = $doc.Descendants([System.Xml.Linq.XName]::Get('dxfs', $NS.main)) | Select-Object -First 1
    if ($dxfsParent) {
        foreach ($dxf in $dxfsParent.Elements([System.Xml.Linq.XName]::Get('dxf', $NS.main))) {
            $dxfs += ConvertFrom-DxfElement -Element $dxf
        }
    }

    return @{
        CellXfs       = $cellXfs
        Dxfs          = $dxfs
        NumberFormats = $numFmts
        Fonts         = $fonts
        Fills         = $fills
        Borders       = $borders
    }
}

# ---- Style sub-element parsers --------------------------------------------

function ConvertFrom-FontElement {
    param([System.Xml.Linq.XElement]$Element)
    $r = [ordered]@{}
    $name = $Element.Element([System.Xml.Linq.XName]::Get('name', $NS.main))
    if ($name) { $r['name'] = $name.Attribute('val').Value }
    $sz = $Element.Element([System.Xml.Linq.XName]::Get('sz', $NS.main))
    if ($sz) { $r['size'] = [double]$sz.Attribute('val').Value }
    if ($Element.Element([System.Xml.Linq.XName]::Get('b', $NS.main)))      { $r['bold']      = $true }
    if ($Element.Element([System.Xml.Linq.XName]::Get('i', $NS.main)))      { $r['italic']    = $true }
    if ($Element.Element([System.Xml.Linq.XName]::Get('u', $NS.main)))      { $r['underline'] = $true }
    if ($Element.Element([System.Xml.Linq.XName]::Get('strike', $NS.main))) { $r['strike']    = $true }
    $color = $Element.Element([System.Xml.Linq.XName]::Get('color', $NS.main))
    if ($color) {
        if ($color.Attribute('rgb'))     { $r['color'] = '#' + $color.Attribute('rgb').Value.ToUpperInvariant().TrimStart('F').PadLeft(6, '0') }
        elseif ($color.Attribute('theme')) { $r['color'] = "theme:$($color.Attribute('theme').Value)" }
        elseif ($color.Attribute('indexed')) { $r['color'] = "indexed:$($color.Attribute('indexed').Value)" }
    }
    return $r
}

function ConvertFrom-FillElement {
    param([System.Xml.Linq.XElement]$Element)
    $r = [ordered]@{}
    $pf = $Element.Element([System.Xml.Linq.XName]::Get('patternFill', $NS.main))
    if ($pf) {
        $type = if ($pf.Attribute('patternType')) { $pf.Attribute('patternType').Value } else { 'none' }
        if ($type -ne 'none') {
            $r['pattern'] = $type
            $fg = $pf.Element([System.Xml.Linq.XName]::Get('fgColor', $NS.main))
            if ($fg -and $fg.Attribute('rgb')) {
                $r['fill'] = '#' + $fg.Attribute('rgb').Value.ToUpperInvariant().TrimStart('F').PadLeft(6, '0')
            }
        }
    }
    return $r
}

function ConvertFrom-BorderElement {
    param([System.Xml.Linq.XElement]$Element)
    $r = [ordered]@{}
    foreach ($side in @('left', 'right', 'top', 'bottom')) {
        $b = $Element.Element([System.Xml.Linq.XName]::Get($side, $NS.main))
        if ($b -and $b.Attribute('style')) {
            $r[$side] = [ordered]@{ style = $b.Attribute('style').Value }
            $col = $b.Element([System.Xml.Linq.XName]::Get('color', $NS.main))
            if ($col -and $col.Attribute('rgb')) {
                $r[$side]['color'] = '#' + $col.Attribute('rgb').Value.ToUpperInvariant().TrimStart('F').PadLeft(6, '0')
            }
        }
    }
    return $r
}

function ConvertFrom-DxfElement {
    # A dxf is a partial override — only the bits that differ from the base style.
    param([System.Xml.Linq.XElement]$Element)
    $r = [ordered]@{}
    $font = $Element.Element([System.Xml.Linq.XName]::Get('font', $NS.main))
    if ($font) { $r['font'] = ConvertFrom-FontElement -Element $font }
    $fill = $Element.Element([System.Xml.Linq.XName]::Get('fill', $NS.main))
    if ($fill) {
        $f = ConvertFrom-FillElement -Element $fill
        if ($f.Count -gt 0) { foreach ($k in $f.Keys) { $r[$k] = $f[$k] } }
    }
    $border = $Element.Element([System.Xml.Linq.XName]::Get('border', $NS.main))
    if ($border) {
        $b = ConvertFrom-BorderElement -Element $border
        if ($b.Count -gt 0) { $r['border'] = $b }
    }
    $numFmt = $Element.Element([System.Xml.Linq.XName]::Get('numFmt', $NS.main))
    if ($numFmt -and $numFmt.Attribute('formatCode')) {
        $r['numberFormat'] = $numFmt.Attribute('formatCode').Value
    }
    return $r
}

# Resolve a cell's s="N" style index into a flat hashtable of non-default attrs.
function Resolve-CellStyle {
    param([hashtable]$Styles, [int]$StyleIndex)
    if (-not $Styles -or $StyleIndex -lt 0 -or $StyleIndex -ge $Styles.CellXfs.Count) { return @{} }
    $xf = $Styles.CellXfs[$StyleIndex]
    $r = [ordered]@{}

    # Number format (skip if General / id 0)
    if ($xf.NumFmtId -ne 0 -and $Styles.NumberFormats.ContainsKey($xf.NumFmtId)) {
        $r['numberFormat'] = $Styles.NumberFormats[$xf.NumFmtId]
    }
    # Font (skip default font index 0)
    if ($xf.FontId -gt 0 -and $xf.FontId -lt $Styles.Fonts.Count) {
        $f = $Styles.Fonts[$xf.FontId]
        if ($f.Count -gt 0) { $r['font'] = $f }
    }
    # Fill (default fill is index 0; index 1 is also default-ish "gray125"; skip both)
    if ($xf.FillId -gt 1 -and $xf.FillId -lt $Styles.Fills.Count) {
        $fi = $Styles.Fills[$xf.FillId]
        if ($fi.Count -gt 0) { foreach ($k in $fi.Keys) { $r[$k] = $fi[$k] } }
    }
    # Border (default border is index 0; skip)
    if ($xf.BorderId -gt 0 -and $xf.BorderId -lt $Styles.Borders.Count) {
        $bo = $Styles.Borders[$xf.BorderId]
        if ($bo.Count -gt 0) { $r['border'] = $bo }
    }
    return $r
}

# ---- Range merging --------------------------------------------------------
# Given a hashtable mapping "col,row" -> value (any type, but compared by canonical
# JSON serialisation), produce a list of (range, value) where adjacent cells with
# equal values are merged into greedy rectangles.
#
# Algorithm: scan top-to-bottom, left-to-right. For each unvisited cell with
# a value, expand right while same value, then expand down while every cell in
# the row stripe matches. Mark visited, emit range. O(n) where n = cell count.
function Compress-CellsToRanges {
    param([hashtable]$Cells)
    if (-not $Cells -or $Cells.Count -eq 0) { return @() }

    # Build a parallel hashtable of canonical JSON strings for value comparison
    # (so we can compare nested hashtables/arrays as values without walking them
    # each time).
    $serialised = @{}
    foreach ($k in $Cells.Keys) {
        $v = $Cells[$k]
        $serialised[$k] = if ($v -is [string]) { $v } else { ConvertTo-Json $v -Depth 10 -Compress }
    }

    # Determine bounds and bucket cells by (row, col) parsed from "col,row" key
    $minRow = [int]::MaxValue; $maxRow = 0
    $minCol = [int]::MaxValue; $maxCol = 0
    foreach ($k in $Cells.Keys) {
        $parts = $k.Split(',')
        $c = [int]$parts[0]; $r = [int]$parts[1]
        if ($r -lt $minRow) { $minRow = $r }
        if ($r -gt $maxRow) { $maxRow = $r }
        if ($c -lt $minCol) { $minCol = $c }
        if ($c -gt $maxCol) { $maxCol = $c }
    }

    $visited = @{}
    $output = New-Object System.Collections.Generic.List[object]

    for ($r = $minRow; $r -le $maxRow; $r++) {
        for ($c = $minCol; $c -le $maxCol; $c++) {
            $key = "$c,$r"
            if ($visited.ContainsKey($key) -or -not $Cells.ContainsKey($key)) { continue }
            $val = $serialised[$key]

            # Extend right
            $endCol = $c
            while ($endCol + 1 -le $maxCol) {
                $nextKey = "$($endCol + 1),$r"
                if (-not $Cells.ContainsKey($nextKey) -or $visited.ContainsKey($nextKey) -or $serialised[$nextKey] -ne $val) { break }
                $endCol++
            }
            # Extend down: each candidate row must have all cells $c..$endCol matching
            $endRow = $r
            $canExtend = $true
            while ($endRow + 1 -le $maxRow -and $canExtend) {
                for ($cc = $c; $cc -le $endCol; $cc++) {
                    $nextKey = "$cc,$($endRow + 1)"
                    if (-not $Cells.ContainsKey($nextKey) -or $visited.ContainsKey($nextKey) -or $serialised[$nextKey] -ne $val) {
                        $canExtend = $false; break
                    }
                }
                if ($canExtend) { $endRow++ }
            }

            # Mark visited
            for ($rr = $r; $rr -le $endRow; $rr++) {
                for ($cc = $c; $cc -le $endCol; $cc++) {
                    $visited["$cc,$rr"] = $true
                }
            }

            $startRef = ConvertTo-A1Ref -Col $c -Row $r
            $endRef   = ConvertTo-A1Ref -Col $endCol -Row $endRow
            $rangeStr = if ($startRef -eq $endRef) { $startRef } else { "$startRef`:$endRef" }
            $output.Add(@{ Range = $rangeStr; Value = $Cells[$key] })
        }
    }
    return $output.ToArray()
}

# ---- TSV escaping ---------------------------------------------------------
# Escape a cell value for inclusion in a TSV. Handles \, \t, \n, \r.
# Empty/null becomes ''.
function Format-TsvCell {
    param($Value)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    if ($s.Length -eq 0) { return '' }
    # Escape backslash first, then tab/newline/carriage return
    $s = $s.Replace('\', '\\')
    $s = $s.Replace("`t", '\t')
    $s = $s.Replace("`n", '\n')
    $s = $s.Replace("`r", '\r')
    return $s
}

function Get-IgnorableNamespaces {
    # Resolve the prefixes listed in mc:Ignorable on the root element to their
    # namespace URIs, then drop the ones in our allowlist.
    param([System.Xml.Linq.XDocument]$Doc)
    $root = $Doc.Root
    $mcIgnorable = $root.Attribute([System.Xml.Linq.XName]::Get('Ignorable', $NS.mc))
    if (-not $mcIgnorable) { return @() }
    $prefixes = $mcIgnorable.Value -split '\s+'
    $uris = @()
    foreach ($prefix in $prefixes) {
        if (-not $prefix) { continue }
        $ns = $root.GetNamespaceOfPrefix($prefix)
        if (-not $ns) { continue }   # prefix listed but not declared on root -- skip
        $uri = $ns.NamespaceName
        if (-not $uri) { continue }
        if ($NS_KEEP -notcontains $uri) {
            $uris += $uri
        }
    }
    return $uris
}

function Remove-VolatileNamespaces {
    param([System.Xml.Linq.XDocument]$Doc)
    $ignorableUris = @(Get-IgnorableNamespaces -Doc $Doc)  # @() guards against PS5.1 unwrapping empty arrays to $null
    if ($ignorableUris.Count -eq 0) { return }

    # Remove elements in volatile namespaces (descendant-walk first, then attributes)
    $elementsToRemove = @($Doc.Descendants() | Where-Object {
        $ignorableUris -contains $_.Name.Namespace.NamespaceName
    })
    foreach ($el in $elementsToRemove) { $el.Remove() }

    # Remove attributes in volatile namespaces
    $attrsToRemove = New-Object System.Collections.Generic.List[System.Xml.Linq.XAttribute]
    foreach ($el in $Doc.Descendants()) {
        foreach ($attr in $el.Attributes()) {
            if ($ignorableUris -contains $attr.Name.Namespace.NamespaceName) {
                $attrsToRemove.Add($attr)
            }
        }
    }
    foreach ($attr in $attrsToRemove) { $attr.Remove() }

    # Remove the now-orphaned xmlns:xr=... declarations on the root, and update mc:Ignorable
    $rootDeclsToRemove = @($Doc.Root.Attributes() | Where-Object {
        $_.IsNamespaceDeclaration -and ($ignorableUris -contains $_.Value)
    })
    foreach ($attr in $rootDeclsToRemove) { $attr.Remove() }

    # Rewrite mc:Ignorable to drop the prefixes we removed
    $mcIgnorable = $Doc.Root.Attribute([System.Xml.Linq.XName]::Get('Ignorable', $NS.mc))
    if ($mcIgnorable) {
        $remaining = @($mcIgnorable.Value -split '\s+' | Where-Object {
            if (-not $_) { return $false }
            $ns = $Doc.Root.GetNamespaceOfPrefix($_)
            if (-not $ns) { return $false }
            return ($NS_KEEP -contains $ns.NamespaceName)
        })
        if ($remaining.Count -eq 0) {
            $mcIgnorable.Remove()
        } else {
            $mcIgnorable.Value = ($remaining -join ' ')
        }
    }
}

function Remove-VolatileAttributes {
    param([System.Xml.Linq.XDocument]$Doc)
    $toRemove = New-Object System.Collections.Generic.List[System.Xml.Linq.XAttribute]
    foreach ($el in $Doc.Descendants()) {
        foreach ($attr in $el.Attributes()) {
            $qname = if ($attr.Name.Namespace.NamespaceName) {
                "{$($attr.Name.Namespace.NamespaceName)}$($attr.Name.LocalName)"
            } else {
                $attr.Name.LocalName
            }
            $bareName = $attr.Name.LocalName

            # codeName is special: volatile on <fileVersion>, stable on <workbookPr>/<sheetPr>
            if ($bareName -eq 'codeName' -and $el.Name.LocalName -ne 'fileVersion') {
                continue
            }

            if (($VOLATILE_ATTRS -contains $qname) -or ($VOLATILE_ATTRS -contains $bareName)) {
                $toRemove.Add($attr)
            }
        }
    }
    foreach ($attr in $toRemove) { $attr.Remove() }
}

function Remove-VolatileElements {
    param([System.Xml.Linq.XDocument]$Doc, [string]$FileLabel)
    foreach ($qname in $VOLATILE_ELEMENTS) {
        $hits = @($Doc.Descendants([System.Xml.Linq.XName]::Get($qname)))
        if ($hits.Count -gt 0) {
            if ($qname -eq "{$($NS.main)}fileSharing") {
                $script:RedactedFileSharing = $true
                Write-NormalizeLog -Level INFO -Message "${FileLabel}: redacted <fileSharing> (workbook-open password hash)"
            }
            foreach ($hit in $hits) { $hit.Remove() }
        }
    }
}

function Strip-ProtectionSecrets {
    param([System.Xml.Linq.XDocument]$Doc, [string]$FileLabel)
    $protectionElems = @($Doc.Descendants() | Where-Object {
        $_.Name.LocalName -eq 'workbookProtection' -or $_.Name.LocalName -eq 'sheetProtection'
    })
    foreach ($el in $protectionElems) {
        $stripped = $false
        $attrsToRemove = @($el.Attributes() | Where-Object {
            $PROTECTION_STRIP_ATTRS -contains $_.Name.LocalName
        })
        if ($attrsToRemove.Count -gt 0) {
            foreach ($attr in $attrsToRemove) { $attr.Remove() }
            $stripped = $true
        }
        if ($stripped) {
            if ($el.Name.LocalName -eq 'workbookProtection') {
                $script:RedactedWorkbookProtection = $true
                Write-NormalizeLog -Level INFO -Message "${FileLabel}: redacted <workbookProtection> hash/salt"
            } else {
                $script:RedactedSheetProtection.Add($FileLabel)
                Write-NormalizeLog -Level INFO -Message "${FileLabel}: redacted <sheetProtection> hash/salt"
            }
        }
    }
}

function Sort-AttributesAlphabetically {
    # Recursive. Sort: namespace declarations first, then by (namespaceUri, localName).
    param([System.Xml.Linq.XElement]$Element)
    $attrs = @($Element.Attributes())
    if ($attrs.Count -gt 1) {
        $sorted = $attrs | Sort-Object -Property `
            @{ Expression = { -not $_.IsNamespaceDeclaration } }, `
            @{ Expression = { $_.Name.Namespace.NamespaceName } }, `
            @{ Expression = { $_.Name.LocalName } }
        $Element.RemoveAttributes()
        foreach ($attr in $sorted) { $Element.Add($attr) }
    }
    foreach ($child in $Element.Elements()) {
        Sort-AttributesAlphabetically -Element $child
    }
}

function Save-PrettyXml {
    param([System.Xml.Linq.XDocument]$Doc, [string]$Path)
    Sort-AttributesAlphabetically -Element $Doc.Root

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = '  '
    $settings.NewLineChars = "`r`n"
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
    $settings.OmitXmlDeclaration = $false
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)

    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $Doc.Save($writer)
    } finally {
        $writer.Dispose()
    }
}

# ---- Workbook + manifest --------------------------------------------------

function Read-WorkbookRels {
    # Returns a hashtable: rId -> relative target path (e.g. 'rId7' -> 'worksheets/sheet5.xml')
    param([string]$RelsPath)
    $map = @{}
    if (-not (Test-Path -LiteralPath $RelsPath)) { return $map }
    $relsDoc = Load-Xml -Path $RelsPath
    foreach ($rel in $relsDoc.Descendants([System.Xml.Linq.XName]::Get('Relationship', $NS.rel))) {
        $idAttr = $rel.Attribute([System.Xml.Linq.XName]::Get('Id'))
        $targetAttr = $rel.Attribute([System.Xml.Linq.XName]::Get('Target'))
        if ($idAttr -and $targetAttr) {
            $map[$idAttr.Value] = $targetAttr.Value
        }
    }
    return $map
}

# Read a worksheet's _rels file to find which tables (and other parts) it
# references. Returns a hashtable: rId -> @{Type='table|drawing|comments|...';
# Target='absolute xl-relative path'}.
function Read-WorksheetRels {
    param([string]$SourceXlPath, [string]$WorksheetSourceFile)
    # WorksheetSourceFile is like 'worksheets/sheet5.xml' (relative to xl/)
    $sheetName = [System.IO.Path]::GetFileName($WorksheetSourceFile)  # 'sheet5.xml'
    $relsPath = Join-Path $SourceXlPath "worksheets\_rels\$sheetName.rels"
    $map = @{}
    if (-not (Test-Path -LiteralPath $relsPath)) { return $map }
    $relsDoc = Load-Xml -Path $relsPath
    foreach ($rel in $relsDoc.Descendants([System.Xml.Linq.XName]::Get('Relationship', $NS.rel))) {
        $id = $rel.Attribute('Id').Value
        $type = $rel.Attribute('Type').Value
        $target = $rel.Attribute('Target').Value
        # Resolve relative paths like '../tables/table1.xml' to xl-relative form
        if ($target.StartsWith('../')) {
            $target = $target.Substring(3)
        } elseif (-not $target.Contains('/')) {
            $target = "worksheets/$target"
        }
        # Extract the short type name from the schema URL
        $shortType = $type
        if ($type.Contains('/')) { $shortType = $type.Substring($type.LastIndexOf('/') + 1) }
        $map[$id] = @{ Type = $shortType; Target = $target }
    }
    return $map
}

function Build-SheetIndex {
    # Reads the post-normalisation workbook XDocument and returns an array of
    # PSCustomObjects describing each sheet, in workbook tab order.
    param([System.Xml.Linq.XDocument]$WorkbookDoc, [hashtable]$RelsMap)

    $sheets = @()
    $sheetElements = @($WorkbookDoc.Descendants([System.Xml.Linq.XName]::Get('sheet', $NS.main)))
    foreach ($s in $sheetElements) {
        $rId = $s.Attribute([System.Xml.Linq.XName]::Get('id', $NS.r)).Value
        $target = $RelsMap[$rId]
        # target is e.g. 'worksheets/sheet5.xml' relative to the workbook file
        $sheets += [PSCustomObject]@{
            sheetId    = [int]$s.Attribute('sheetId').Value
            name       = $s.Attribute('name').Value
            state      = if ($s.Attribute('state')) { $s.Attribute('state').Value } else { 'visible' }
            rId        = $rId
            sourceFile = $target  # e.g. 'worksheets/sheet5.xml'
        }
    }
    return $sheets
}

function Annotate-SheetMetadata {
    # Walk each worksheet XML to pick up codeName + tabColor (these live in <sheetPr>,
    # not <sheet>). Mutates the sheet objects in place.
    param([array]$Sheets, [string]$SourceXlPath)
    foreach ($sheet in $Sheets) {
        $sheetPath = Join-Path $SourceXlPath $sheet.sourceFile
        if (-not (Test-Path -LiteralPath $sheetPath)) {
            $sheet | Add-Member -NotePropertyName codeName -NotePropertyValue $null
            $sheet | Add-Member -NotePropertyName tabColor -NotePropertyValue $null
            continue
        }
        try {
            $doc = Load-Xml -Path $sheetPath
            $sheetPr = $doc.Descendants([System.Xml.Linq.XName]::Get('sheetPr', $NS.main)) | Select-Object -First 1
            $codeName = $null
            $tabColor = $null
            if ($sheetPr) {
                if ($sheetPr.Attribute('codeName')) { $codeName = $sheetPr.Attribute('codeName').Value }
                $tc = $sheetPr.Element([System.Xml.Linq.XName]::Get('tabColor', $NS.main))
                if ($tc -and $tc.Attribute('rgb')) { $tabColor = '#' + $tc.Attribute('rgb').Value.Substring(2) }
            }
            $sheet | Add-Member -NotePropertyName codeName -NotePropertyValue $codeName
            $sheet | Add-Member -NotePropertyName tabColor -NotePropertyValue $tabColor
        } catch {
            Write-NormalizeLog -Level WARN -Message "$($sheet.sourceFile): could not read sheet metadata: $_"
            $sheet | Add-Member -NotePropertyName codeName -NotePropertyValue $null
            $sheet | Add-Member -NotePropertyName tabColor -NotePropertyValue $null
        }
    }
}

function Extract-Lambdas {
    # Returns a hashtable: lambdaName -> body. Removes the <definedName> entries
    # from $Doc and re-adds a single global representative per lambda body.
    # Local-sheet copies of an identical body collapse to one shared body.
    param([System.Xml.Linq.XDocument]$Doc)

    $lambdas = [ordered]@{}      # name -> body  (canonical order: alphabetical name)

    $definedNames = @($Doc.Descendants([System.Xml.Linq.XName]::Get('definedName', $NS.main)))
    foreach ($dn in $definedNames) {
        $body = $dn.Value
        if (-not $body) { continue }

        # Heuristic for "this is a lambda": body starts with _xlfn.LAMBDA(
        $isLambda = $body -match '^\s*_xlfn\.LAMBDA\s*\('
        if (-not $isLambda) { continue }

        $nameAttr = $dn.Attribute('name')
        if (-not $nameAttr) { continue }
        $name = $nameAttr.Value

        # Whitespace-normalised body for dedup comparison
        $normBody = ($body -replace '\s+', ' ').Trim()

        if (-not $lambdas.Contains($name)) {
            $lambdas[$name] = $body  # keep original whitespace in the body file
        } else {
            # Same name, second occurrence (e.g. localSheetId copy). Keep the first
            # body verbatim. If they differ after whitespace normalisation, log a warning.
            $existingNorm = ($lambdas[$name] -replace '\s+', ' ').Trim()
            if ($existingNorm -ne $normBody) {
                Write-NormalizeLog -Level WARN -Message "Lambda '$name' has divergent bodies across sheets; keeping first occurrence"
            }
        }
        # Remove this definedName from the workbook regardless
        $dn.Remove()
    }

    return $lambdas
}

function Build-DefinedNames {
    # After Extract-Lambdas has stripped the lambdas, collect remaining defined names.
    # Returns an [ordered] hashtable, sorted by (name, localSheetId) with globals first.
    # Local-sheet keys use the sheet *name* (e.g. "_FilterDatabase@TASKPRED") not
    # the raw localSheetId index, which is opaque (Excel's localSheetId is a
    # 0-based index into the sheets-in-tab-order list, not the sheetId).
    param([System.Xml.Linq.XDocument]$Doc, [array]$Sheets)

    $entries = @()
    foreach ($dn in $Doc.Descendants([System.Xml.Linq.XName]::Get('definedName', $NS.main))) {
        $name = $dn.Attribute('name').Value
        $localSheetId = if ($dn.Attribute('localSheetId')) { [int]$dn.Attribute('localSheetId').Value } else { $null }
        $entries += [PSCustomObject]@{
            name = $name
            localSheetId = $localSheetId
            value = $dn.Value
        }
    }

    $sorted = $entries | Sort-Object -Property `
        @{ Expression = { $_.name } }, `
        @{ Expression = { if ($null -eq $_.localSheetId) { -1 } else { $_.localSheetId } } }

    $out = [ordered]@{}
    foreach ($e in $sorted) {
        if ($null -eq $e.localSheetId) {
            $key = $e.name
        } else {
            # localSheetId is 0-based index into tab-order sheets list
            $sheetName = if ($Sheets -and $e.localSheetId -lt $Sheets.Count -and $e.localSheetId -ge 0) {
                $Sheets[$e.localSheetId].name
            } else {
                "sheet$($e.localSheetId)"  # fallback if the index is out of range
            }
            $key = "$($e.name)@$sheetName"
        }
        $out[$key] = $e.value
    }
    return $out
}

function Get-WorkbookProperties {
    param([System.Xml.Linq.XDocument]$Doc)
    $wbPr = $Doc.Descendants([System.Xml.Linq.XName]::Get('workbookPr', $NS.main)) | Select-Object -First 1
    $calcPr = $Doc.Descendants([System.Xml.Linq.XName]::Get('calcPr', $NS.main)) | Select-Object -First 1

    $props = [ordered]@{
        codeName = if ($wbPr -and $wbPr.Attribute('codeName')) { $wbPr.Attribute('codeName').Value } else { $null }
        date1904 = if ($wbPr -and $wbPr.Attribute('date1904')) { $wbPr.Attribute('date1904').Value -eq '1' -or $wbPr.Attribute('date1904').Value -eq 'true' } else { $false }
        calcMode = if ($calcPr -and $calcPr.Attribute('calcMode')) { $calcPr.Attribute('calcMode').Value } else { 'auto' }
    }
    return $props
}

# ---- Per-artefact processors ----------------------------------------------

function Process-Workbook {
    param([string]$SourceXlPath, [string]$DestPath)

    $workbookXml = Join-Path $SourceXlPath 'workbook.xml'
    $relsXml = Join-Path $SourceXlPath '_rels\workbook.xml.rels'

    if (-not (Test-Path -LiteralPath $workbookXml)) {
        throw "Source workbook.xml not found at: $workbookXml"
    }

    $doc = Load-Xml -Path $workbookXml
    $relsMap = Read-WorkbookRels -RelsPath $relsXml

    Remove-VolatileNamespaces -Doc $doc
    Remove-VolatileElements -Doc $doc -FileLabel 'workbook.xml'
    Strip-ProtectionSecrets -Doc $doc -FileLabel 'workbook.xml'
    Remove-VolatileAttributes -Doc $doc

    # Sheet index BEFORE we strip more (we need <sheet> elements to read names/ids)
    $sheets = Build-SheetIndex -WorkbookDoc $doc -RelsMap $relsMap
    Annotate-SheetMetadata -Sheets $sheets -SourceXlPath $SourceXlPath

    # Lambda extraction (mutates $doc)
    $lambdas = Extract-Lambdas -Doc $doc

    # Defined names (after lambda strip)
    $definedNames = Build-DefinedNames -Doc $doc -Sheets $sheets

    # Workbook properties
    $wbProps = Get-WorkbookProperties -Doc $doc

    # Build the manifest. The tables array stays empty here -- it gets populated
    # later by Process-Table calls in Invoke-Normalize, before Save-Manifest.
    $manifest = [ordered]@{
        schemaVersion    = 1
        generatorVersion = '0.2'
        workbook         = $wbProps
        sheets           = @()
        tables           = @()
        definedNames     = $definedNames
        lambdas          = @($lambdas.Keys | Sort-Object)
    }

    # Compute padding width: at least 2, more if any sheetId exceeds 99 (e.g.
    # P6 Tools has sheetIds up to 146). Lexicographic sort of zero-padded
    # numbers only works when every number has the same width.
    $maxSheetId = ($sheets | Measure-Object -Property sheetId -Maximum).Maximum
    $padWidth = [Math]::Max(2, ([string]$maxSheetId).Length)
    foreach ($s in $sheets) {
        $s | Add-Member -NotePropertyName padWidth -NotePropertyValue $padWidth -Force
        $safeName = Get-SafeFileName -Name $s.name
        $padded = ([string]$s.sheetId).PadLeft($padWidth, '0')
        $manifest.sheets += [ordered]@{
            sheetId    = $s.sheetId
            name       = $s.name
            codeName   = $s.codeName
            state      = $s.state
            tabColor   = $s.tabColor
            file       = "worksheets/$padded - $safeName.xml"
        }
    }

    return [PSCustomObject]@{
        Manifest    = $manifest
        Sheets      = $sheets
        Lambdas     = $lambdas
    }
}

function Save-Manifest {
    param([System.Collections.Specialized.OrderedDictionary]$Manifest, [string]$Path)
    # PS 5.1's ConvertTo-Json indents oddly (4 spaces + double-space after colon).
    # Use ConvertTo-Json then post-process to standard 2-space, single-space.
    $raw = $Manifest | ConvertTo-Json -Depth 20 -Compress:$false
    $pretty = Format-Json -Json $raw
    # Collapse empty containers: {\n  \n} -> {} and [\n  \n] -> []
    $pretty = [regex]::Replace($pretty, '\{\s*\n\s*\}', '{}')
    $pretty = [regex]::Replace($pretty, '\[\s*\n\s*\]', '[]')
    Write-Utf8NoBom -Path $Path -Text $pretty -LineEnding "`n"
}

function Format-Json {
    # Reformat a PS5.1 ConvertTo-Json output to canonical 2-space indent
    # with single-space after colon. Walks the string char-by-char respecting
    # quoted strings (no JSON parse needed -- ConvertTo-Json output is always
    # valid and balanced).
    param([string]$Json)
    $sb = New-Object System.Text.StringBuilder
    $depth = 0
    $inString = $false
    $escaped = $false
    $atLineStart = $false
    for ($i = 0; $i -lt $Json.Length; $i++) {
        $c = $Json[$i]
        if ($inString) {
            [void]$sb.Append($c)
            if ($escaped) { $escaped = $false }
            elseif ($c -eq '\') { $escaped = $true }
            elseif ($c -eq '"') { $inString = $false }
            continue
        }
        switch ($c) {
            '"' {
                $inString = $true
                [void]$sb.Append($c)
                $atLineStart = $false
            }
            '{' {
                $depth++
                [void]$sb.Append('{')
                [void]$sb.Append("`n")
                [void]$sb.Append(('  ' * $depth))
                $atLineStart = $true
            }
            '[' {
                $depth++
                [void]$sb.Append('[')
                [void]$sb.Append("`n")
                [void]$sb.Append(('  ' * $depth))
                $atLineStart = $true
            }
            '}' {
                $depth--
                [void]$sb.Append("`n")
                [void]$sb.Append(('  ' * $depth))
                [void]$sb.Append('}')
                $atLineStart = $false
            }
            ']' {
                $depth--
                [void]$sb.Append("`n")
                [void]$sb.Append(('  ' * $depth))
                [void]$sb.Append(']')
                $atLineStart = $false
            }
            ',' {
                [void]$sb.Append(',')
                [void]$sb.Append("`n")
                [void]$sb.Append(('  ' * $depth))
                $atLineStart = $true
            }
            ':' {
                [void]$sb.Append(': ')
                $atLineStart = $false
            }
            { $_ -eq ' ' -or $_ -eq "`t" -or $_ -eq "`r" -or $_ -eq "`n" } {
                # Skip whitespace from PS's formatter; we control our own
                if (-not $atLineStart) { }
            }
            default {
                [void]$sb.Append($c)
                $atLineStart = $false
            }
        }
    }
    return $sb.ToString()
}

function Save-Lambdas {
    param([System.Collections.Specialized.OrderedDictionary]$Lambdas, [string]$LambdasDir)
    if ($Lambdas.Count -eq 0) { return }
    Ensure-Directory -Path $LambdasDir
    foreach ($name in $Lambdas.Keys) {
        $body = $Lambdas[$name]
        $safe = Get-SafeFileName -Name $name
        $path = Join-Path $LambdasDir "$safe.lambda"
        Write-Utf8NoBom -Path $path -Text $body -LineEnding "`n"
    }
}

function Process-Worksheet {
    # v1.0: per-sheet folder with split files instead of one normalised XML.
    # The folder layout mirrors what we agreed in the spec:
    #   worksheets/<NN> - <Name>/
    #     data.tsv             - cached cell values for non-table cells, with
    #                            marker rows at table positions
    #     formulas.json        - cell formulas (A1 form, range-collapsed)
    #     styles.json          - resolved per-cell styles, range-merged
    #     tables/<TableName>/
    #       definition.json    - schema, columns, calc formulas, overrides
    #       data.tsv           - clean dataset (headers row 1, data rows below)
    # Returns an array of table-manifest entries (consumed by Invoke-Normalize
    # to populate MANIFEST.tables).
    param(
        [string]$SourceXmlPath,
        [PSCustomObject]$SheetMeta,
        [string]$DestDir,
        [string]$SourceXlPath,
        [string[]]$SharedStrings,
        [hashtable]$Styles
    )

    $doc = Load-Xml -Path $SourceXmlPath
    Remove-VolatileNamespaces -Doc $doc
    $fileLabel = "worksheets/$($SheetMeta.name)"
    Remove-VolatileElements -Doc $doc -FileLabel $fileLabel
    Strip-ProtectionSecrets -Doc $doc -FileLabel $fileLabel
    Remove-VolatileAttributes -Doc $doc

    $padding = if ($SheetMeta.PSObject.Properties['padWidth']) { $SheetMeta.padWidth } else { 2 }
    $safeName = Get-SafeFileName -Name $SheetMeta.name
    $padded = ([string]$SheetMeta.sheetId).PadLeft($padding, '0')
    $sheetFolder = Join-Path $DestDir "$padded - $safeName"
    Ensure-Directory -Path $sheetFolder

    # Discover tables hosted on this sheet: walk <tableParts> for r:id refs,
    # resolve each via the worksheet's _rels file to a tableN.xml path, parse
    # each table to learn its name+ref so we can exclude its cells from the
    # sheet's data.tsv and emit a marker row at the table's start.
    $tableEntries = New-Object System.Collections.Generic.List[object]
    $tableRanges  = New-Object System.Collections.Generic.List[object]   # @{ Name; Range = @{StartCol,StartRow,EndCol,EndRow}; CellsByKey }
    $tablePartsEl = $doc.Descendants([System.Xml.Linq.XName]::Get('tableParts', $NS.main)) | Select-Object -First 1
    if ($tablePartsEl -and $SourceXlPath) {
        $rels = Read-WorksheetRels -SourceXlPath $SourceXlPath -WorksheetSourceFile $SheetMeta.sourceFile
        foreach ($tp in $tablePartsEl.Elements([System.Xml.Linq.XName]::Get('tablePart', $NS.main))) {
            $rIdAttr = $tp.Attribute([System.Xml.Linq.XName]::Get('id', $NS.r))
            if (-not $rIdAttr) { continue }
            $rel = $rels[$rIdAttr.Value]
            if (-not $rel) { continue }
            $tableXmlPath = Join-Path $SourceXlPath ($rel.Target -replace '/', '\')
            if (-not (Test-Path -LiteralPath $tableXmlPath)) { continue }

            try {
                $tDoc = Load-Xml -Path $tableXmlPath
                Remove-VolatileNamespaces -Doc $tDoc
                Remove-VolatileAttributes -Doc $tDoc
                $tName = $tDoc.Root.Attribute('name').Value
                $tRef  = $tDoc.Root.Attribute('ref').Value
                $tableSafe = Get-SafeFileName -Name $tName

                $tableFolder = Join-Path (Join-Path $sheetFolder 'tables') $tableSafe
                Ensure-Directory -Path $tableFolder

                # Parse the range so we can exclude its cells from the sheet
                $rng = ConvertFrom-A1Range -Range $tRef
                if ($rng) {
                    $tableRanges.Add(@{
                        Name = $tName
                        SafeName = $tableSafe
                        Range = $rng
                        Folder = $tableFolder
                        XmlDoc = $tDoc
                        Ref = $tRef
                    })
                }

                $tableEntries.Add([ordered]@{
                    name        = $tName
                    displayName = if ($tDoc.Root.Attribute('displayName')) { $tDoc.Root.Attribute('displayName').Value } else { $tName }
                    sheet       = $SheetMeta.name
                    ref         = $tRef
                    folder      = "worksheets/$padded - $safeName/tables/$tableSafe/"
                })
            } catch {
                Write-NormalizeLog -Level WARN -Message "$($fileLabel): could not process table at $tableXmlPath - $_"
            }
        }
    }

    # Walk <sheetData> once and collect three parallel views: values, formulas, styles
    $sheetData = $doc.Descendants([System.Xml.Linq.XName]::Get('sheetData', $NS.main)) | Select-Object -First 1
    $cellValues   = @{}   # "col,row" -> resolved value (string or number)
    $cellFormulas = @{}   # "col,row" -> A1 formula string
    $cellStyles   = @{}   # "col,row" -> resolved style hashtable
    $sharedFormulasMaster = @{}   # si -> @{ Formula = '...'; Ref = 'A1:A100' }

    if ($sheetData) {
        foreach ($row in $sheetData.Elements([System.Xml.Linq.XName]::Get('row', $NS.main))) {
            foreach ($cell in $row.Elements([System.Xml.Linq.XName]::Get('c', $NS.main))) {
                $refAttr = $cell.Attribute('r')
                if (-not $refAttr) { continue }
                $coord = ConvertFrom-A1Ref -Ref $refAttr.Value
                if (-not $coord) { continue }
                $key = "$($coord.Col),$($coord.Row)"

                # Cell type: t="s" (sharedString), t="b" (boolean), t="str" (inline string from formula),
                #   t="inlineStr" (inline string), t="e" (error), default = number
                $type = if ($cell.Attribute('t')) { $cell.Attribute('t').Value } else { 'n' }

                # Value
                $vEl = $cell.Element([System.Xml.Linq.XName]::Get('v', $NS.main))
                $isEl = $cell.Element([System.Xml.Linq.XName]::Get('is', $NS.main))
                $rawValue = $null
                if ($vEl) { $rawValue = $vEl.Value }
                elseif ($isEl) {
                    # inline string: concat all <t> contents
                    $tList = @($isEl.Descendants([System.Xml.Linq.XName]::Get('t', $NS.main)))
                    $rawValue = ($tList | ForEach-Object { $_.Value }) -join ''
                }

                if ($null -ne $rawValue -and $rawValue -ne '') {
                    if ($type -eq 's') {
                        $idx = [int]$rawValue
                        if ($idx -ge 0 -and $idx -lt $SharedStrings.Count) {
                            $cellValues[$key] = $SharedStrings[$idx]
                        } else {
                            $cellValues[$key] = "[!sst-$idx]"
                        }
                    } elseif ($type -eq 'b') {
                        $cellValues[$key] = if ($rawValue -eq '1') { 'TRUE' } else { 'FALSE' }
                    } elseif ($type -eq 'e') {
                        $cellValues[$key] = $rawValue   # e.g. "#N/A", "#REF!"
                    } else {
                        # number, str, inlineStr -- emit as-is
                        $cellValues[$key] = $rawValue
                    }
                }

                # Formula
                $fEl = $cell.Element([System.Xml.Linq.XName]::Get('f', $NS.main))
                if ($fEl) {
                    $fType = if ($fEl.Attribute('t')) { $fEl.Attribute('t').Value } else { 'normal' }
                    $si    = if ($fEl.Attribute('si')) { [int]$fEl.Attribute('si').Value } else { -1 }
                    $body  = $fEl.Value

                    if ($fType -eq 'shared') {
                        if ($body) {
                            # Master formula for this si group
                            $ref = if ($fEl.Attribute('ref')) { $fEl.Attribute('ref').Value } else { $refAttr.Value }
                            $sharedFormulasMaster[$si] = @{ Formula = $body; Ref = $ref }
                            $cellFormulas[$key] = "=$body"
                        } else {
                            # Reference to a master we'll resolve in pass 2
                            $cellFormulas[$key] = "[!shared-$si]"
                        }
                    } elseif ($body) {
                        $cellFormulas[$key] = "=$body"
                    }
                }

                # Style
                $sAttr = $cell.Attribute('s')
                if ($sAttr -and $Styles) {
                    $resolved = Resolve-CellStyle -Styles $Styles -StyleIndex ([int]$sAttr.Value)
                    if ($resolved.Count -gt 0) {
                        $cellStyles[$key] = $resolved
                    }
                }
            }
        }

        # Pass 2: resolve shared-formula references. Excel only stores the body
        # on one cell of each shared-formula group; siblings just reference si=N.
        # We could R1C1-translate the master to each sibling's position, but
        # that's risky without a real formula parser. Instead emit "=[shared:N]"
        # markers so the diff stays readable, and add a "sharedFormulas" section
        # to formulas.json that gives the masters their own entries.
        $needsShared = @($cellFormulas.GetEnumerator() | Where-Object { $_.Value -like '`[!shared-*' })
        foreach ($e in $needsShared) {
            $si = [int]($e.Value -replace '\[!shared-', '' -replace '\]', '')
            if ($sharedFormulasMaster.ContainsKey($si)) {
                $cellFormulas[$e.Key] = "[shared:$si]"
            } else {
                # Master not seen -- shouldn't happen but fall back to empty
                $cellFormulas.Remove($e.Key)
            }
        }
    }

    # Phase C: extract sheet-level metadata before tables modify cellValues
    Save-MetaJson -Doc $doc -SheetMeta $SheetMeta -TableRanges $tableRanges -Path (Join-Path $sheetFolder '_meta.json')
    Save-ValidationsJson -Doc $doc -Path (Join-Path $sheetFolder 'validations.json')
    Save-ConditionalFormatsJson -Doc $doc -Path (Join-Path $sheetFolder 'conditional_formats.json')
    Save-CommentsJson -SourceXlPath $SourceXlPath -SheetMeta $SheetMeta -Path (Join-Path $sheetFolder 'comments.json')

    # Phase B: extract table cells from $cellValues and $cellFormulas, then
    # write per-table data.tsv + definition.json. Cells inside table ranges
    # are removed from the sheet's $cellValues so they don't appear twice.
    # A marker row is added at the table's first row so the sheet's data.tsv
    # still shows where the table sits positionally.
    $tableMarkers = @{}    # "col,row" -> marker text (single cell at table's start col,row)
    foreach ($t in $tableRanges) {
        # Carve out the table's cells
        $tableCellValues   = @{}
        $tableCellFormulas = @{}
        for ($r = $t.Range.StartRow; $r -le $t.Range.EndRow; $r++) {
            for ($c = $t.Range.StartCol; $c -le $t.Range.EndCol; $c++) {
                $k = "$c,$r"
                if ($cellValues.ContainsKey($k)) {
                    $tableCellValues[$k] = $cellValues[$k]
                    $cellValues.Remove($k)
                }
                if ($cellFormulas.ContainsKey($k)) {
                    $tableCellFormulas[$k] = $cellFormulas[$k]
                    $cellFormulas.Remove($k)
                }
            }
        }

        # Marker row at the table's start (first column of first row of table range)
        $markerKey = "$($t.Range.StartCol),$($t.Range.StartRow)"
        $tableMarkers[$markerKey] = "[table:$($t.Name) ref=$($t.Ref) -> tables/$($t.SafeName)/]"

        # Write the table's data.tsv (clean: header row from cells in startRow,
        # data rows below; coordinates rebased so that table's first row is row 1)
        Save-TableDataTsv -Cells $tableCellValues -TableRange $t.Range -Path (Join-Path $t.Folder 'data.tsv')

        # Write the table's definition.json (schema, columns, calc formulas,
        # overrides). Built from the table's XML doc plus the table-cell formulas.
        Save-TableDefinitionJson -TableXmlDoc $t.XmlDoc -TableRange $t.Range -CellFormulas $tableCellFormulas -Path (Join-Path $t.Folder 'definition.json')
    }

    # Merge marker rows into the sheet's cellValues so they appear at the right row
    foreach ($k in $tableMarkers.Keys) { $cellValues[$k] = $tableMarkers[$k] }

    # Write data.tsv
    Save-DataTsv -Cells $cellValues -Path (Join-Path $sheetFolder 'data.tsv')

    # Write formulas.json
    Save-FormulasJson -Cells $cellFormulas -SharedMasters $sharedFormulasMaster -Path (Join-Path $sheetFolder 'formulas.json')

    # Write styles.json
    Save-StylesJson -Cells $cellStyles -Path (Join-Path $sheetFolder 'styles.json')

    # Return table manifest entries to the caller
    return ,$tableEntries.ToArray()
}

# Write a table's clean data.tsv. Headers row comes from the cells in the
# table's first row (range.StartRow); data rows are everything below. Cell
# coordinates are rebased so the table's first column is "Col1" (we use the
# original sheet column letters in the header for traceability).
function Save-TableDataTsv {
    param([hashtable]$Cells, [hashtable]$TableRange, [string]$Path)

    $startCol = $TableRange.StartCol
    $endCol   = $TableRange.EndCol
    $startRow = $TableRange.StartRow
    $endRow   = $TableRange.EndRow

    # Group by row
    $byRow = @{}
    foreach ($k in $Cells.Keys) {
        $parts = $k.Split(',')
        $c = [int]$parts[0]; $r = [int]$parts[1]
        if (-not $byRow.ContainsKey($r)) { $byRow[$r] = @{} }
        $byRow[$r][$c] = $Cells[$k]
    }

    # Header: take values from the first row (table headers); fall back to
    # column-letter for any blank header cell
    $headers = @()
    if ($byRow.ContainsKey($startRow)) {
        for ($c = $startCol; $c -le $endCol; $c++) {
            $h = if ($byRow[$startRow].ContainsKey($c)) { [string]$byRow[$startRow][$c] } else { '' }
            if ([string]::IsNullOrEmpty($h)) { $h = ConvertTo-ColLetters -Col $c }
            $headers += $h
        }
    } else {
        for ($c = $startCol; $c -le $endCol; $c++) { $headers += (ConvertTo-ColLetters -Col $c) }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(($headers | ForEach-Object { Format-TsvCell -Value $_ }) -join "`t")
    [void]$sb.Append("`r`n")

    # Data rows (skip startRow which is the header row)
    for ($r = $startRow + 1; $r -le $endRow; $r++) {
        $line = @()
        for ($c = $startCol; $c -le $endCol; $c++) {
            $v = if ($byRow.ContainsKey($r) -and $byRow[$r].ContainsKey($c)) { $byRow[$r][$c] } else { '' }
            $line += (Format-TsvCell -Value $v)
        }
        [void]$sb.Append($line -join "`t")
        [void]$sb.Append("`r`n")
    }

    Write-Utf8NoBom -Path $Path -Text $sb.ToString() -LineEnding "`r`n"
}

# Write a table's definition.json. Schema, columns (with calc formulas where
# present), and overrides (cells in a calc column whose formula differs from
# the column default).
function Save-TableDefinitionJson {
    param(
        [System.Xml.Linq.XDocument]$TableXmlDoc,
        [hashtable]$TableRange,
        [hashtable]$CellFormulas,
        [string]$Path
    )

    $root = $TableXmlDoc.Root
    $name = $root.Attribute('name').Value
    $displayName = if ($root.Attribute('displayName')) { $root.Attribute('displayName').Value } else { $name }
    $totalsRowShown = if ($root.Attribute('totalsRowShown')) { ($root.Attribute('totalsRowShown').Value -ne '0') } else { $true }

    $columns = @()
    $tableColumnsEl = $root.Element([System.Xml.Linq.XName]::Get('tableColumns', $NS.main))
    $colIdx = 0
    if ($tableColumnsEl) {
        foreach ($tc in $tableColumnsEl.Elements([System.Xml.Linq.XName]::Get('tableColumn', $NS.main))) {
            $colName = $tc.Attribute('name').Value
            $col = [ordered]@{ name = $colName }
            $calcFormulaEl = $tc.Element([System.Xml.Linq.XName]::Get('calculatedColumnFormula', $NS.main))
            if ($calcFormulaEl) {
                $defaultFormula = "=$($calcFormulaEl.Value)"
                $col['formula'] = $defaultFormula

                # Find overrides: walk this column's data cells, flag any whose
                # formula differs from the default. Also flag cells that have
                # no formula at all (manually overridden to a literal).
                $sheetColIdx = $TableRange.StartCol + $colIdx
                $overrides = [ordered]@{}
                for ($r = $TableRange.StartRow + 1; $r -le $TableRange.EndRow; $r++) {
                    $k = "$sheetColIdx,$r"
                    $tableRowIdx = $r - $TableRange.StartRow   # 1-based table row
                    if ($CellFormulas.ContainsKey($k)) {
                        $cellF = $CellFormulas[$k]
                        # Resolved shared-formula references match the default
                        if ($cellF -ne $defaultFormula -and $cellF -notlike '*[shared:*]*') {
                            $overrides["$tableRowIdx"] = $cellF
                        }
                    }
                }
                if ($overrides.Count -gt 0) { $col['overrides'] = $overrides }
            }
            $columns += $col
            $colIdx++
        }
    }

    $obj = [ordered]@{
        name           = $name
        displayName    = $displayName
        ref            = $TableRange.StartCol.ToString() + ',' + $TableRange.StartRow.ToString() + ':' + $TableRange.EndCol.ToString() + ',' + $TableRange.EndRow.ToString()
        a1Ref          = (ConvertTo-A1Ref -Col $TableRange.StartCol -Row $TableRange.StartRow) + ':' + (ConvertTo-A1Ref -Col $TableRange.EndCol -Row $TableRange.EndRow)
        totalsRowShown = $totalsRowShown
        columns        = $columns
    }

    # Replace the long winded "ref" with just a1Ref; ref above was just for completeness
    $obj.Remove('ref')
    $obj['ref'] = $obj['a1Ref']
    $obj.Remove('a1Ref')

    Save-Manifest -Manifest $obj -Path $Path
}

function Save-DataTsv {
    param([hashtable]$Cells, [string]$Path)
    if ($Cells.Count -eq 0) {
        # Write empty header so the file's existence still signals "this sheet processed"
        Write-Utf8NoBom -Path $Path -Text "Row`r`n" -LineEnding "`r`n"
        return
    }

    # Determine extents
    $maxCol = 0; $maxRow = 0; $minRow = [int]::MaxValue
    foreach ($k in $Cells.Keys) {
        $parts = $k.Split(',')
        $c = [int]$parts[0]; $r = [int]$parts[1]
        if ($c -gt $maxCol) { $maxCol = $c }
        if ($r -gt $maxRow) { $maxRow = $r }
        if ($r -lt $minRow) { $minRow = $r }
    }

    # Group cells by row for fast lookup
    $byRow = @{}
    foreach ($k in $Cells.Keys) {
        $parts = $k.Split(',')
        $c = [int]$parts[0]; $r = [int]$parts[1]
        if (-not $byRow.ContainsKey($r)) { $byRow[$r] = @{} }
        $byRow[$r][$c] = $Cells[$k]
    }

    $sb = New-Object System.Text.StringBuilder
    # Header: Row | A | B | C | ...
    [void]$sb.Append('Row')
    for ($c = 1; $c -le $maxCol; $c++) {
        [void]$sb.Append("`t")
        [void]$sb.Append((ConvertTo-ColLetters -Col $c))
    }
    [void]$sb.Append("`r`n")

    # Body: only emit rows that have at least one populated cell
    $sortedRows = @($byRow.Keys | Sort-Object)
    foreach ($r in $sortedRows) {
        [void]$sb.Append([string]$r)
        $rowCells = $byRow[$r]
        for ($c = 1; $c -le $maxCol; $c++) {
            [void]$sb.Append("`t")
            if ($rowCells.ContainsKey($c)) {
                [void]$sb.Append((Format-TsvCell -Value $rowCells[$c]))
            }
        }
        [void]$sb.Append("`r`n")
    }

    Write-Utf8NoBom -Path $Path -Text $sb.ToString() -LineEnding "`r`n"
}

function Save-FormulasJson {
    param([hashtable]$Cells, [hashtable]$SharedMasters, [string]$Path)
    if ($Cells.Count -eq 0 -and (-not $SharedMasters -or $SharedMasters.Count -eq 0)) {
        # Don't write empty formulas.json -- absence means "no formulas"
        return
    }

    # Convert hashtable keyed by "col,row" to one keyed by A1 ref
    $byRef = @{}
    foreach ($k in $Cells.Keys) {
        $parts = $k.Split(',')
        $byRef[(ConvertTo-A1Ref -Col ([int]$parts[0]) -Row ([int]$parts[1]))] = $Cells[$k]
    }

    # Range-merge: greedy rectangles where adjacent cells share an identical formula
    $merged = Compress-CellsToRanges -Cells $Cells

    $ranges = [ordered]@{}
    $cellsOnly = [ordered]@{}
    $sortedMerged = @($merged | Sort-Object -Property @{ Expression = {
        $r = ConvertFrom-A1Range -Range ($_.Range -split ':' | Select-Object -First 1)
        if (-not $r) { return 0 }
        return $r.StartRow * 100000 + $r.StartCol
    } })
    foreach ($entry in $sortedMerged) {
        if ($entry.Range -match ':') {
            $ranges[$entry.Range] = $entry.Value
        } else {
            $cellsOnly[$entry.Range] = $entry.Value
        }
    }

    $obj = [ordered]@{}
    if ($ranges.Count -gt 0) { $obj['ranges'] = $ranges }
    if ($cellsOnly.Count -gt 0) { $obj['cells'] = $cellsOnly }
    if ($SharedMasters -and $SharedMasters.Count -gt 0) {
        $sm = [ordered]@{}
        foreach ($si in ($SharedMasters.Keys | Sort-Object)) {
            $sm["shared$si"] = [ordered]@{
                ref     = $SharedMasters[$si].Ref
                formula = "=$($SharedMasters[$si].Formula)"
            }
        }
        $obj['sharedFormulas'] = $sm
    }

    Save-Manifest -Manifest $obj -Path $Path
}

function Save-StylesJson {
    param([hashtable]$Cells, [string]$Path)
    if ($Cells.Count -eq 0) { return }   # Don't write empty styles.json

    $merged = Compress-CellsToRanges -Cells $Cells
    $sortedMerged = @($merged | Sort-Object -Property @{ Expression = {
        $r = ConvertFrom-A1Range -Range ($_.Range -split ':' | Select-Object -First 1)
        if (-not $r) { return 0 }
        return $r.StartRow * 100000 + $r.StartCol
    } })

    $obj = [ordered]@{}
    foreach ($entry in $sortedMerged) {
        $obj[$entry.Range] = $entry.Value
    }

    Save-Manifest -Manifest $obj -Path $Path
}

# ---- Phase C: per-sheet metadata files ------------------------------------

function Save-MetaJson {
    # Sheet-level metadata that's neither cell data nor formula nor style:
    # tab colour, frozen panes, column widths, hidden state, hosted-tables
    # references. Always written so a sheet folder is self-describing.
    param(
        [System.Xml.Linq.XDocument]$Doc,
        [PSCustomObject]$SheetMeta,
        [System.Collections.Generic.List[object]]$TableRanges,
        [string]$Path
    )

    $obj = [ordered]@{
        sheetId  = $SheetMeta.sheetId
        name     = $SheetMeta.name
        codeName = $SheetMeta.codeName
        state    = $SheetMeta.state
        tabColor = $SheetMeta.tabColor
    }

    # Frozen panes (from <sheetView><pane>)
    $pane = $Doc.Descendants([System.Xml.Linq.XName]::Get('pane', $NS.main)) | Select-Object -First 1
    if ($pane -and ($pane.Attribute('state') -and $pane.Attribute('state').Value -eq 'frozen')) {
        $frozen = [ordered]@{}
        if ($pane.Attribute('xSplit')) { $frozen['xSplit'] = [int]$pane.Attribute('xSplit').Value }
        if ($pane.Attribute('ySplit')) { $frozen['ySplit'] = [int]$pane.Attribute('ySplit').Value }
        $obj['frozenPanes'] = $frozen
    }

    # Show gridlines (default true; only emit when overridden)
    $sv = $Doc.Descendants([System.Xml.Linq.XName]::Get('sheetView', $NS.main)) | Select-Object -First 1
    if ($sv -and $sv.Attribute('showGridLines') -and $sv.Attribute('showGridLines').Value -eq '0') {
        $obj['showGridLines'] = $false
    }

    # Column widths (only non-default ones; default is 8.43-ish, captured in defaultColWidth)
    $cols = New-Object System.Collections.Generic.List[object]
    foreach ($colsEl in $Doc.Descendants([System.Xml.Linq.XName]::Get('cols', $NS.main))) {
        foreach ($colEl in $colsEl.Elements([System.Xml.Linq.XName]::Get('col', $NS.main))) {
            $entry = [ordered]@{}
            if ($colEl.Attribute('min')) { $entry['min'] = [int]$colEl.Attribute('min').Value }
            if ($colEl.Attribute('max')) { $entry['max'] = [int]$colEl.Attribute('max').Value }
            if ($colEl.Attribute('width')) { $entry['width'] = [double]$colEl.Attribute('width').Value }
            if ($colEl.Attribute('customWidth') -and $colEl.Attribute('customWidth').Value -eq '1') { $entry['customWidth'] = $true }
            if ($colEl.Attribute('hidden') -and $colEl.Attribute('hidden').Value -eq '1') { $entry['hidden'] = $true }
            if ($colEl.Attribute('bestFit') -and $colEl.Attribute('bestFit').Value -eq '1') { $entry['bestFit'] = $true }
            $cols.Add($entry)
        }
    }
    if ($cols.Count -gt 0) { $obj['columns'] = $cols.ToArray() }

    # Merged cells
    $mergedCells = New-Object System.Collections.Generic.List[string]
    foreach ($mc in $Doc.Descendants([System.Xml.Linq.XName]::Get('mergeCell', $NS.main))) {
        $mergedCells.Add($mc.Attribute('ref').Value)
    }
    if ($mergedCells.Count -gt 0) {
        $obj['mergedCells'] = @($mergedCells.ToArray() | Sort-Object)
    }

    # Hosted tables (cross-reference for context)
    if ($TableRanges -and $TableRanges.Count -gt 0) {
        $hostedTables = @()
        foreach ($t in $TableRanges) { $hostedTables += [ordered]@{ name = $t.Name; ref = $t.Ref } }
        $obj['tables'] = @($hostedTables | Sort-Object -Property name)
    }

    Save-Manifest -Manifest $obj -Path $Path
}

function Save-ValidationsJson {
    # Data-validation rules (dropdowns, ranges, etc) keyed by sqref range.
    # Sorted by range for stable diffs. Only emitted if any rules exist.
    param([System.Xml.Linq.XDocument]$Doc, [string]$Path)

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($dv in $Doc.Descendants([System.Xml.Linq.XName]::Get('dataValidation', $NS.main))) {
        $entry = [ordered]@{}
        $entry['sqref'] = if ($dv.Attribute('sqref')) { $dv.Attribute('sqref').Value } else { '' }
        if ($dv.Attribute('type')) { $entry['type'] = $dv.Attribute('type').Value }
        if ($dv.Attribute('operator')) { $entry['operator'] = $dv.Attribute('operator').Value }
        if ($dv.Attribute('allowBlank') -and $dv.Attribute('allowBlank').Value -eq '1') { $entry['allowBlank'] = $true }
        if ($dv.Attribute('showInputMessage') -and $dv.Attribute('showInputMessage').Value -eq '1') { $entry['showInputMessage'] = $true }
        if ($dv.Attribute('showErrorMessage') -and $dv.Attribute('showErrorMessage').Value -eq '1') { $entry['showErrorMessage'] = $true }
        if ($dv.Attribute('errorTitle')) { $entry['errorTitle'] = $dv.Attribute('errorTitle').Value }
        if ($dv.Attribute('error')) { $entry['error'] = $dv.Attribute('error').Value }
        if ($dv.Attribute('promptTitle')) { $entry['promptTitle'] = $dv.Attribute('promptTitle').Value }
        if ($dv.Attribute('prompt')) { $entry['prompt'] = $dv.Attribute('prompt').Value }
        $f1 = $dv.Element([System.Xml.Linq.XName]::Get('formula1', $NS.main))
        if ($f1) { $entry['formula1'] = $f1.Value }
        $f2 = $dv.Element([System.Xml.Linq.XName]::Get('formula2', $NS.main))
        if ($f2) { $entry['formula2'] = $f2.Value }
        $entries.Add($entry)
    }

    if ($entries.Count -eq 0) { return }
    $sorted = @($entries.ToArray() | Sort-Object -Property sqref)
    $obj = [ordered]@{ validations = $sorted }
    Save-Manifest -Manifest $obj -Path $Path
}

function Save-ConditionalFormatsJson {
    # Conditional-formatting rules keyed by sqref. Each <conditionalFormatting>
    # block has a sqref attribute and one or more <cfRule> children. Sort by
    # sqref then by priority for stable diffs.
    param([System.Xml.Linq.XDocument]$Doc, [string]$Path)

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($cf in $Doc.Descendants([System.Xml.Linq.XName]::Get('conditionalFormatting', $NS.main))) {
        $sqref = if ($cf.Attribute('sqref')) { $cf.Attribute('sqref').Value } else { '' }
        foreach ($rule in $cf.Elements([System.Xml.Linq.XName]::Get('cfRule', $NS.main))) {
            $r = [ordered]@{}
            $r['sqref'] = $sqref
            if ($rule.Attribute('type')) { $r['type'] = $rule.Attribute('type').Value }
            if ($rule.Attribute('priority')) { $r['priority'] = [int]$rule.Attribute('priority').Value }
            if ($rule.Attribute('operator')) { $r['operator'] = $rule.Attribute('operator').Value }
            if ($rule.Attribute('dxfId')) { $r['dxfId'] = [int]$rule.Attribute('dxfId').Value }
            if ($rule.Attribute('text')) { $r['text'] = $rule.Attribute('text').Value }
            if ($rule.Attribute('aboveAverage')) { $r['aboveAverage'] = ($rule.Attribute('aboveAverage').Value -ne '0') }
            $formulae = @()
            foreach ($f in $rule.Elements([System.Xml.Linq.XName]::Get('formula', $NS.main))) {
                $formulae += $f.Value
            }
            if ($formulae.Count -gt 0) { $r['formulae'] = $formulae }
            $entries.Add($r)
        }
    }

    if ($entries.Count -eq 0) { return }
    $sorted = @($entries.ToArray() | Sort-Object -Property `
        @{ Expression = { $_.sqref } }, `
        @{ Expression = { if ($_.PSObject.Properties['priority']) { $_.priority } else { 0 } } })
    $obj = [ordered]@{ conditionalFormats = $sorted }
    Save-Manifest -Manifest $obj -Path $Path
}

function Save-CommentsJson {
    # Cell comments: text + author + cell ref. Source is xl/commentsN.xml,
    # referenced via the worksheet's _rels file. Sorted by cell ref.
    param([string]$SourceXlPath, [PSCustomObject]$SheetMeta, [string]$Path)

    if (-not $SourceXlPath) { return }
    $rels = Read-WorksheetRels -SourceXlPath $SourceXlPath -WorksheetSourceFile $SheetMeta.sourceFile
    $commentsRel = $rels.Values | Where-Object { $_.Type -eq 'comments' } | Select-Object -First 1
    if (-not $commentsRel) { return }
    $commentsXmlPath = Join-Path $SourceXlPath ($commentsRel.Target -replace '/', '\')
    if (-not (Test-Path -LiteralPath $commentsXmlPath)) { return }

    $cDoc = Load-Xml -Path $commentsXmlPath
    Remove-VolatileNamespaces -Doc $cDoc

    # Authors: index -> name
    $authors = @()
    foreach ($a in $cDoc.Descendants([System.Xml.Linq.XName]::Get('author', $NS.main))) {
        $authors += $a.Value
    }

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($cm in $cDoc.Descendants([System.Xml.Linq.XName]::Get('comment', $NS.main))) {
        $ref = $cm.Attribute('ref').Value
        $authorIdx = if ($cm.Attribute('authorId')) { [int]$cm.Attribute('authorId').Value } else { -1 }
        $author = if ($authorIdx -ge 0 -and $authorIdx -lt $authors.Count) { $authors[$authorIdx] } else { $null }
        # Concat all <t> contents
        $tList = @($cm.Descendants([System.Xml.Linq.XName]::Get('t', $NS.main)))
        $text = ($tList | ForEach-Object { $_.Value }) -join ''
        $entry = [ordered]@{ ref = $ref; text = $text }
        if ($author) { $entry['author'] = $author }
        $entries.Add($entry)
    }

    if ($entries.Count -eq 0) { return }
    $sorted = @($entries.ToArray() | Sort-Object -Property @{ Expression = {
        $r = ConvertFrom-A1Ref -Ref $_.ref
        if ($r) { return $r.Row * 16384 + $r.Col } else { return 0 }
    } })
    $obj = [ordered]@{ comments = $sorted }
    Save-Manifest -Manifest $obj -Path $Path
}

function Process-Table {
    # Returns a hashtable with the table's manifest entry (name, ref, file).
    param([string]$SourceXmlPath, [string]$DestDir)

    $doc = Load-Xml -Path $SourceXmlPath
    Remove-VolatileNamespaces -Doc $doc
    Remove-VolatileAttributes -Doc $doc

    $tableName = $doc.Root.Attribute('name').Value
    $displayName = if ($doc.Root.Attribute('displayName')) { $doc.Root.Attribute('displayName').Value } else { $tableName }
    $ref = if ($doc.Root.Attribute('ref')) { $doc.Root.Attribute('ref').Value } else { $null }
    $safeName = Get-SafeFileName -Name $tableName
    $destPath = Join-Path $DestDir "$safeName.xml"

    Sort-AttributesAlphabetically -Element $doc.Root
    Save-PrettyXml -Doc $doc -Path $destPath

    return [ordered]@{
        name        = $tableName
        displayName = $displayName
        ref         = $ref
        file        = "tables/$safeName.xml"
    }
}

# Passthrough normalisation for the OOXML files that aren't workbook/sheet/table:
# strip volatile namespaces + attributes, sort attrs, pretty-print, write to dest.
# Used for xl/sharedStrings.xml and xl/styles.xml so worksheet <v>N</v> indices
# and dxfId references can actually be resolved.
#
# For sharedStrings.xml in particular, mark <sst> as xml:space="preserve" so the
# (potentially tens of thousands of) <si> entries don't blow up under XmlWriter
# indentation -- same trick as <sheetData>.
function Process-AuxiliaryXml {
    param(
        [string]$SourceXmlPath,
        [string]$DestPath,
        [string]$PreserveWhitespaceForRootChild = ''
    )
    $doc = Load-Xml -Path $SourceXmlPath
    Remove-VolatileNamespaces -Doc $doc
    Remove-VolatileAttributes -Doc $doc

    if ($PreserveWhitespaceForRootChild -and $doc.Root) {
        $xmlNs = [System.Xml.Linq.XNamespace]::Xml
        $doc.Root.SetAttributeValue($xmlNs + 'space', 'preserve')
    }

    Sort-AttributesAlphabetically -Element $doc.Root
    Save-PrettyXml -Doc $doc -Path $DestPath
}

# ---- main -----------------------------------------------------------------

function Invoke-Normalize {
    param([string]$Source, [string]$Destination)

    Ensure-Directory -Path $Destination
    $script:LogPath = Join-Path $Destination '.normalize.log'
    if (Test-Path -LiteralPath $script:LogPath) {
        Remove-Item -LiteralPath $script:LogPath -Force
    }

    $sourceXl = Join-Path $Source 'xl'
    if (-not (Test-Path -LiteralPath $sourceXl)) {
        throw "Source 'xl' folder not found at: $sourceXl. Was the .xlsx extracted correctly?"
    }

    Write-NormalizeLog -Level INFO -Message "Normalizing from $Source to $Destination"

    # Step 1: workbook + lambdas (manifest written at the end so tables array
    # can be populated by Step 3)
    $wbResult = $null
    try {
        $wbResult = Process-Workbook -SourceXlPath $sourceXl -DestPath $Destination
        Save-Lambdas -Lambdas $wbResult.Lambdas -LambdasDir (Join-Path $Destination 'lambdas')
    } catch {
        Write-NormalizeLog -Level ERROR -Message "Failed to process workbook.xml: $_"
        $script:FailedCount++
        return  # Without sheet metadata we can't process worksheets meaningfully
    }

    # Step 1b: load sharedStrings + styles ONCE (per-worksheet processing
    # references both)
    $sharedStrings = Load-SharedStrings -SourceXlPath $sourceXl
    $styles = Load-Styles -SourceXlPath $sourceXl
    Write-NormalizeLog -Level INFO -Message "Loaded $($sharedStrings.Count) shared strings, $(if ($styles) { $styles.CellXfs.Count } else { 0 }) cellXfs"

    # Step 2: worksheets (per-sheet folder with split files). Tables are
    # processed inside Process-Worksheet (nested under their host sheet folder),
    # not as a separate step. Returned table-manifest entries get aggregated
    # into MANIFEST.tables.
    $allTableEntries = @()
    $worksheetsDir = Join-Path $Destination 'worksheets'
    Ensure-Directory -Path $worksheetsDir
    foreach ($sheet in $wbResult.Sheets) {
        $sourcePath = Join-Path $sourceXl $sheet.sourceFile
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            Write-NormalizeLog -Level WARN -Message "Sheet '$($sheet.name)' (rId=$($sheet.rId)) source missing: $sourcePath"
            $script:FailedCount++
            continue
        }
        try {
            $tableEntries = Process-Worksheet -SourceXmlPath $sourcePath -SheetMeta $sheet -DestDir $worksheetsDir -SourceXlPath $sourceXl -SharedStrings $sharedStrings -Styles $styles
            if ($tableEntries) { $allTableEntries += $tableEntries }
        } catch {
            Write-NormalizeLog -Level ERROR -Message "Failed to process sheet '$($sheet.name)' from $($sheet.sourceFile): $_"
            $script:FailedCount++
        }
    }

    # Step 3: tables are now processed per-worksheet (Phase B). The old
    # Excel/tables/ folder no longer exists. Aggregate entries into manifest.
    $sortedTables = @($allTableEntries | Sort-Object -Property @{Expression={$_.sheet}}, @{Expression={$_.name}})
    $wbResult.Manifest['tables'] = $sortedTables

    # Step 4: auxiliary files needed to interpret worksheet content.
    # styles.xml is pretty-printed (small, useful to read for dxfId resolution).
    # sharedStrings.xml is byte-copied: it's an opaque Excel-managed string table,
    # the diff signal IS "string N changed" regardless of formatting, and
    # XmlWriter doesn't honor xml:space="preserve" so the splice trick used for
    # sheetData would be needed -- not worth the complexity for an opaque file.
    $stylesSrc = Join-Path $sourceXl 'styles.xml'
    if (Test-Path -LiteralPath $stylesSrc) {
        try {
            Process-AuxiliaryXml -SourceXmlPath $stylesSrc -DestPath (Join-Path $Destination 'styles.xml') -PreserveWhitespaceForRootChild ''
        } catch {
            Write-NormalizeLog -Level ERROR -Message "Failed to process styles.xml: $_"
            $script:FailedCount++
        }
    }
    $sstSrc = Join-Path $sourceXl 'sharedStrings.xml'
    if (Test-Path -LiteralPath $sstSrc) {
        try {
            # Byte-copy with one cosmetic normalisation: lowercase the encoding
            # name in the XML declaration so it matches the rest of our output
            # (XmlWriter writes encoding="utf-8"; Excel writes encoding="UTF-8").
            $sstDest = Join-Path $Destination 'sharedStrings.xml'
            $bytes = [System.IO.File]::ReadAllBytes($sstSrc)
            $text = [System.Text.UTF8Encoding]::new($false).GetString($bytes)
            $text = [regex]::Replace($text, '^(<\?xml[^?]*?encoding=")UTF-8(")', '${1}utf-8${2}', 'IgnoreCase')
            [System.IO.File]::WriteAllText($sstDest, $text, [System.Text.UTF8Encoding]::new($false))
        } catch {
            Write-NormalizeLog -Level ERROR -Message "Failed to copy sharedStrings.xml: $_"
            $script:FailedCount++
        }
    }

    # Step 5: write the manifest (now that tables array is populated)
    try {
        Save-Manifest -Manifest $wbResult.Manifest -Path (Join-Path $Destination 'MANIFEST.json')
        Write-NormalizeLog -Level INFO -Message "Wrote MANIFEST.json with $($wbResult.Sheets.Count) sheets, $($sortedTables.Count) tables, $($wbResult.Lambdas.Count) lambdas"
    } catch {
        Write-NormalizeLog -Level ERROR -Message "Failed to write MANIFEST.json: $_"
        $script:FailedCount++
    }

    # Final log line summarising redactions for VBA to surface
    if ($script:RedactedFileSharing) {
        Write-NormalizeLog -Level INFO -Message "REDACTED: workbook open password (fileSharing)"
    }
    if ($script:RedactedWorkbookProtection) {
        Write-NormalizeLog -Level INFO -Message "REDACTED: workbook structure protection password"
    }
    if ($script:RedactedSheetProtection.Count -gt 0) {
        Write-NormalizeLog -Level INFO -Message "REDACTED: sheet protection passwords for: $($script:RedactedSheetProtection -join ', ')"
    }

    Write-NormalizeLog -Level INFO -Message "Done. Failed files: $($script:FailedCount)"
}

# ---- entry point ----------------------------------------------------------

try {
    Invoke-Normalize -Source $Source -Destination $Destination
    if ($script:FailedCount -gt 0) { exit 2 } else { exit 0 }
} catch {
    if ($script:LogPath) {
        Write-NormalizeLog -Level ERROR -Message "Unrecoverable: $_"
    } else {
        [Console]::Error.WriteLine("ERROR: $_")
    }
    exit 1
}
