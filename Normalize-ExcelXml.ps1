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
    # selection / scroll position
    'activeCell', 'sqref', 'topLeftCell', 'pane', 'activePane'
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
    $ignorableUris = Get-IgnorableNamespaces -Doc $Doc
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
                Write-NormalizeLog -Level WARN -Message "${FileLabel}: redacted <fileSharing> (workbook-open password hash)"
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
                Write-NormalizeLog -Level WARN -Message "${FileLabel}: redacted <workbookProtection> hash/salt"
            } else {
                $script:RedactedSheetProtection.Add($FileLabel)
                Write-NormalizeLog -Level WARN -Message "${FileLabel}: redacted <sheetProtection> hash/salt"
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
    $bodyToName = @{}            # bodyHash -> name (for cross-name dedup, optional)

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
    param([System.Xml.Linq.XDocument]$Doc)

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
        $key = if ($null -eq $e.localSheetId) { $e.name } else { "$($e.name)@sheet$($e.localSheetId)" }
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
    $definedNames = Build-DefinedNames -Doc $doc

    # Workbook properties
    $wbProps = Get-WorkbookProperties -Doc $doc

    # Build the manifest
    $manifest = [ordered]@{
        schemaVersion = 1
        workbook      = $wbProps
        sheets        = @()
        definedNames  = $definedNames
        lambdas       = @($lambdas.Keys | Sort-Object)
    }

    foreach ($s in $sheets) {
        $safeName = Get-SafeFileName -Name $s.name
        $padded = ([string]$s.sheetId).PadLeft(2, '0')
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
    param([string]$SourceXmlPath, [PSCustomObject]$SheetMeta, [string]$DestDir)

    $doc = Load-Xml -Path $SourceXmlPath
    Remove-VolatileNamespaces -Doc $doc
    $fileLabel = "worksheets/$($SheetMeta.name)"
    Remove-VolatileElements -Doc $doc -FileLabel $fileLabel
    Strip-ProtectionSecrets -Doc $doc -FileLabel $fileLabel
    Remove-VolatileAttributes -Doc $doc

    # We do NOT recurse into <sheetData> for attribute reordering -- it can have
    # tens of thousands of <c>/<v>/<f> elements with semantic content. Sort-
    # AttributesAlphabetically is recursive though, so temporarily detach
    # sheetData, sort the rest, then re-attach.
    $sheetData = $doc.Descendants([System.Xml.Linq.XName]::Get('sheetData', $NS.main)) | Select-Object -First 1
    $sheetDataParent = $null
    $sheetDataIndex = -1
    if ($sheetData) {
        $sheetDataParent = $sheetData.Parent
        $siblings = @($sheetDataParent.Elements())
        for ($i = 0; $i -lt $siblings.Count; $i++) {
            if ([object]::ReferenceEquals($siblings[$i], $sheetData)) { $sheetDataIndex = $i; break }
        }
        $sheetData.Remove()
    }

    Sort-AttributesAlphabetically -Element $doc.Root

    if ($sheetData -and $sheetDataParent) {
        # Re-insert at the same index. XElement doesn't have InsertAt, but we can
        # add then move via the previous sibling.
        $allElements = @($sheetDataParent.Elements())
        if ($sheetDataIndex -ge $allElements.Count) {
            $sheetDataParent.Add($sheetData)
        } elseif ($sheetDataIndex -le 0) {
            $sheetDataParent.AddFirst($sheetData)
        } else {
            $allElements[$sheetDataIndex - 1].AddAfterSelf($sheetData)
        }
    }

    $safeName = Get-SafeFileName -Name $SheetMeta.name
    $padded = ([string]$SheetMeta.sheetId).PadLeft(2, '0')
    $destPath = Join-Path $DestDir "$padded - $safeName.xml"
    Save-PrettyXml -Doc $doc -Path $destPath
}

function Process-Table {
    param([string]$SourceXmlPath, [string]$DestDir)

    $doc = Load-Xml -Path $SourceXmlPath
    Remove-VolatileNamespaces -Doc $doc
    Remove-VolatileAttributes -Doc $doc

    $tableName = $doc.Root.Attribute('name').Value
    $safeName = Get-SafeFileName -Name $tableName
    $destPath = Join-Path $DestDir "$safeName.xml"

    Sort-AttributesAlphabetically -Element $doc.Root
    Save-PrettyXml -Doc $doc -Path $destPath
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

    # Step 1: workbook + manifest + lambdas
    $wbResult = $null
    try {
        $wbResult = Process-Workbook -SourceXlPath $sourceXl -DestPath $Destination
        Save-Manifest -Manifest $wbResult.Manifest -Path (Join-Path $Destination 'MANIFEST.json')
        Save-Lambdas -Lambdas $wbResult.Lambdas -LambdasDir (Join-Path $Destination 'lambdas')
        Write-NormalizeLog -Level INFO -Message "Wrote MANIFEST.json with $($wbResult.Sheets.Count) sheets, $($wbResult.Lambdas.Count) lambdas"
    } catch {
        Write-NormalizeLog -Level ERROR -Message "Failed to process workbook.xml: $_"
        $script:FailedCount++
        return  # Without sheet metadata we can't process worksheets meaningfully
    }

    # Step 2: worksheets (per-sheet, with name lookup)
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
            Process-Worksheet -SourceXmlPath $sourcePath -SheetMeta $sheet -DestDir $worksheetsDir
        } catch {
            Write-NormalizeLog -Level ERROR -Message "Failed to process sheet '$($sheet.name)' from $($sheet.sourceFile): $_"
            $script:FailedCount++
        }
    }

    # Step 3: tables
    $sourceTablesDir = Join-Path $sourceXl 'tables'
    if (Test-Path -LiteralPath $sourceTablesDir) {
        $destTablesDir = Join-Path $Destination 'tables'
        Ensure-Directory -Path $destTablesDir
        $tableFiles = @(Get-ChildItem -LiteralPath $sourceTablesDir -Filter '*.xml' -File)
        foreach ($tf in $tableFiles) {
            try {
                Process-Table -SourceXmlPath $tf.FullName -DestDir $destTablesDir
            } catch {
                Write-NormalizeLog -Level ERROR -Message "Failed to process table $($tf.Name): $_"
                $script:FailedCount++
            }
        }
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
