# Verbose walk: print each significant step + position in the f stream
# so we can map each f byte to the OFORMS field it belongs to.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Path)
. (Join-Path $PSScriptRoot '..\Cfbf.ps1')
. (Join-Path $PSScriptRoot '..\Oforms.ps1')

$ole = Read-CfbfStreams $Path
$f = $ole.Streams | Where-Object { $_.Name.Trim() -eq 'f' } | Select-Object -First 1
$fBytes = [byte[]](ConvertTo-StreamBytes $ole.Bytes $f)
$s = [PadStream]::new($fBytes, 'f')

function P($label) { '  pos=0x{0:X4}/{1}  --  {2}' -f $s.Pos, $fBytes.Length, $label | Write-Host }

P 'before FormControl version'
$null = $s.Read(2)                                  # versions
$cb = $s.U16()
P "cbForm = $cb"
$s.WillJumpTo($cb)
$maskVal = $s.U32()
$mask = New-Mask $maskVal $Script:FormPropMaskNames
P "FormPropMask = 0x{0:X8}" -f $maskVal
P "  fBooleanProperties = $($mask.fBooleanProperties)"
P "  fBooleanProperties2 = $($mask.fBooleanProperties2)"
P "  fMouseIcon=$($mask.fMouseIcon)  fFont=$($mask.fFont)  fPicture=$($mask.fPicture)"
Read-Propmask-Consume $s $mask @(@('fBackColor', 4), @('fForeColor', 4), @('fNextAvailableID', 4))
P 'after fBackColor/fForeColor/fNextAvailableID consume'
if ($mask.fBooleanProperties) {
    $bp = $s.U32()
    P "  BooleanProperties = 0x{0:X8}" -f $bp
}
$s.End()                                            # finish FormDataBlock jump
P 'after FormDataBlock jump'
if ($mask.fMouseIcon) { Read-GuidAndPicture $s; P 'after MouseIcon' }
if ($mask.fFont)      { Read-GuidAndFont    $s; P 'after Font' }
if ($mask.fPicture)   { Read-GuidAndPicture $s; P 'after Picture' }

$countSiteClassInfo = $s.U16()
P "countSiteClassInfo = $countSiteClassInfo"
for ($i = 0; $i -lt $countSiteClassInfo; $i++) { Read-SiteClassInfo $s; P "  after SiteClassInfo[$i]" }
$countOfSites = $s.U32()
$countOfBytes = $s.U32()
P "countOfSites=$countOfSites  countOfBytes=$countOfBytes"
$remaining = $countOfSites
$s.WillJumpTo($countOfBytes)
$s.WillPad()
while ($remaining -gt 0) {
    $consumed = Read-FormObjectDepthTypeCount $s
    $remaining -= $consumed
}
$s.End()
P 'after FormObjectDepthTypeCount loop (padded to 4)'
for ($i = 0; $i -lt $countOfSites; $i++) {
    P "  --- site[$i] start at pos=0x$('{0:X}' -f $s.Pos) ---"
    $null = $s.Read(2)
    $cbSite = $s.U16()
    P "    cbSite = $cbSite"
    $s.WillJumpTo($cbSite)
    $sm = New-Mask ($s.U32()) $Script:SitePropMaskNames
    P "    SiteMask: fName=$($sm.fName) fTag=$($sm.fTag) fID=$($sm.fID) fHelpContextID=$($sm.fHelpContextID) fBitFlags=$($sm.fBitFlags) fObjectStreamSize=$($sm.fObjectStreamSize) fTabIndex=$($sm.fTabIndex) fClsidCacheIndex=$($sm.fClsidCacheIndex) fPosition=$($sm.fPosition) fGroupID=$($sm.fGroupID) fControlTipText=$($sm.fControlTipText) fRuntimeLicKey=$($sm.fRuntimeLicKey) fControlSource=$($sm.fControlSource) fRowSource=$($sm.fRowSource)"
    $s.PaddedStruct()
    $nameLen = 0; $tagLen = 0
    if ($sm.fName) { $nameLen = Read-CountOfBytesWithCompressionFlag $s; P "      nameLen=$nameLen (pos=0x$('{0:X}' -f $s.Pos))" }
    if ($sm.fTag)  { $tagLen  = Read-CountOfBytesWithCompressionFlag $s; P "      tagLen=$tagLen (pos=0x$('{0:X}' -f $s.Pos))" }
    if ($sm.fID)   { $null = $s.U32(); P "      fID read" }
    Read-Propmask-Consume $s $sm @(@('fHelpContextID', 4), @('fBitFlags', 4), @('fObjectStreamSize', 4))
    if ($sm.fTabIndex)        { $null = $s.U16() }
    if ($sm.fClsidCacheIndex) { $cci = $s.U16(); P "      ClsidCacheIndex=$cci" }
    if ($sm.fGroupID)         { $null = $s.Read(2) }
    $tipLen = 0
    if ($sm.fControlTipText)  { $tipLen = Read-CountOfBytesWithCompressionFlag $s; P "      tipLen=$tipLen" }
    Read-Propmask-Consume $s $sm @(@('fRuntimeLicKey', 4), @('fControlSource', 4), @('fRowSource', 4))
    $s.End()
    P "      pos after padded_struct = 0x$('{0:X}' -f $s.Pos)"
    if ($nameLen -gt 0) { $null = $s.Read($nameLen); P "      name read (now pos=0x$('{0:X}' -f $s.Pos))" }
    if ($tagLen  -gt 0) { $null = $s.Read($tagLen);  P "      tag read"  }
    if ($sm.fPosition)  { $null = $s.Read(8);        P "      position read" }
    if ($tipLen   -gt 0) { $null = $s.Read($tipLen); P "      tip read" }
    $s.End()
    P "    end site[$i] at pos=0x$('{0:X}' -f $s.Pos)"
}
$s.End()    # countOfBytes jump
P "done. final pos=0x$('{0:X}' -f $s.Pos), streamLen=$($fBytes.Length)"
"pad ranges:"
foreach ($p in $s.PadRanges) { "  off=0x{0:X} len={1}" -f $p.Offset, $p.Length | Write-Host }
