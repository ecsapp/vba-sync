# MS-OFORMS stream walker. Port of oletools/oleform.py.
#
# Goal: enumerate every byte inside the `f` and `o` streams that is either
#   (a) alignment padding consumed by ExtendedStream._pad(), or
#   (b) tail-skip bytes inside a will_jump_to() block (the parser declares
#       cbSize and reads `size - consummed` to fast-forward to the next record).
# Both classes are uninitialised by Excel's writer and are the source of
# .frx churn. Zeroing them is safe: Excel's importer never inspects them.
#
# The walker returns a list of (offset, length) ranges in stream-relative
# coordinates. The caller maps each range to file bytes via the CFBF
# sector tables and zeroes them in place.
#
# This file is a faithful structural port -- the ExtendedStream class
# (PadStream below) records padding ranges instead of just skipping them.
# Property mask layouts and consume_* control dispatch mirror oleform.py
# line-for-line where possible.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---- Property masks: list of bit-name + per-bit byte size when the bit is set.
# Each table mirrors the corresponding _names list in oleform.py. The byte
# size column comes from each control's consume_* function (which calls
# propmask.consume(stream, [(name, size), ...])). Bits that the parser
# never reads via .consume (e.g. fName/fTag/fSize/fValue/fCaption/fPicture
# in OleSite/MorphData -- handled by hand) carry size=0; we handle them
# in the per-control switch.

$Script:FormPropMaskNames = @(
    'Unused1', 'fBackColor', 'fForeColor', 'fNextAvailableID', 'Unused2_0', 'Unused2_1',
    'fBooleanProperties', 'fBooleanProperties2', 'fMousePointer', 'fScrollBars',
    'fDisplayedSize', 'fLogicalSize', 'fScrollPosition', 'fGroupCnt', 'Reserved',
    'fMouseIcon', 'fCycle', 'fSpecialEffect', 'fBorderColor', 'fCaption', 'fFont',
    'fPicture', 'fZoom', 'fPictureAlignment', 'fPictureTiling', 'fPictureSizeMode',
    'fShapeCookie', 'fDrawBuffer'
)

$Script:SitePropMaskNames = @(
    'fName', 'fTag', 'fID', 'fHelpContextID', 'fBitFlags', 'fObjectStreamSize',
    'fTabIndex', 'fClsidCacheIndex', 'fPosition', 'fGroupID', 'Unused1',
    'fControlTipText', 'fRuntimeLicKey', 'fControlSource', 'fRowSource'
)

$Script:MorphDataPropMaskNames = @(
    'fVariousPropertyBits', 'fBackColor', 'fForeColor', 'fMaxLength', 'fBorderStyle',
    'fScrollBars', 'fDisplayStyle', 'fMousePointer', 'fSize', 'fPasswordChar',
    'fListWidth', 'fBoundColumn', 'fTextColumn', 'fColumnCount', 'fListRows',
    'fcColumnInfo', 'fMatchEntry', 'fListStyle', 'fShowDropButtonWhen', 'UnusedBits1',
    'fDropButtonStyle', 'fMultiSelect', 'fValue', 'fCaption', 'fPicturePosition',
    'fBorderColor', 'fSpecialEffect', 'fMouseIcon', 'fPicture', 'fAccelerator',
    'UnusedBits2', 'Reserved', 'fGroupName'
)

$Script:ImagePropMaskNames = @(
    'UnusedBits1_1', 'UnusedBits1_2', 'fAutoSize', 'fBorderColor', 'fBackColor',
    'fBorderStyle', 'fMousePointer', 'fPictureSizeMode', 'fSpecialEffect', 'fSize',
    'fPicture', 'fPictureAlignment', 'fPictureTiling', 'fVariousPropertyBits',
    'fMouseIcon'
)

$Script:CommandButtonPropMaskNames = @(
    'fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fCaption', 'fPicturePosition',
    'fSize', 'fMousePointer', 'fPicture', 'fAccelerator', 'fTakeFocusOnClick',
    'fMouseIcon'
)

$Script:SpinButtonPropMaskNames = @(
    'fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fSize', 'UnusedBits1',
    'fMin', 'fMax', 'fPosition', 'fPrevEnabled', 'fNextEnabled', 'fSmallChange',
    'fOrientation', 'fDelay', 'fMouseIcon', 'fMousePointer'
)

$Script:TabStripPropMaskNames = @(
    'fListIndex', 'fBackColor', 'fForeColor', 'Unused1', 'fSize', 'fItems',
    'fMousePointer', 'Unused2', 'fTabOrientation', 'fTabStyle', 'fMultiRow',
    'fTabFixedWidth', 'fTabFixedHeight', 'fTooltips', 'Unused3', 'fTipStrings',
    'Unused4', 'fNames', 'fVariousPropertyBits', 'fNewVersion', 'fTabsAllocated',
    'fTags', 'fTabData', 'fAccelerator', 'fMouseIcon'
)

$Script:LabelPropMaskNames = @(
    'fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fCaption',
    'fPicturePosition', 'fSize', 'fMousePointer', 'fBorderColor', 'fBorderStyle',
    'fSpecialEffect', 'fPicture', 'fAccelerator', 'fMouseIcon'
)

$Script:ScrollBarPropMaskNames = @(
    'fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fSize', 'fMousePointer',
    'fMin', 'fMax', 'fPosition', 'UnusedBits1', 'fPrevEnabled', 'fNextEnabled',
    'fSmallChange', 'fLargeChange', 'fOrientation', 'fProportionalThumb',
    'fDelay', 'fMouseIcon'
)

function New-Mask([uint64]$val, [string[]]$names) {
    $bits = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    for ($i = 0; $i -lt $names.Length; $i++) {
        $bit = [int](($val -shr $i) -band 1UL)
        # Multiple names with same key (e.g. fBooleanProperties + fBooleanProperties2)
        # behave like Python's "last wins" for __getattr__ since _names.index returns first;
        # but propmask.consume() only checks each name independently. Use dictionary
        # keyed by name; OR results when duplicate name.
        if ($bits.ContainsKey($names[$i])) {
            $bits[$names[$i]] = $bits[$names[$i]] -bor $bit
        } else {
            $bits[$names[$i]] = $bit
        }
    }
    return $bits
}

# PadStream: wraps a byte[] (the concatenated stream content) and tracks
# position + padding context. read() returns bytes but the only thing we
# care about is the set of padding/skip ranges recorded for the caller.
class PadStream {
    [byte[]]$Bytes
    [int]$Pos
    [bool]$Padding
    [int]$PadStart
    # Stack of context frames: each frame is a hashtable
    #   { Kind = 'jump'|'pad'|'padded'; Data = <type-specific> }
    [System.Collections.Stack]$Frames
    # Accumulated padding ranges (stream-relative offsets) as
    # objects: { Offset = ..., Length = ... }
    [System.Collections.ArrayList]$PadRanges
    [string]$Path     # diagnostic label only

    PadStream([byte[]]$bytes, [string]$path) {
        $this.Bytes = $bytes
        $this.Pos = 0
        $this.Padding = $false
        $this.PadStart = 0
        $this.Frames = [System.Collections.Stack]::new()
        $this.PadRanges = [System.Collections.ArrayList]::new()
        $this.Path = $path
    }

    [void] RecordPad([int]$offset, [int]$length) {
        if ($length -le 0) { return }
        # Guard: don't record ranges past the end of the stream.
        if ($offset -ge $this.Bytes.Length) { return }
        $end = [Math]::Min($offset + $length, $this.Bytes.Length)
        $effective = $end - $offset
        if ($effective -le 0) { return }
        [void]$this.PadRanges.Add([pscustomobject]@{ Offset = $offset; Length = $effective })
    }

    [byte[]] ReadRaw([int]$size) {
        if ($size -lt 0) { throw "Negative read size at $($this.Pos): $size" }
        if ($size -eq 0) { return [byte[]]::new(0) }
        $end = $this.Pos + $size
        if ($end -gt $this.Bytes.Length) {
            # Truncate gracefully; do not throw -- some streams have trailing
            # data we can't predict (e.g. StdPicture preamble inside CompObj
            # is not OFORMS; we only call into f/o). Throw to surface bugs.
            throw "Read past end of stream '$($this.Path)': pos=$($this.Pos) size=$size streamLen=$($this.Bytes.Length)"
        }
        $out = New-Object byte[] $size
        [Array]::Copy($this.Bytes, $this.Pos, $out, 0, $size)
        $this.Pos = $end
        return $out
    }

    [void] PadTo([int]$start, [int]$size) {
        $offset = ($this.Pos - $start) % $size
        if ($offset -ne 0) {
            $padLen = $size - $offset
            $this.RecordPad($this.Pos, $padLen)
            $null = $this.ReadRaw($padLen)
        }
    }

    [byte[]] Read([int]$size) {
        if ($this.Padding) {
            $this.PadTo($this.PadStart, $size)
        }
        return $this.ReadRaw($size)
    }

    [void] WillJumpTo([int]$size) {
        $this.Frames.Push(@{ Kind = 'jump'; Start = $this.Pos; Size = $size })
    }

    [void] WillPad() {
        $this.Frames.Push(@{ Kind = 'pad'; Start = $this.Pos })
    }

    [void] PaddedStruct() {
        $this.Frames.Push(@{ Kind = 'padded'; PrevPadding = $this.Padding; PrevPadStart = $this.PadStart })
        $this.Padding = $true
        $this.PadStart = $this.Pos
    }

    [void] End() {
        $frame = $this.Frames.Pop()
        switch ($frame.Kind) {
            'jump' {
                $consumed = $this.Pos - $frame.Start
                if ($consumed -gt $frame.Size) {
                    throw "Bad jump in '$($this.Path)' at $($this.Pos): consumed=$consumed > declared=$($frame.Size)"
                }
                $tail = $frame.Size - $consumed
                if ($tail -gt 0) {
                    # Tail bytes are skipped by the parser, BUT they contain real
                    # deterministic content (rest of DataBlock / ExtraDataBlock that
                    # oleform doesn't bother to parse). They are NOT churn. We must
                    # advance the position but not record padding.
                    $null = $this.ReadRaw($tail)
                }
            }
            'pad' {
                $this.PadTo($frame.Start, 4)
            }
            'padded' {
                $this.PadTo($this.PadStart, 4)
                $this.Padding = $frame.PrevPadding
                $this.PadStart = $frame.PrevPadStart
            }
        }
    }

    [uint16] U16() {
        $b = $this.Read(2)
        return [uint16]([uint16]$b[0] -bor ([uint16]$b[1] -shl 8))
    }

    [uint32] U32() {
        $b = $this.Read(4)
        return ([uint32]$b[0]) -bor (([uint32]$b[1]) -shl 8) -bor `
               (([uint32]$b[2]) -shl 16) -bor (([uint32]$b[3]) -shl 24)
    }

    [uint64] U64() {
        $b = $this.Read(8)
        $lo = ([uint64]$b[0]) -bor (([uint64]$b[1]) -shl 8) -bor (([uint64]$b[2]) -shl 16) -bor (([uint64]$b[3]) -shl 24)
        $hi = ([uint64]$b[4]) -bor (([uint64]$b[5]) -shl 8) -bor (([uint64]$b[6]) -shl 16) -bor (([uint64]$b[7]) -shl 24)
        return $lo -bor ($hi -shl 32)
    }

    [byte] U8() {
        $b = $this.Read(1)
        return $b[0]
    }
}

function Read-Propmask-Consume([PadStream]$s, $mask, [object[]]$pairs) {
    foreach ($p in $pairs) {
        $name = $p[0]; $size = [int]$p[1]
        if ($mask[$name] -ne 0) { $null = $s.Read($size) }
    }
}

function Read-CountOfBytesWithCompressionFlag([PadStream]$s) {
    $v = $s.U32()
    return [uint32]($v -band 0x7FFFFFFFu)
}

function Read-TextProps([PadStream]$s) {
    # 2.3.1
    $null = $s.Read(2)               # versions: BB (0, 2)
    $cb = $s.U16()
    $null = $s.Read($cb)
}

function Read-GuidAndFont([PadStream]$s) {
    # 2.4.7 -- UUID 16 bytes
    $u = $s.Read(16)
    # Decide on StdFont vs TextProps via the leading DWORD.
    $uuid0 = ([uint32]$u[0]) -bor (([uint32]$u[1]) -shl 8) -bor (([uint32]$u[2]) -shl 16) -bor (([uint32]$u[3]) -shl 24)
    if ($uuid0 -eq 0x0BE35203u) {
        # StdFont: version (1), then 9 bytes (sCharset, bFlags, sWeight, ulHeight), bFaceLen (1), faceName
        $null = $s.Read(1)
        $null = $s.Read(9)
        $faceLen = $s.U8()
        $null = $s.Read($faceLen)
    } elseif ($uuid0 -eq 0xAFC20920u) {
        Read-TextProps $s
    } else {
        throw "GuidAndFont: unknown UUID prefix 0x$('{0:X8}' -f $uuid0) at stream pos $($s.Pos - 16)"
    }
}

function Read-GuidAndPicture([PadStream]$s) {
    # 2.4.8 -- 16-byte UUID, then StdPicture
    $null = $s.Read(16)
    $null = $s.U32()                 # preamble 0x0000746C
    $sz = $s.U32()
    $null = $s.Read($sz)
}

function Read-SiteClassInfo([PadStream]$s) {
    # 2.2.10.10.1
    $null = $s.U16()
    $cb = $s.U16()
    $null = $s.Read($cb)
}

function Read-FormObjectDepthTypeCount([PadStream]$s) {
    # 2.2.10.7
    $b = $s.Read(2)
    $depth = $b[0]
    $mixed = $b[1]
    if ($mixed -band 0x80) {
        $null = $s.U8()              # SITE_TYPE byte = 1
        return ($mixed -bxor 0x80)
    }
    if ($mixed -ne 1) {
        throw "FormObjectDepthTypeCount: expected 1 got $mixed at pos $($s.Pos - 1)"
    }
    return 1
}

function Read-OleSiteConcreteControl([PadStream]$s) {
    # 2.2.10.12.1
    $null = $s.U16()                                    # version 0
    $cb = $s.U16()
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:SitePropMaskNames
        $s.PaddedStruct()
        try {
            $nameLen = 0; $tagLen = 0
            if ($mask.fName) { $nameLen = Read-CountOfBytesWithCompressionFlag $s }
            if ($mask.fTag)  { $tagLen  = Read-CountOfBytesWithCompressionFlag $s }
            if ($mask.fID)   { $null = $s.U32() }
            Read-Propmask-Consume $s $mask @(@('fHelpContextID', 4), @('fBitFlags', 4), @('fObjectStreamSize', 4))
            $tabindex = 0; $clsidCacheIndex = 0
            if ($mask.fTabIndex)        { $tabindex = $s.U16() }
            if ($mask.fClsidCacheIndex) { $clsidCacheIndex = $s.U16() }
            if ($mask.fGroupID)         { $null = $s.Read(2) }
            $tipLen = 0
            if ($mask.fControlTipText)  { $tipLen = Read-CountOfBytesWithCompressionFlag $s }
            Read-Propmask-Consume $s $mask @(@('fRuntimeLicKey', 4), @('fControlSource', 4), @('fRowSource', 4))
        } finally { $s.End() }
        # SiteExtraDataBlock 2.2.10.12.4
        # Each variable-length string (Name / Tag / ControlTipText) is followed by
        # 0..3 bytes of alignment padding so the next field starts on a 4-byte
        # boundary relative to the SiteExtraDataBlock start. oleform.py omits this
        # (its inline comment notes "Sometimes it looks like 2 extra null bytes"),
        # but Excel always pads -- and those pad bytes are uninitialised, which is
        # the second-biggest source of .frx churn after the cbForm tail-skip.
        $extraStart = $s.Pos
        if ($nameLen -gt 0) {
            $null = $s.Read($nameLen)
            $s.PadTo($extraStart, 4)
        }
        if ($tagLen  -gt 0) {
            $null = $s.Read($tagLen)
            $s.PadTo($extraStart, 4)
        }
        if ($mask.fPosition) { $null = $s.Read(8) }
        if ($tipLen -gt 0) {
            $null = $s.Read($tipLen)
            $s.PadTo($extraStart, 4)
        }
        return [pscustomobject]@{ ClsidCacheIndex = $clsidCacheIndex; TabIndex = $tabindex }
    } finally { $s.End() }
}

function Read-FormControl([PadStream]$s) {
    # 2.2.10.1 -- returns the list of children (OleSites).
    $null = $s.Read(2)                                  # versions BB (0, 4)
    $cb = $s.U16()
    $dontSaveClassTable = 0
    $maskFMouseIcon = 0; $maskFFont = 0; $maskFPicture = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:FormPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        $maskFFont      = $mask.fFont
        $maskFPicture   = $mask.fPicture
        Read-Propmask-Consume $s $mask @(@('fBackColor', 4), @('fForeColor', 4), @('fNextAvailableID', 4))
        if ($mask.fBooleanProperties) {
            $bp = $s.U32()
            $dontSaveClassTable = ([int](($bp -shr 15) -band 1u))
        }
        # The rest of FormDataBlock and FormExtraDataBlock is jumped over by
        # will_jump_to. We leave it to the End() padding-record path.
    } finally { $s.End() }
    # FormStreamData 2.2.10.5
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
    if ($maskFFont)      { Read-GuidAndFont    $s }
    if ($maskFPicture)   { Read-GuidAndPicture $s }
    # FormSiteData 2.2.10.6
    if (-not $dontSaveClassTable) {
        $countSiteClassInfo = $s.U16()
        for ($i = 0; $i -lt $countSiteClassInfo; $i++) { Read-SiteClassInfo $s }
    }
    $countOfSites = $s.U32()
    $countOfBytes = $s.U32()
    $remaining = $countOfSites
    $sites = New-Object System.Collections.ArrayList
    $s.WillJumpTo($countOfBytes)
    try {
        $s.WillPad()
        try {
            while ($remaining -gt 0) {
                $consumed = Read-FormObjectDepthTypeCount $s
                $remaining -= $consumed
            }
        } finally { $s.End() }
        for ($i = 0; $i -lt $countOfSites; $i++) {
            [void]$sites.Add((Read-OleSiteConcreteControl $s))
        }
    } finally { $s.End() }
    return ,$sites.ToArray()
}

function Read-MorphDataControl([PadStream]$s) {
    # 2.2.5.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0; $maskFPicture = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U64()
        $mask = New-Mask $maskVal $Script:MorphDataPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        $maskFPicture   = $mask.fPicture
        $s.PaddedStruct()
        try {
            Read-Propmask-Consume $s $mask @(
                @('fVariousPropertyBits', 4), @('fBackColor', 4),
                @('fForeColor', 4),           @('fMaxLength', 4),
                @('fBorderStyle', 1),         @('fScrollBars', 1),
                @('fDisplayStyle', 1),        @('fMousePointer', 1),
                @('fPasswordChar', 2),        @('fListWidth', 4),
                @('fBoundColumn', 2),         @('fTextColumn', 2),
                @('fColumnCount', 2),         @('fListRows', 2),
                @('fcColumnInfo', 2),         @('fMatchEntry', 1),
                @('fListStyle', 1),           @('fShowDropButtonWhen', 1),
                @('fDropButtonStyle', 1),     @('fMultiSelect', 1)
            )
            $valueSize = 0
            if ($mask.fValue)   { $valueSize = Read-CountOfBytesWithCompressionFlag $s }
            $captionSize = 0
            if ($mask.fCaption) { $captionSize = Read-CountOfBytesWithCompressionFlag $s }
            Read-Propmask-Consume $s $mask @(
                @('fPicturePosition', 4), @('fBorderColor', 4),
                @('fSpecialEffect', 4),   @('fMouseIcon', 2),
                @('fPicture', 2),         @('fAccelerator', 2)
            )
            $groupNameSize = 0
            if ($mask.fGroupName) { $groupNameSize = Read-CountOfBytesWithCompressionFlag $s }
        } finally { $s.End() }
        # MorphDataExtraDataBlock 2.2.5.4
        $null = $s.Read(8)
        if ($valueSize     -gt 0) { $null = $s.Read($valueSize) }
        if ($captionSize   -gt 0) { $null = $s.Read($captionSize) }
        if ($groupNameSize -gt 0) { $null = $s.Read($groupNameSize) }
    } finally { $s.End() }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
    if ($maskFPicture)   { Read-GuidAndPicture $s }
    Read-TextProps $s
}

function Read-ImageControl([PadStream]$s) {
    # 2.2.3.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0; $maskFPicture = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:ImagePropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        $maskFPicture   = $mask.fPicture
    } finally { $s.End() }
    if ($maskFPicture)   { Read-GuidAndPicture $s }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
}

function Read-CommandButtonControl([PadStream]$s) {
    # 2.2.1.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0; $maskFPicture = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:CommandButtonPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        $maskFPicture   = $mask.fPicture
    } finally { $s.End() }
    if ($maskFPicture)   { Read-GuidAndPicture $s }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
    Read-TextProps $s
}

function Read-SpinButtonControl([PadStream]$s) {
    # 2.2.8.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:SpinButtonPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
    } finally { $s.End() }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
}

function Read-TabStripControl([PadStream]$s) {
    # 2.2.9.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0
    $tabData = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:TabStripPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        Read-Propmask-Consume $s $mask @(
            @('fListIndex', 4), @('fBackColor', 4),
            @('fForeColor', 4), @('fSize', 4),
            @('fMousePointer', 1), @('fTabOrientation', 4),
            @('fTabStyle', 4), @('fTabFixedWidth', 4),
            @('fTabFixedHeight', 4), @('fTipStrings', 4),
            @('fNames', 4), @('fVariousPropertyBits', 4),
            @('fTabsAllocated', 4), @('fTags', 4)
        )
        if ($mask.fTabData) { $tabData = $s.U32() }
    } finally { $s.End() }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
    Read-TextProps $s
    for ($i = 0; $i -lt $tabData; $i++) { $null = $s.Read(4) }
}

function Read-LabelControl([PadStream]$s) {
    # 2.2.4.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0; $maskFPicture = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:LabelPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
        $maskFPicture   = $mask.fPicture
        $s.PaddedStruct()
        try {
            Read-Propmask-Consume $s $mask @(@('fForeColor', 4), @('fBackColor', 4), @('fVariousPropertyBits', 4))
            $captionSize = 0
            if ($mask.fCaption) { $captionSize = Read-CountOfBytesWithCompressionFlag $s }
            Read-Propmask-Consume $s $mask @(
                @('fPicturePosition', 4), @('fMousePointer', 1),
                @('fBorderColor', 4),     @('fBorderStyle', 2),
                @('fSpecialEffect', 2),   @('fPicture', 2),
                @('fAccelerator', 2),     @('fMouseIcon', 2)
            )
        } finally { $s.End() }
        if ($captionSize -gt 0) { $null = $s.Read($captionSize) }
        $null = $s.Read(8)
    } finally { $s.End() }
    if ($maskFPicture)   { Read-GuidAndPicture $s }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
    Read-TextProps $s
}

function Read-ScrollBarControl([PadStream]$s) {
    # 2.2.7.1
    $null = $s.Read(2)
    $cb = $s.U16()
    $maskFMouseIcon = 0
    $s.WillJumpTo($cb)
    try {
        $maskVal = $s.U32()
        $mask = New-Mask $maskVal $Script:ScrollBarPropMaskNames
        $maskFMouseIcon = $mask.fMouseIcon
    } finally { $s.End() }
    if ($maskFMouseIcon) { Read-GuidAndPicture $s }
}

# Top-level walker. Takes:
#   $fStreamBytes : byte[] of the 'f' stream content
#   $oStreamBytes : byte[] of the 'o' stream content
# Returns hashtable @{ FPad = [pscustomobject[]]; OPad = [pscustomobject[]] }
# where each entry is { Offset, Length } in stream-relative coordinates.
function Get-OformsPaddingRanges {
    param(
        [Parameter(Mandatory)][byte[]]$FStreamBytes,
        [Parameter(Mandatory)][byte[]]$OStreamBytes
    )

    $fStream = [PadStream]::new($FStreamBytes, 'f')
    $oStream = [PadStream]::new($OStreamBytes, 'o')

    $sites = Read-FormControl $fStream

    foreach ($var in $sites) {
        $idx = $var.ClsidCacheIndex
        switch ($idx) {
            7  { $null = Read-FormControl        $oStream }
            12 {        Read-ImageControl        $oStream }
            14 { $null = Read-FormControl        $oStream }
            15 {        Read-MorphDataControl    $oStream }
            16 {        Read-SpinButtonControl   $oStream }
            17 {        Read-CommandButtonControl $oStream }
            18 {        Read-TabStripControl     $oStream }
            21 {        Read-LabelControl        $oStream }
            23 {        Read-MorphDataControl    $oStream }
            24 {        Read-MorphDataControl    $oStream }
            25 {        Read-MorphDataControl    $oStream }
            26 {        Read-MorphDataControl    $oStream }
            27 {        Read-MorphDataControl    $oStream }
            28 {        Read-MorphDataControl    $oStream }
            47 {        Read-ScrollBarControl    $oStream }
            57 { $null = Read-FormControl        $oStream }
            default { throw "Unsupported ClsidCacheIndex $idx in OFORMS o-stream" }
        }
    }

    return @{
        FPad = $fStream.PadRanges.ToArray()
        OPad = $oStream.PadRanges.ToArray()
    }
}
