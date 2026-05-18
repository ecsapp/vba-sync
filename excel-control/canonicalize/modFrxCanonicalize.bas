Attribute VB_Name = "modFrxCanonicalize"
'  modFrxCanonicalize
'
'  Pure-VBA MS-OFORMS canonicalizer for .frx files.
'  Reads FrxPath, walks the CFBF container, finds the f/o streams,
'  walks the MS-OFORMS records inside them, and zeros every
'  alignment-padding byte (the source of Excel's non-deterministic
'  .frx writer).  Writes the result back in place.
'
'  Returns the number of bytes zeroed, or -1 on parse error.
'
'  Self-contained: no References beyond default .xlam (Scripting
'  Runtime not even required -- uses Open ... For Binary).
'
'  Port of decalage2/oletools/oleform.py + a minimal CFBF reader.

Option Explicit
Option Compare Binary

'-- CFBF constants
Private Const CFBF_HEADER_LEN As Long = 512
Private Const CFBF_DIR_ENTRY_LEN As Long = 128
Private Const CFBF_ENDOFCHAIN As Long = -2     ' 0xFFFFFFFE as signed Long
Private Const CFBF_FREESECT As Long = -1       ' 0xFFFFFFFF
Private Const CFBF_FATSECT As Long = -3        ' 0xFFFFFFFD
Private Const CFBF_DIFSECT As Long = -4        ' 0xFFFFFFFC

'-- The file-byte buffer and the padding-mask we accumulate
Private mBytes()    As Byte
Private mBytesLen   As Long
Private mPadMask()  As Byte     '  1 = zero-this-byte, 0 = keep

'-- CFBF metadata (populated by ReadCfbfHeader)
Private mCfbfStart      As Long          ' offset of CFBF magic in file (24 for .frx)
Private mSectorSize     As Long
Private mMiniSectorSize As Long
Private mMiniCutoff     As Long
Private mFirstDirSec    As Long
Private mFirstMiniFatSec As Long
Private mNumMiniFatSec  As Long
Private mFirstDifatSec  As Long
Private mNumDifatSec    As Long
Private mFat()          As Long          ' FAT entries (signed Long)
Private mFatLen         As Long
Private mMiniFat()      As Long
Private mMiniFatLen     As Long
Private mMiniStreamSec  As Long          ' first sector of the mini-stream (Root Entry's start)
Private mMiniStreamSize As Long          ' Root Entry size

'-- Per-stream buffer + file-offset mapping
'   When we read a stream, we copy bytes from the file into mStreamBuf,
'   and remember (per byte) which file offset it came from in mStreamMap.
'   Then when the OFORMS walker says "byte i is padding", we set
'   mPadMask(mStreamMap(i)) = 1.
Private mStreamBuf()    As Byte
Private mStreamMap()    As Long
Private mStreamLen      As Long
Private mStreamPos      As Long          ' read cursor

'-- Padded-struct epilog stack (mimics ExtendedStream.padded_struct)
Private mPaddingActive  As Boolean
Private mPadStart       As Long          ' stream offset where current padded run started
Private Const PAD_STACK_MAX As Long = 16
Private mPadStack(0 To PAD_STACK_MAX - 1) As Long
Private mPadStackActive(0 To PAD_STACK_MAX - 1) As Boolean
Private mPadStackTop    As Long

'-- Error flag
Private mParseError     As Boolean
Private mParseErrorMsg  As String
Private mStage          As String


'======================================================================
'  PUBLIC ENTRY POINT
'======================================================================
Public Function CanonicalizeFrxLastError() As String
    CanonicalizeFrxLastError = mParseErrorMsg
End Function

Public Function CanonicalizeFrx(ByVal FrxPath As String) As Long
    On Error GoTo Fail
    mParseErrorMsg = ""
    mParseError = False
    mPadStackTop = 0
    mPaddingActive = False
    mPadStart = 0

    mStage = "LoadFile"
    If Not LoadFile(FrxPath) Then
        CanonicalizeFrx = -1
        Exit Function
    End If

    ' Re-init pad mask to zero (no bytes flagged)
    ReDim mPadMask(0 To mBytesLen - 1)

    ' --- Locate CFBF magic (typical .frx wraps it at offset 24) ---
    mStage = "LocateCfbfStart"
    If Not LocateCfbfStart() Then
        CanonicalizeFrx = -1
        Exit Function
    End If

    ' --- Parse CFBF header + FAT + miniFAT ---
    mStage = "ReadCfbfHeader"
    If Not ReadCfbfHeader() Then GoTo Fail
    mStage = "ReadFat"
    If Not ReadFat() Then GoTo Fail
    mStage = "ReadMiniFat"
    If Not ReadMiniFat() Then GoTo Fail

    ' --- Walk directory, find Root Entry (mini-stream chain), f, o ---
    Dim fStartSec As Long, fSize As Long
    Dim oStartSec As Long, oSize As Long
    Dim foundF As Boolean, foundO As Boolean
    mStage = "WalkDirectory"
    If Not WalkDirectory(fStartSec, fSize, foundF, oStartSec, oSize, foundO) Then GoTo Fail

    If Not foundF Then
        mParseErrorMsg = "No 'f' stream"
        GoTo Fail
    End If
    If Not foundO Then
        mParseErrorMsg = "No 'o' stream"
        GoTo Fail
    End If

    ' --- Walk the 'f' stream (FormControl + OleSiteConcreteControl array) ---
    Dim siteClsids() As Long
    Dim siteCount As Long
    mStage = "LoadStream(f)"
    If Not LoadStream(fStartSec, fSize) Then GoTo Fail
    mStage = "ParseFormControlStream"
    If Not ParseFormControlStream(siteClsids, siteCount) Then GoTo Fail

    ' --- Walk the 'o' stream (ObjectStreamRecord array, one per site) ---
    mStage = "LoadStream(o)"
    If Not LoadStream(oStartSec, oSize) Then GoTo Fail
    mStage = "ParseObjectStream"
    If Not ParseObjectStream(siteClsids, siteCount) Then GoTo Fail

    ' --- Zero out flagged bytes and write file back ---
    Dim zeroed As Long
    zeroed = ApplyPadMaskAndCount()
    If Not SaveFile(FrxPath) Then GoTo Fail

    CanonicalizeFrx = zeroed
    Exit Function

Fail:
    If Err.Number <> 0 Then
        mParseErrorMsg = "VBA error #" & Err.Number & ": " & Err.Description & " (stage=" & mStage & ", streamPos=" & mStreamPos & "/" & mStreamLen & ")"
    ElseIf Len(mParseErrorMsg) = 0 Then
        mParseErrorMsg = "Unknown error (stage=" & mStage & ", streamPos=" & mStreamPos & "/" & mStreamLen & ")"
    Else
        mParseErrorMsg = mParseErrorMsg & " (stage=" & mStage & ", streamPos=" & mStreamPos & "/" & mStreamLen & ")"
    End If
    Debug.Print "CanonicalizeFrx error: " & mParseErrorMsg
    CanonicalizeFrx = -1
End Function


'======================================================================
'  FILE I/O
'======================================================================
Private Function LoadFile(ByVal Path As String) As Boolean
    On Error GoTo Bail
    Dim fnum As Integer
    fnum = FreeFile
    Open Path For Binary Access Read As #fnum
    mBytesLen = LOF(fnum)
    If mBytesLen <= 0 Then
        Close #fnum
        mParseErrorMsg = "Empty file"
        Exit Function
    End If
    ReDim mBytes(0 To mBytesLen - 1)
    Get #fnum, 1, mBytes
    Close #fnum
    LoadFile = True
    Exit Function
Bail:
    mParseErrorMsg = "LoadFile: " & Err.Description
End Function

Private Function SaveFile(ByVal Path As String) As Boolean
    On Error GoTo Bail
    Dim fnum As Integer
    fnum = FreeFile
    Open Path For Binary Access Write As #fnum
    Put #fnum, 1, mBytes
    Close #fnum
    SaveFile = True
    Exit Function
Bail:
    mParseErrorMsg = "SaveFile: " & Err.Description
End Function

Private Function ApplyPadMaskAndCount() As Long
    Dim i As Long, n As Long
    n = 0
    For i = 0 To mBytesLen - 1
        If mPadMask(i) <> 0 Then
            If mBytes(i) <> 0 Then
                mBytes(i) = 0
                n = n + 1
            End If
        End If
    Next i
    ApplyPadMaskAndCount = n
End Function


'======================================================================
'  CFBF: header, FAT, miniFAT, directory, stream materialization
'======================================================================

' Scan the first 64 bytes for CFBF magic D0 CF 11 E0 A1 B1 1A E1.
' In a typical .frx that's at offset 24.
Private Function LocateCfbfStart() As Boolean
    Dim i As Long
    Dim probe As Long
    probe = mBytesLen - 8
    If probe > 64 Then probe = 64
    For i = 0 To probe
        If mBytes(i) = &HD0 And mBytes(i + 1) = &HCF And _
           mBytes(i + 2) = &H11 And mBytes(i + 3) = &HE0 And _
           mBytes(i + 4) = &HA1 And mBytes(i + 5) = &HB1 And _
           mBytes(i + 6) = &H1A And mBytes(i + 7) = &HE1 Then
            mCfbfStart = i
            LocateCfbfStart = True
            Exit Function
        End If
    Next i
    mParseErrorMsg = "CFBF magic not found in first 64 bytes"
End Function

Private Function ReadCfbfHeader() As Boolean
    Dim h As Long
    h = mCfbfStart
    If h + CFBF_HEADER_LEN > mBytesLen Then
        mParseErrorMsg = "Truncated CFBF header"
        Exit Function
    End If
    ' Sector shift at +30 (WORD LE) -> 2^x = sector size
    Dim sectorShift As Long
    sectorShift = ReadWordLE(h + 30)
    mSectorSize = 2 ^ sectorShift
    ' Mini sector shift at +32
    Dim miniShift As Long
    miniShift = ReadWordLE(h + 32)
    mMiniSectorSize = 2 ^ miniShift
    ' First dir sector at +48
    mFirstDirSec = ReadDwordLE(h + 48)
    ' Mini-stream cutoff at +56
    mMiniCutoff = ReadDwordLE(h + 56)
    ' First mini FAT sector at +60
    mFirstMiniFatSec = ReadDwordLE(h + 60)
    ' Num mini FAT sectors at +64
    mNumMiniFatSec = ReadDwordLE(h + 64)
    ' First DIFAT sector at +68 (typically ENDOFCHAIN for small files)
    mFirstDifatSec = ReadDwordLE(h + 68)
    ' Num DIFAT sectors at +72
    mNumDifatSec = ReadDwordLE(h + 72)
    ReadCfbfHeader = True
End Function

' Convert a sector number to a byte offset inside the file
Private Function SectorOffset(ByVal sec As Long) As Long
    SectorOffset = mCfbfStart + CFBF_HEADER_LEN + sec * mSectorSize
End Function

' Read the FAT.  For small files (<= 109 FAT sectors) the DIFAT lives
' entirely in the header at offsets +76..+509 (109 entries).  We don't
' bother with DIFAT-sector chasing for the common case.
Private Function ReadFat() As Boolean
    Dim h As Long
    h = mCfbfStart
    ' Number of FAT sectors at +44
    Dim numFat As Long
    numFat = ReadDwordLE(h + 44)
    If numFat > 109 Then
        ' For .frx this never happens (tiny files), but bail loudly if it does.
        mParseErrorMsg = "Unsupported: DIFAT chain (numFat=" & numFat & ")"
        Exit Function
    End If
    ' Each FAT sector holds mSectorSize/4 entries.
    Dim entriesPerSec As Long
    entriesPerSec = mSectorSize \ 4
    mFatLen = numFat * entriesPerSec
    ReDim mFat(0 To mFatLen - 1)
    Dim i As Long, secIdx As Long, secNum As Long, off As Long, j As Long
    For i = 0 To numFat - 1
        secNum = ReadDwordLE(h + 76 + i * 4)
        off = SectorOffset(secNum)
        If off + mSectorSize > mBytesLen Then
            mParseErrorMsg = "FAT sector " & secNum & " past EOF"
            Exit Function
        End If
        For j = 0 To entriesPerSec - 1
            mFat(i * entriesPerSec + j) = ReadDwordLE(off + j * 4)
        Next j
    Next i
    ReadFat = True
End Function

Private Function ReadMiniFat() As Boolean
    If mNumMiniFatSec = 0 Then
        mMiniFatLen = 0
        ReadMiniFat = True
        Exit Function
    End If
    Dim entriesPerSec As Long
    entriesPerSec = mSectorSize \ 4
    mMiniFatLen = mNumMiniFatSec * entriesPerSec
    ReDim mMiniFat(0 To mMiniFatLen - 1)
    Dim sec As Long, idx As Long, off As Long, j As Long
    sec = mFirstMiniFatSec
    idx = 0
    Do While sec >= 0 And sec < mFatLen
        off = SectorOffset(sec)
        If off + mSectorSize > mBytesLen Then
            mParseErrorMsg = "MiniFAT sector " & sec & " past EOF"
            Exit Function
        End If
        For j = 0 To entriesPerSec - 1
            If idx >= mMiniFatLen Then Exit For
            mMiniFat(idx) = ReadDwordLE(off + j * 4)
            idx = idx + 1
        Next j
        sec = mFat(sec)
    Loop
    ReadMiniFat = True
End Function

' Walk the directory chain (starts at mFirstDirSec, follows FAT) and
' pull out start-sector + size for Root Entry, 'f', 'o'.  We don't care
' about CompObj.  Names are UTF-16LE in the first 64 bytes.
Private Function WalkDirectory(ByRef fStartSec As Long, ByRef fSize As Long, ByRef foundF As Boolean, _
                               ByRef oStartSec As Long, ByRef oSize As Long, ByRef foundO As Boolean) As Boolean
    foundF = False: foundO = False
    Dim sec As Long
    sec = mFirstDirSec
    Dim entriesPerSec As Long
    entriesPerSec = mSectorSize \ CFBF_DIR_ENTRY_LEN  ' 4 for 512-byte sectors
    Do While sec >= 0
        Dim off As Long
        off = SectorOffset(sec)
        If off + mSectorSize > mBytesLen Then
            mParseErrorMsg = "Dir sector " & sec & " past EOF"
            Exit Function
        End If
        Dim i As Long
        For i = 0 To entriesPerSec - 1
            Dim base As Long
            base = off + i * CFBF_DIR_ENTRY_LEN
            Dim objType As Byte
            objType = mBytes(base + 66)
            If objType = 0 Then GoTo NextEntry  ' empty
            Dim name As String
            name = ReadDirName(base)
            Dim startSec As Long, streamSize As Long
            startSec = ReadDwordLE(base + 116)
            streamSize = ReadDwordLE(base + 120)
            ' Type 5 = Root, Type 2 = Stream, Type 1 = Storage
            If objType = 5 Then
                ' Root Entry: its start sector is first mini-stream sector,
                ' its size is the mini-stream total size.
                mMiniStreamSec = startSec
                mMiniStreamSize = streamSize
            ElseIf objType = 2 Then
                If name = "f" Then
                    fStartSec = startSec: fSize = streamSize: foundF = True
                ElseIf name = "o" Then
                    oStartSec = startSec: oSize = streamSize: foundO = True
                End If
            End If
NextEntry:
        Next i
        If sec >= mFatLen Then Exit Do
        sec = mFat(sec)
        If sec = CFBF_ENDOFCHAIN Then Exit Do
    Loop
    WalkDirectory = True
End Function

' Read a NUL-terminated UTF-16LE name from a directory entry.
' Field is "DirEntryNameLength" at +64 (WORD, in bytes, includes \0).
Private Function ReadDirName(ByVal base As Long) As String
    Dim nameLenBytes As Long
    nameLenBytes = ReadWordLE(base + 64)
    If nameLenBytes < 2 Then ReadDirName = "": Exit Function
    ' Strip trailing \0
    Dim chars As Long
    chars = (nameLenBytes - 2) \ 2
    Dim s As String, j As Long, ch As Long
    For j = 0 To chars - 1
        ch = mBytes(base + j * 2) + mBytes(base + j * 2 + 1) * 256&
        s = s & ChrW$(ch)
    Next j
    ReadDirName = s
End Function

' Materialize a stream into mStreamBuf, with mStreamMap[i] = file offset.
' If size <= miniCutoff, walk the mini-FAT through the mini-stream;
' otherwise walk the regular FAT.
Private Function LoadStream(ByVal startSec As Long, ByVal streamSize As Long) As Boolean
    mStreamLen = streamSize
    mStreamPos = 0
    If streamSize <= 0 Then
        ReDim mStreamBuf(0 To 0)
        ReDim mStreamMap(0 To 0)
        LoadStream = True
        Exit Function
    End If
    ReDim mStreamBuf(0 To streamSize - 1)
    ReDim mStreamMap(0 To streamSize - 1)
    Dim isMini As Boolean
    isMini = (streamSize < mMiniCutoff)
    Dim chunk As Long
    chunk = IIf(isMini, mMiniSectorSize, mSectorSize)
    Dim sec As Long
    sec = startSec
    Dim copied As Long
    copied = 0
    Do While copied < streamSize And sec >= 0
        Dim off As Long
        off = MiniOrSectorOffset(sec, isMini)
        If off < 0 Then
            mParseErrorMsg = "Bad sector " & sec & " (isMini=" & isMini & ")"
            Exit Function
        End If
        Dim toCopy As Long
        toCopy = chunk
        If copied + toCopy > streamSize Then toCopy = streamSize - copied
        Dim k As Long
        For k = 0 To toCopy - 1
            mStreamBuf(copied + k) = mBytes(off + k)
            mStreamMap(copied + k) = off + k
        Next k
        copied = copied + toCopy
        ' Walk to next sector
        If isMini Then
            If sec >= mMiniFatLen Then Exit Do
            sec = mMiniFat(sec)
        Else
            If sec >= mFatLen Then Exit Do
            sec = mFat(sec)
        End If
        If sec = CFBF_ENDOFCHAIN Then Exit Do
    Loop
    LoadStream = True
End Function

' Resolve sector number -> file byte offset.  For mini sectors, walk
' the mini-stream chain (Root Entry) by sector number.
Private Function MiniOrSectorOffset(ByVal sec As Long, ByVal isMini As Boolean) As Long
    If Not isMini Then
        MiniOrSectorOffset = SectorOffset(sec)
        Exit Function
    End If
    ' Mini stream: walk Root Entry's FAT chain to find the sector that
    ' covers mini-sector `sec`.  Each big sector holds (mSectorSize/mMiniSectorSize) mini sectors.
    Dim miniPerBig As Long
    miniPerBig = mSectorSize \ mMiniSectorSize
    Dim bigIdx As Long
    bigIdx = sec \ miniPerBig
    Dim miniOff As Long
    miniOff = (sec Mod miniPerBig) * mMiniSectorSize
    Dim curSec As Long, curIdx As Long
    curSec = mMiniStreamSec: curIdx = 0
    Do While curIdx < bigIdx And curSec >= 0 And curSec < mFatLen
        curSec = mFat(curSec)
        curIdx = curIdx + 1
        If curSec = CFBF_ENDOFCHAIN Then
            MiniOrSectorOffset = -1
            Exit Function
        End If
    Loop
    If curSec < 0 Then MiniOrSectorOffset = -1: Exit Function
    MiniOrSectorOffset = SectorOffset(curSec) + miniOff
End Function


'======================================================================
'  STREAM READERS (operate on mStreamBuf with cursor mStreamPos)
'======================================================================

Private Function StreamRead(ByVal n As Long) As Boolean
    ' If padding mode is active, flag the alignment gap as padding first.
    If mPaddingActive Then PadIfNeeded n
    If mStreamPos + n > mStreamLen Then
        mParseErrorMsg = "Stream read past end (pos=" & mStreamPos & " n=" & n & " len=" & mStreamLen & ")"
        Exit Function
    End If
    mStreamPos = mStreamPos + n
    StreamRead = True
End Function

Private Function ReadByte() As Long
    If Not StreamRead(1) Then Exit Function
    ReadByte = mStreamBuf(mStreamPos - 1)
End Function

Private Function ReadWord() As Long
    If Not StreamRead(2) Then Exit Function
    ReadWord = mStreamBuf(mStreamPos - 2) + mStreamBuf(mStreamPos - 1) * 256&
End Function

Private Function ReadDword() As Double
    ' Unsigned DWORD as Double (avoid Long overflow)
    If Not StreamRead(4) Then Exit Function
    Dim p As Long: p = mStreamPos - 4
    ReadDword = CDbl(mStreamBuf(p)) _
              + CDbl(mStreamBuf(p + 1)) * 256# _
              + CDbl(mStreamBuf(p + 2)) * 65536# _
              + CDbl(mStreamBuf(p + 3)) * 16777216#
End Function

' Read a count-of-bytes/chars field (CountOfBytesWithCompressionFlag).
Private Function ReadCountWithCompressionFlag() As Long
    Dim v As Double
    v = ReadDword()
    ' Mask off the high compression flag bit (0x80000000)
    If v >= 2147483648# Then v = v - 2147483648#
    ReadCountWithCompressionFlag = CLng(v)
End Function

'-- Padding semantics (matches ExtendedStream._pad in oleform.py) -----
'   _pad(start, size=4) -> if (pos - start) % size != 0, consume the
'                          remainder.  Those consumed bytes are padding.
Private Sub PadIfNeeded(ByVal nextRead As Long)
    Dim off As Long
    off = (mStreamPos - mPadStart) Mod nextRead
    If off <> 0 Then
        Dim need As Long
        need = nextRead - off
        FlagPadding mStreamPos, need
        mStreamPos = mStreamPos + need
    End If
End Sub

' Align to 4 bytes from a given start (matches will_pad/__exit__).
Private Sub PadTo4(ByVal startPos As Long)
    Dim off As Long
    off = (mStreamPos - startPos) Mod 4
    If off <> 0 Then
        Dim need As Long
        need = 4 - off
        FlagPadding mStreamPos, need
        mStreamPos = mStreamPos + need
    End If
End Sub

' Same as PadTo4 but the caller passes the "alignment anchor" -- not
' tied to mPadStart.  Used for ExtraDataBlock string alignment where
' strings are 4-aligned relative to the ExtraDataBlock start.
Private Sub PadTo4Relative(ByVal alignAnchor As Long)
    PadTo4 alignAnchor
End Sub

' Mark `n` bytes starting at stream offset `streamOff` as padding (zero on write-back).
Private Sub FlagPadding(ByVal streamOff As Long, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        If streamOff + i < mStreamLen Then
            mPadMask(mStreamMap(streamOff + i)) = 1
        End If
    Next i
End Sub

'-- padded_struct() context manager -----------------------------------
Private Sub PushPadded()
    If mPadStackTop >= PAD_STACK_MAX Then
        mParseErrorMsg = "Padded-struct stack overflow"
        Exit Sub
    End If
    mPadStack(mPadStackTop) = mPadStart
    mPadStackActive(mPadStackTop) = mPaddingActive
    mPadStackTop = mPadStackTop + 1
    mPaddingActive = True
    mPadStart = mStreamPos
End Sub

Private Sub PopPadded()
    ' On exit: flush alignment to 4 from mPadStart, restore prior state.
    PadTo4 mPadStart
    mPadStackTop = mPadStackTop - 1
    mPaddingActive = mPadStackActive(mPadStackTop)
    mPadStart = mPadStack(mPadStackTop)
End Sub

'-- will_jump_to: read cbRecord bytes from here.  After consuming what
'   we want, fill the rest to cbRecord (those trailing bytes are padding
'   between MS-OFORMS records and ARE non-deterministic).
Private Function JumpTo(ByVal cbRecord As Long, ByVal startPos As Long) As Boolean
    Dim consumed As Long
    consumed = mStreamPos - startPos
    If consumed > cbRecord Then
        mParseErrorMsg = "Bad jump: consumed=" & consumed & " > cbRecord=" & cbRecord & " at pos=" & mStreamPos
        Exit Function
    End If
    Dim remaining As Long
    remaining = cbRecord - consumed
    If remaining > 0 Then
        FlagPadding mStreamPos, remaining
        mStreamPos = mStreamPos + remaining
    End If
    JumpTo = True
End Function


'======================================================================
'  PRIMITIVE FILE-LEVEL READERS (operate on mBytes(), file offsets)
'======================================================================
Private Function ReadWordLE(ByVal off As Long) As Long
    ReadWordLE = mBytes(off) + mBytes(off + 1) * 256&
End Function

Private Function ReadDwordLE(ByVal off As Long) As Long
    ' Returns signed Long.  ENDOFCHAIN/FREESECT/etc. are negative when interpreted signed,
    ' which matches our constants above.
    Dim v As Double
    v = CDbl(mBytes(off)) _
      + CDbl(mBytes(off + 1)) * 256# _
      + CDbl(mBytes(off + 2)) * 65536# _
      + CDbl(mBytes(off + 3)) * 16777216#
    If v >= 2147483648# Then v = v - 4294967296#
    ReadDwordLE = CLng(v)
End Function


'======================================================================
'  MS-OFORMS PARSERS
'======================================================================

'  PropMask helpers
'  We model PropMask as a Long containing up to 32 bits.  Bit i corresponds
'  to the i-th entry in the propmask's name list.
Private Function MaskBit(ByVal mask As Double, ByVal bit As Long) As Long
    ' Returns 0 or 1.  mask is Double to hold unsigned 32-bit.
    ' Avoid CLng (overflows for mask >= 2^31): use mod-2 on a Double.
    Dim shifted As Double
    shifted = Int(mask / (2 ^ bit))
    If (shifted - Int(shifted / 2) * 2) >= 1 Then
        MaskBit = 1
    Else
        MaskBit = 0
    End If
End Function

' Consume a list of (bitIndex, size) pairs from the propmask.  Each entry
' is read only if the corresponding mask bit is set.  This mimics
' Mask.consume() in oleform.py.
Private Function ConsumeProps(ByVal mask As Double, ByRef bits() As Long, ByRef sizes() As Long, ByVal n As Long) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If MaskBit(mask, bits(i)) Then
            If Not StreamRead(sizes(i)) Then Exit Function
        End If
    Next i
    ConsumeProps = True
End Function


'  Parse the 'f' stream: FormControl record + array of OleSiteConcreteControl.
'  Returns the ClsidCacheIndex per site so we know what to expect in 'o'.
Private Function ParseFormControlStream(ByRef siteClsids() As Long, ByRef siteCount As Long) As Boolean
    Dim clsids() As Long, n As Long
    n = 0
    ReDim clsids(0 To 127)
    If Not ConsumeFormControl(clsids, n, True) Then Exit Function
    siteCount = n
    If n > 0 Then
        ReDim siteClsids(0 To n - 1)
        Dim i As Long
        For i = 0 To n - 1
            siteClsids(i) = clsids(i)
        Next i
    End If
    ParseFormControlStream = True
End Function

' Parse the 'o' stream: one ObjectStreamRecord per site, dispatched by ClsidCacheIndex.
Private Function ParseObjectStream(ByRef siteClsids() As Long, ByVal siteCount As Long) As Boolean
    Dim i As Long
    For i = 0 To siteCount - 1
        Select Case siteClsids(i)
            Case 7
                Dim dummy() As Long, n As Long
                n = 0
                ReDim dummy(0 To 0)
                If Not ConsumeFormControl(dummy, n, False) Then Exit Function
            Case 12
                If Not ConsumeImageControl() Then Exit Function
            Case 14
                Dim dummy2() As Long, n2 As Long
                n2 = 0
                ReDim dummy2(0 To 0)
                If Not ConsumeFormControl(dummy2, n2, False) Then Exit Function
            Case 15, 23, 24, 25, 26, 27, 28
                If Not ConsumeMorphDataControl() Then Exit Function
            Case 16
                If Not ConsumeSpinButtonControl() Then Exit Function
            Case 17
                If Not ConsumeCommandButtonControl() Then Exit Function
            Case 18
                If Not ConsumeTabStripControl() Then Exit Function
            Case 21
                If Not ConsumeLabelControl() Then Exit Function
            Case 47
                If Not ConsumeScrollBarControl() Then Exit Function
            Case 57
                Dim dummy3() As Long, n3 As Long
                n3 = 0
                ReDim dummy3(0 To 0)
                If Not ConsumeFormControl(dummy3, n3, False) Then Exit Function
            Case Else
                mParseErrorMsg = "Unsupported ClsidCacheIndex: " & siteClsids(i)
                Exit Function
        End Select
    Next i
    ParseObjectStream = True
End Function


'-- FormControl: [MS-OFORMS] 2.2.10.1 ---------------------------------
Private Function ConsumeFormControl(ByRef siteClsids() As Long, ByRef siteCount As Long, ByVal collectSites As Boolean) As Boolean
    ' Version bytes
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    ' (We don't enforce 0,4 strictly -- but the standard says so.)
    ' cbForm
    Dim cbform As Long
    cbform = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    ' DataBlock
    Dim mask As Double
    mask = ReadDword()
    ' Build bit list:
    '   bit 1 fBackColor (4), bit 2 fForeColor (4), bit 3 fNextAvailableID (4)
    Dim bits() As Long, sizes() As Long
    ReDim bits(0 To 2): ReDim sizes(0 To 2)
    bits(0) = 1: sizes(0) = 4
    bits(1) = 2: sizes(1) = 4
    bits(2) = 3: sizes(2) = 4
    If Not ConsumeProps(mask, bits, sizes, 3) Then Exit Function
    ' fBooleanProperties bits: 6 OR 7 (Python uses both)
    Dim hasBoolProps As Long
    hasBoolProps = MaskBit(mask, 6) Or MaskBit(mask, 7)
    Dim formDontSaveClass As Long
    formDontSaveClass = 0
    If hasBoolProps Then
        Dim boolProps As Double
        boolProps = ReadDword()
        ' bit 15 = DONTSAVECLASSTABLE
        If MaskBit(boolProps, 15) Then formDontSaveClass = 1
    End If
    ' Skip the rest of the FormDataBlock + FormExtraDataBlock via JumpTo
    If Not JumpTo(cbform, startPos) Then Exit Function

    ' FormStreamData: GuidAndPicture for MouseIcon, GuidAndFont for Font, GuidAndPicture for Picture
    If MaskBit(mask, 15) Then  ' fMouseIcon
        If Not ConsumeGuidAndPicture() Then Exit Function
    End If
    If MaskBit(mask, 20) Then  ' fFont
        If Not ConsumeGuidAndFont() Then Exit Function
    End If
    If MaskBit(mask, 21) Then  ' fPicture
        If Not ConsumeGuidAndPicture() Then Exit Function
    End If

    ' FormSiteData
    Dim countSiteClass As Long
    If formDontSaveClass = 0 Then
        countSiteClass = ReadWord()
        Dim i As Long
        For i = 1 To countSiteClass
            If Not ConsumeSiteClassInfo() Then Exit Function
        Next i
    End If

    Dim countOfSites As Double, countOfBytes As Double
    countOfSites = ReadDword()
    countOfBytes = ReadDword()
    Dim cob As Long
    cob = CLng(countOfBytes)
    Dim cobStart As Long
    cobStart = mStreamPos

    ' will_pad starting here, while consuming SiteDepthsAndTypes for all sites
    Dim padStart As Long
    padStart = mStreamPos
    Dim remaining As Long
    remaining = CLng(countOfSites)
    Do While remaining > 0
        Dim consumed As Long
        consumed = ConsumeFormObjectDepthTypeCount()
        If consumed < 0 Then Exit Function
        remaining = remaining - consumed
    Loop
    ' Padding at end of SiteDepthsAndTypes block (to 4)
    PadTo4 padStart

    ' Now CountOfSites entries of OleSiteConcreteControl
    For i = 1 To CLng(countOfSites)
        Dim clsid As Long
        If Not ConsumeOleSiteConcreteControl(clsid) Then Exit Function
        If collectSites Then
            If siteCount >= UBound(siteClsids) Then
                ReDim Preserve siteClsids(0 To siteCount * 2 + 16)
            End If
            siteClsids(siteCount) = clsid
            siteCount = siteCount + 1
        End If
    Next i

    ' will_jump_to(CountOfBytes): consume any remaining bytes inside the
    ' SiteData block (those are inter-site padding).
    If Not JumpTo(cob, cobStart) Then Exit Function

    ConsumeFormControl = True
End Function

Private Function ConsumeFormObjectDepthTypeCount() As Long
    ' Returns "count consumed from remaining" (1 or N).  Negative = error.
    Dim depth As Long, mixed As Long
    depth = ReadByte()
    mixed = ReadByte()
    If (mixed And &H80&) <> 0 Then
        ' Site type byte follows; expect 1
        Dim st As Long
        st = ReadByte()
        If st <> 1 Then
            mParseErrorMsg = "FormObjectDepthTypeCount: expected SITE_TYPE=1, got " & st
            ConsumeFormObjectDepthTypeCount = -1
            Exit Function
        End If
        ConsumeFormObjectDepthTypeCount = mixed Xor &H80&
        Exit Function
    End If
    If mixed <> 1 Then
        mParseErrorMsg = "FormObjectDepthTypeCount: expected 1, got " & mixed
        ConsumeFormObjectDepthTypeCount = -1
        Exit Function
    End If
    ConsumeFormObjectDepthTypeCount = 1
End Function

Private Function ConsumeSiteClassInfo() As Boolean
    Dim ver As Long
    ver = ReadWord()
    Dim cb As Long
    cb = ReadWord()
    If Not StreamRead(cb) Then Exit Function
    ConsumeSiteClassInfo = True
End Function

'-- OleSiteConcreteControl: [MS-OFORMS] 2.2.10.12.1 -------------------
Private Function ConsumeOleSiteConcreteControl(ByRef outClsidCacheIndex As Long) As Boolean
    Dim ver As Long
    ver = ReadWord()
    Dim cbSite As Long
    cbSite = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()

    ' SiteDataBlock (padded_struct) -- bit indices from SitePropMask:
    '   0 fName, 1 fTag, 2 fID, 3 fHelpContextID, 4 fBitFlags, 5 fObjectStreamSize,
    '   6 fTabIndex, 7 fClsidCacheIndex, 8 fPosition, 9 fGroupID, 10 Unused1,
    '   11 fControlTipText, 12 fRuntimeLicKey, 13 fControlSource, 14 fRowSource
    PushPadded
    Dim nameLen As Long, tagLen As Long, ctlTipLen As Long
    nameLen = 0: tagLen = 0: ctlTipLen = 0
    Dim id As Double: id = 0
    Dim tabindex As Long: tabindex = 0
    Dim clsidCache As Long: clsidCache = 0

    If MaskBit(mask, 0) Then nameLen = ReadCountWithCompressionFlag()
    If MaskBit(mask, 1) Then tagLen = ReadCountWithCompressionFlag()
    If MaskBit(mask, 2) Then id = ReadDword()
    ' (fHelpContextID, fBitFlags, fObjectStreamSize) -- each 4
    Dim bits() As Long, sizes() As Long
    ReDim bits(0 To 2): ReDim sizes(0 To 2)
    bits(0) = 3: sizes(0) = 4
    bits(1) = 4: sizes(1) = 4
    bits(2) = 5: sizes(2) = 4
    If Not ConsumeProps(mask, bits, sizes, 3) Then Exit Function

    If MaskBit(mask, 6) Then tabindex = ReadWord()
    If MaskBit(mask, 7) Then clsidCache = ReadWord()
    If MaskBit(mask, 9) Then
        If Not StreamRead(2) Then Exit Function   ' fGroupID
    End If
    If MaskBit(mask, 11) Then ctlTipLen = ReadCountWithCompressionFlag()
    ' (fRuntimeLicKey, fControlSource, fRowSource) -- each 4
    ReDim bits(0 To 2): ReDim sizes(0 To 2)
    bits(0) = 12: sizes(0) = 4
    bits(1) = 13: sizes(1) = 4
    bits(2) = 14: sizes(2) = 4
    If Not ConsumeProps(mask, bits, sizes, 3) Then Exit Function
    PopPadded

    ' SiteExtraDataBlock: name, tag, optional position(8), control_tip_text
    If nameLen > 0 Then
        If Not StreamRead(nameLen) Then Exit Function
    End If
    If tagLen > 0 Then
        If Not StreamRead(tagLen) Then Exit Function
    End If
    If MaskBit(mask, 8) Then
        If Not StreamRead(8) Then Exit Function   ' SitePosition
    End If
    If ctlTipLen > 0 Then
        If Not StreamRead(ctlTipLen) Then Exit Function
    End If

    outClsidCacheIndex = clsidCache
    ' Trailing padding inside cbSite
    If Not JumpTo(cbSite, startPos) Then Exit Function
    ConsumeOleSiteConcreteControl = True
End Function

'-- TextProps: [MS-OFORMS] 2.3.1 ---------------------------------------
' Excel-written TextProps body (empirical, mask=0x35, cb=24):
'   [PropMask 4] [Count 4] [Field 4] [Field 4] [FaceName 6] [pad 2] = 24
'
' We can't trust the per-bit field sizes without a definitive MS-OFORMS
' spec, BUT we know:
'   - PropMask is 4 bytes.
'   - If a FaceName is present, it ends right at (cb - trailing-pad)
'     where the trailing pad is whatever's needed to make the string-end
'     reach a 4-byte boundary from startPos.
'   - FaceName length is encoded as a CountWithCompressionFlag DWORD
'     somewhere in the binary-props area.
'
' Strategy: skip past PropMask, scan forward looking for a
' CountWithCompressionFlag DWORD whose `count` value (low 31 bits)
' equals (cb - currentPos - count_padToEnd).  In practice the count
' is always the very first DWORD after PropMask when fFontName is set
' (mask bit 0 in the layout below), so just probe that.
'
' If we find a valid count, the FaceName lives at the end of cb,
' aligned to 4.  We pad-fill anything between the count's position+4
' and the string start as opaque (binary props -- left alone), and
' flag only the trailing alignment pad as padding.
Private Function ConsumeTextProps() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    If cb <= 0 Then ConsumeTextProps = True: Exit Function

    Dim mask As Double
    mask = ReadDword()      ' 4 bytes
    Dim faceNameLen As Long: faceNameLen = 0
    If MaskBit(mask, 0) Then
        Dim count As Long
        count = ReadCountWithCompressionFlag()
        If count > 0 And count <= cb - 8 Then faceNameLen = count
    End If
    ' Skip remaining binary props as opaque (read but don't flag).  The
    ' FaceName, if present, lives at the very end of cb (aligned to 4),
    ' so its starting offset relative to startPos = cb - faceNameLen,
    ' rounded down to a multiple of 4 from the StreamPos side.
    If faceNameLen > 0 Then
        ' Compute string start, then skip-without-flag to that offset.
        Dim stringStart As Long
        stringStart = startPos + (cb - faceNameLen)
        ' Round DOWN to 4-aligned (relative to startPos) -- Excel pads
        ' to 4 within the cb scope after the FaceName.
        Dim relOff As Long
        relOff = stringStart - startPos
        Dim mod4 As Long
        mod4 = relOff Mod 4
        If mod4 <> 0 Then stringStart = stringStart + (4 - mod4)
        If stringStart < mStreamPos Then stringStart = mStreamPos
        ' Skip-without-flag (preserve binary props).
        If stringStart - mStreamPos > 0 Then
            If Not StreamReadOpaque(stringStart - mStreamPos) Then Exit Function
        End If
        If Not StreamReadOpaque(faceNameLen) Then Exit Function
    End If
    ' Fill cb -- trailing alignment is the only true padding.
    If Not JumpTo(cb, startPos) Then Exit Function
    ConsumeTextProps = True
End Function

' Like StreamRead but never flags padding (used inside TextProps for
' opaque binary props between the count and the FaceName).
Private Function StreamReadOpaque(ByVal n As Long) As Boolean
    If n <= 0 Then StreamReadOpaque = True: Exit Function
    If mStreamPos + n > mStreamLen Then
        mParseErrorMsg = "Stream read past end (pos=" & mStreamPos & " n=" & n & ")"
        Exit Function
    End If
    mStreamPos = mStreamPos + n
    StreamReadOpaque = True
End Function

'-- GuidAndFont: [MS-OFORMS] 2.4.7 ------------------------------------
Private Function ConsumeGuidAndFont() As Boolean
    ' Read 16-byte GUID
    Dim g1 As Double, g2 As Long, g3 As Long
    g1 = ReadDword()
    g2 = ReadWord()
    g3 = ReadWord()
    ' Next 8 bytes (big-endian Q) - just read them, identify by first DWORD pattern
    Dim q1 As Long, q2 As Long, q3 As Long, q4 As Long
    q1 = ReadByte(): q2 = ReadByte(): q3 = ReadByte(): q4 = ReadByte()
    Dim q5 As Long, q6 As Long, q7 As Long, q8 As Long
    q5 = ReadByte(): q6 = ReadByte(): q7 = ReadByte(): q8 = ReadByte()
    ' StdFont GUID: 0BE35203-8F91-11CE-9DE3-00AA004BB851
    ' g1=0x0BE35203 (199447043), g2=0x8F91 (36753), g3=0x11CE (4558)
    If CLng(g1) = 199447043 And g2 = 36753 And g3 = 4558 _
       And q1 = &H9D And q2 = &HE3 And q3 = &H0 And q4 = &HAA _
       And q5 = &H0 And q6 = &H4B And q7 = &HB8 And q8 = &H51 Then
        ' StdFont [2.4.12]
        Dim v As Long
        v = ReadByte()
        If Not StreamRead(9) Then Exit Function   ' sCharset(2)+bFlags(1)+sWeight(2)+ulHeight(4)
        Dim bFaceLen As Long
        bFaceLen = ReadByte()
        If bFaceLen > 0 Then
            If Not StreamRead(bFaceLen) Then Exit Function
        End If
    ElseIf CLng(g1) = -1346237176 + 4294967296# - 4294967296# And False Then
        ' Unreachable branch (placeholder for TextProps GUID), see below
    Else
        ' Try the TextProps GUID: AFC20920-DA4E-11CE-B943-00AA006887B4
        ' (oletools encodes high-bit g1 as 2948729120 which is 0xAFC20920 unsigned)
        ' g1 came back as Double from ReadDword -- compare unsigned.
        If g1 = 2948729120# And g2 = 55886 And g3 = 4558 Then
            If Not ConsumeTextProps() Then Exit Function
        Else
            mParseErrorMsg = "GuidAndFont: unknown GUID"
            Exit Function
        End If
    End If
    ConsumeGuidAndFont = True
End Function

'-- GuidAndPicture: [MS-OFORMS] 2.4.8 ---------------------------------
Private Function ConsumeGuidAndPicture() As Boolean
    ' 16-byte GUID + StdPicture
    Dim g1 As Double, g2 As Long, g3 As Long
    g1 = ReadDword(): g2 = ReadWord(): g3 = ReadWord()
    ' 8 byte tail of GUID
    Dim i As Long
    For i = 1 To 8: If Not StreamRead(1) Then Exit Function: Next i
    ' StdPicture preamble: 0x0000746C
    Dim preamble As Double
    preamble = ReadDword()
    If CLng(preamble) <> &H746C Then
        mParseErrorMsg = "StdPicture: bad preamble 0x" & Hex$(CLng(preamble))
        Exit Function
    End If
    Dim sz As Double
    sz = ReadDword()
    If sz > 0 Then
        If Not StreamRead(CLng(sz)) Then Exit Function
    End If
    ConsumeGuidAndPicture = True
End Function

'-- MorphDataControl: [MS-OFORMS] 2.2.5.1 -----------------------------
Private Function ConsumeMorphDataControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    ' PropMask is QWORD here (8 bytes) -- 33 bits used
    Dim m_lo As Double, m_hi As Double
    m_lo = ReadDword()
    m_hi = ReadDword()
    ' Helper for QWORD bit testing -- bit 32 is the only high-half bit we use (fGroupName)
    PushPadded
    ' Fixed props (bit list / sizes from oleform.py)
    Dim bits() As Long, sizes() As Long
    ReDim bits(0 To 19): ReDim sizes(0 To 19)
    bits(0) = 0: sizes(0) = 4   ' fVariousPropertyBits
    bits(1) = 1: sizes(1) = 4   ' fBackColor
    bits(2) = 2: sizes(2) = 4   ' fForeColor
    bits(3) = 3: sizes(3) = 4   ' fMaxLength
    bits(4) = 4: sizes(4) = 1   ' fBorderStyle
    bits(5) = 5: sizes(5) = 1   ' fScrollBars
    bits(6) = 6: sizes(6) = 1   ' fDisplayStyle
    bits(7) = 7: sizes(7) = 1   ' fMousePointer
    bits(8) = 9: sizes(8) = 2   ' fPasswordChar
    bits(9) = 10: sizes(9) = 4  ' fListWidth
    bits(10) = 11: sizes(10) = 2  ' fBoundColumn
    bits(11) = 12: sizes(11) = 2  ' fTextColumn
    bits(12) = 13: sizes(12) = 2  ' fColumnCount
    bits(13) = 14: sizes(13) = 2  ' fListRows
    bits(14) = 15: sizes(14) = 2  ' fcColumnInfo
    bits(15) = 16: sizes(15) = 1  ' fMatchEntry
    bits(16) = 17: sizes(16) = 1  ' fListStyle
    bits(17) = 18: sizes(17) = 1  ' fShowDropButtonWhen
    bits(18) = 20: sizes(18) = 1  ' fDropButtonStyle
    bits(19) = 21: sizes(19) = 1  ' fMultiSelect
    If Not ConsumeProps(m_lo, bits, sizes, 20) Then Exit Function

    Dim valueSize As Long: valueSize = 0
    Dim captionSize As Long: captionSize = 0
    Dim groupSize As Long: groupSize = 0
    If MaskBit(m_lo, 22) Then valueSize = ReadCountWithCompressionFlag()
    If MaskBit(m_lo, 23) Then captionSize = ReadCountWithCompressionFlag()

    ReDim bits(0 To 5): ReDim sizes(0 To 5)
    bits(0) = 24: sizes(0) = 4  ' fPicturePosition
    bits(1) = 25: sizes(1) = 4  ' fBorderColor
    bits(2) = 26: sizes(2) = 4  ' fSpecialEffect
    bits(3) = 27: sizes(3) = 2  ' fMouseIcon
    bits(4) = 28: sizes(4) = 2  ' fPicture
    bits(5) = 29: sizes(5) = 2  ' fAccelerator
    If Not ConsumeProps(m_lo, bits, sizes, 6) Then Exit Function

    ' fGroupName is bit 32 (high half).  m_hi bit 0.
    If (CLng(m_hi) And 1) Then groupSize = ReadCountWithCompressionFlag()
    PopPadded

    ' ExtraDataBlock: discard 8 (Size), then value, caption, group_name.
    ' Each variable-length string is followed by 4-byte alignment padding
    ' (oletools omits this; Excel actually leaves those bytes uninitialised).
    Dim extraStart As Long
    extraStart = mStreamPos
    If Not StreamRead(8) Then Exit Function
    Dim afterValue As Long, afterCaption As Long
    If valueSize > 0 Then If Not StreamRead(valueSize) Then Exit Function
    afterValue = mStreamPos
    PadTo4Relative extraStart
    If captionSize > 0 Then If Not StreamRead(captionSize) Then Exit Function
    afterCaption = mStreamPos
    PadTo4Relative extraStart
    If groupSize > 0 Then If Not StreamRead(groupSize) Then Exit Function
    PadTo4Relative extraStart

    If Not JumpTo(cb, startPos) Then Exit Function

    ' StreamData
    If MaskBit(m_lo, 27) Then
        If Not ConsumeGuidAndPicture() Then Exit Function
    End If
    If MaskBit(m_lo, 28) Then
        If Not ConsumeGuidAndPicture() Then Exit Function
    End If
    If Not ConsumeTextProps() Then Exit Function
    ConsumeMorphDataControl = True
End Function

'-- ImageControl: [MS-OFORMS] 2.2.3.1 ---------------------------------
Private Function ConsumeImageControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    If Not JumpTo(cb, startPos) Then Exit Function
    ' Bits 10 fPicture, 14 fMouseIcon
    If MaskBit(mask, 10) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If MaskBit(mask, 14) Then If Not ConsumeGuidAndPicture() Then Exit Function
    ConsumeImageControl = True
End Function

'-- CommandButtonControl: [MS-OFORMS] 2.2.1.1 -------------------------
Private Function ConsumeCommandButtonControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    If Not JumpTo(cb, startPos) Then Exit Function
    ' Bits 7 fPicture, 10 fMouseIcon
    If MaskBit(mask, 7) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If MaskBit(mask, 10) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If Not ConsumeTextProps() Then Exit Function
    ConsumeCommandButtonControl = True
End Function

'-- SpinButtonControl: [MS-OFORMS] 2.2.8.1 ----------------------------
Private Function ConsumeSpinButtonControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    If Not JumpTo(cb, startPos) Then Exit Function
    ' Bit 13 = fMouseIcon
    If MaskBit(mask, 13) Then If Not ConsumeGuidAndPicture() Then Exit Function
    ConsumeSpinButtonControl = True
End Function

'-- TabStripControl: [MS-OFORMS] 2.2.9.1 ------------------------------
Private Function ConsumeTabStripControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    Dim bits() As Long, sizes() As Long
    ReDim bits(0 To 13): ReDim sizes(0 To 13)
    bits(0) = 0: sizes(0) = 4
    bits(1) = 1: sizes(1) = 4
    bits(2) = 2: sizes(2) = 4
    bits(3) = 4: sizes(3) = 4
    bits(4) = 6: sizes(4) = 1
    bits(5) = 8: sizes(5) = 4
    bits(6) = 9: sizes(6) = 4
    bits(7) = 11: sizes(7) = 4
    bits(8) = 12: sizes(8) = 4
    bits(9) = 15: sizes(9) = 4
    bits(10) = 17: sizes(10) = 4
    bits(11) = 18: sizes(11) = 4
    bits(12) = 20: sizes(12) = 4
    bits(13) = 21: sizes(13) = 4
    If Not ConsumeProps(mask, bits, sizes, 14) Then Exit Function
    Dim tabData As Long: tabData = 0
    If MaskBit(mask, 22) Then tabData = CLng(ReadDword())
    If Not JumpTo(cb, startPos) Then Exit Function
    If MaskBit(mask, 24) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If Not ConsumeTextProps() Then Exit Function
    Dim i As Long
    For i = 1 To tabData
        If Not StreamRead(4) Then Exit Function
    Next i
    ConsumeTabStripControl = True
End Function

'-- LabelControl: [MS-OFORMS] 2.2.4.1 ---------------------------------
Private Function ConsumeLabelControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    PushPadded
    Dim bits() As Long, sizes() As Long
    ReDim bits(0 To 2): ReDim sizes(0 To 2)
    bits(0) = 0: sizes(0) = 4
    bits(1) = 1: sizes(1) = 4
    bits(2) = 2: sizes(2) = 4
    If Not ConsumeProps(mask, bits, sizes, 3) Then Exit Function
    Dim captionSize As Long: captionSize = 0
    If MaskBit(mask, 3) Then captionSize = ReadCountWithCompressionFlag()
    ReDim bits(0 To 7): ReDim sizes(0 To 7)
    bits(0) = 4: sizes(0) = 4   ' fPicturePosition
    bits(1) = 6: sizes(1) = 1   ' fMousePointer
    bits(2) = 7: sizes(2) = 4   ' fBorderColor
    bits(3) = 8: sizes(3) = 2   ' fBorderStyle
    bits(4) = 9: sizes(4) = 2   ' fSpecialEffect
    bits(5) = 10: sizes(5) = 2  ' fPicture
    bits(6) = 11: sizes(6) = 2  ' fAccelerator
    bits(7) = 12: sizes(7) = 2  ' fMouseIcon
    If Not ConsumeProps(mask, bits, sizes, 8) Then Exit Function
    PopPadded
    ' ExtraDataBlock: caption + (4-aligned) fSize 8.
    ' oleform.py reads the 8 immediately after caption -- that's wrong
    ' for Excel-written .frx, which pads caption to 4 before fSize.
    Dim extraStart As Long
    extraStart = mStreamPos
    If captionSize > 0 Then If Not StreamRead(captionSize) Then Exit Function
    PadTo4Relative extraStart
    If Not StreamRead(8) Then Exit Function
    If Not JumpTo(cb, startPos) Then Exit Function
    If MaskBit(mask, 10) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If MaskBit(mask, 12) Then If Not ConsumeGuidAndPicture() Then Exit Function
    If Not ConsumeTextProps() Then Exit Function
    ConsumeLabelControl = True
End Function

'-- ScrollBarControl: [MS-OFORMS] 2.2.7.1 -----------------------------
Private Function ConsumeScrollBarControl() As Boolean
    Dim ver0 As Long, ver1 As Long
    ver0 = ReadByte(): ver1 = ReadByte()
    Dim cb As Long
    cb = ReadWord()
    Dim startPos As Long
    startPos = mStreamPos
    Dim mask As Double
    mask = ReadDword()
    If Not JumpTo(cb, startPos) Then Exit Function
    If MaskBit(mask, 16) Then If Not ConsumeGuidAndPicture() Then Exit Function
    ConsumeScrollBarControl = True
End Function
