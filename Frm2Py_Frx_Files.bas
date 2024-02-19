Attribute VB_Name = "mod_Frm2Py_Frx_Files"
'
' Attributions:
'
'   A great deal of the work for reading these FRX files was
'   inspired from prior work by LaVolpe, found here:
'       https://www.vbforums.com/showthread.php?778707-VB6-Image-Recovery-from-Project-Files
'
'   And, the code for saving PNG files was
'   inspired from prior work by Dilettante, found here:
'       https://www.vbforums.com/showthread.php?808301-VB6-PicSave-Simple-SavePicture-as-GIF-PNG-JPEG
'
'
Option Explicit
'
Private Type IID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(7&)       As Byte
End Type
Private Type EncoderParameter
    EncoderGUID     As IID
    NumberOfValues  As Long
    Type            As Long
    pValue          As Long
End Type
'
Private Type EncoderParameters
    Count As Long 'Must always be set to 0 or 1 here, we have just one declared below.
    Parameter As EncoderParameter
End Type
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef clsid As IID) As Long
Private Declare Function SHCreateStreamOnFile Lib "shlwapi" Alias "SHCreateStreamOnFileW" (ByVal pszFile As Long, ByVal grfMode As Long, ByRef stm As IUnknown) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal iSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
'
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, ByRef pBitMap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal gdipImage As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal gdipImage As Long, ByVal Stream As IUnknown, ByRef clsidEncoder As IID, ByVal pEncoderParams As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef gdipToken As Long, ByRef StartupInput As GdiplusStartupInput, ByVal pStartupOutput As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal gdipToken As Long) As Long
'
Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type
'
Private ImageFormatGIF      As IID
Private ImageFormatJPEG     As IID
Private ImageFormatPNG      As IID
Private EncoderQuality      As IID
Private gdipStartupInput    As GdiplusStartupInput
Private gdipToken           As Long
Private IID_IPicture        As IID
'
Private miIconCount         As Long
Private miPngCount          As Long
'


Public Function FrxTest() As String
    ' This is just for testing from Debug window.
    gsInputFileSpec = "C:\Users\Elroy\Desktop\VB6_FRM_to_PY\TESTING\Main4.frm"
    gsOutputFilePath = "C:\Users\Elroy\Desktop\VB6_FRM_to_PY\TESTING\"
    gsOutputFileName = "Main4.py"
    gsOutputFileSpec = gsOutputFilePath & gsOutputFileName
    gsOutputFileBase = "Main4"
    
    OpenFrxFile
    
    FrxTest = FrxList("004F")
    
    CloseFrxFile
End Function


Public Sub OpenFrxFile()
    Dim sFrxSpec As String
    sFrxSpec = Left$(gsInputFileSpec, Len(gsInputFileSpec) - 1&) & "x"
    If FileExists(sFrxSpec) Then
        ghFrx = FreeFile
        Open sFrxSpec For Binary As ghFrx
        '
        ' Get the GDI+ going in case we need to save some PNG files.
        CLSIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture
        CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatGIF
        CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatJPEG
        CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatPNG
        CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), EncoderQuality
        gdipStartupInput.GdiplusVersion = 1&
        Const API_NULL As Long = 0&
        If GdiplusStartup(gdipToken, gdipStartupInput, API_NULL) <> 0& Then gdipToken = 0&
    End If
End Sub

Public Sub CloseFrxFile()
    If ghFrx Then
        Close ghFrx
        ghFrx = 0&
        miIconCount = 0&
        miPngCount = 0&
        '
        ' And clean-up the GDI+.
        If gdipToken <> 0& Then
            GdiplusShutdown gdipToken
            gdipToken = 0&
        End If
    End If
End Sub

Public Function FrxString(sHexOffset As String) As String
    ' The sHexOffset can be a complete FRM line, so long as it's got the .frx file in it.
    ' Whether or not it's actually a string is NOT tested, and things are unpredictable if it's not.
    '
    If ghFrx = 0& Then Exit Function ' Return vbNullString if we don't have an FRX file.
    '
    Dim iOffsetIn As Long
    iOffsetIn = OffsetFromLine(sHexOffset) + 1& ' We add +1 because of the way VB6 "Get" works.
    If iOffsetIn <= 0& Then Exit Function
    '
    Dim iLength As Long
    Get #ghFrx, iOffsetIn, iLength
    If iLength <= 0& Then Exit Function
    '
    Dim bb() As Byte
    ReDim bb(iLength - 1&)
    Get #ghFrx, iOffsetIn + 4&, bb()
    '
    FrxString = StrConv(bb, vbUnicode)
End Function

Public Function FrxMultiLineText(sHexOffset As String) As String
    ' This one is a bit strange.  If the first byte is 0xFF, the length is in the 2nd & 3rd byte.
    ' Otherwise, that first byte is the length, with data starting in the 2nd byte.
    '
    If ghFrx = 0& Then Exit Function ' Return vbNullString if we don't have an FRX file.
    '
    Dim iOffsetIn As Long
    iOffsetIn = OffsetFromLine(sHexOffset) + 1& ' We add +1 because of the way VB6 "Get" works.
    If iOffsetIn <= 0& Then Exit Function
    '
    Dim iByte As Byte
    Dim iLength As Integer
    Get #ghFrx, iOffsetIn, iByte
    If iByte = CByte(&HFF) Then
        Get #ghFrx, iOffsetIn + 1&, iLength
        iOffsetIn = iOffsetIn + 3&
    Else
        iLength = iByte
        iOffsetIn = iOffsetIn + 1&
    End If
    If iLength <= 0& Then Exit Function
    '
    Dim bb() As Byte
    ReDim bb(iLength - 1&)
    Get #ghFrx, iOffsetIn, bb()
    '
    FrxMultiLineText = StrConv(bb, vbUnicode)
End Function

Public Function FrxIcon(sHexOffset As String) As String
    ' Returns the full spec of the icon.
    ' The icon file is saved in the folder: gsOutputFilePath\Images
    '
    If ghFrx = 0& Then Exit Function ' Return vbNullString if we don't have an FRX file.
    '
    Dim iOffsetIn As Long
    iOffsetIn = OffsetFromLine(sHexOffset) + 1& ' We add +1 because of the way VB6 "Get" works.
    If iOffsetIn <= 0& Then Exit Function
    '
    Dim iLength As Long
    Get #ghFrx, iOffsetIn, iLength
    If iLength <= 0& Then Exit Function
    '
    ' Make sure our image tags are correct.
    Dim bbTags(3&) As Byte
    Get #ghFrx, iOffsetIn + 4&, bbTags()
    If bbTags(0&) <> &H6C Or bbTags(1&) <> &H74 Or bbTags(2&) <> 0 Or bbTags(3&) <> 0 Then Exit Function
    '
    ' Get data size, and validate it.
    Dim iDataSize As Long
    Get #ghFrx, iOffsetIn + 8&, iDataSize
    If iLength - 8& <> iDataSize Then Exit Function
    '
    ' Make sure it's an icon.
    Dim iType As Long
    Get #ghFrx, iOffsetIn + 12&, iType
    If iType <> &H10000 Then Exit Function
    '
    ' Get our icon data.
    Dim bbData() As Byte
    ReDim bbData(iLength - 9&)
    Get ghFrx, iOffsetIn + 12&, bbData()
    '
    ' And save it.
    miIconCount = miIconCount + 1&
    Dim sIconSpec As String, sIconPath As String, hOut As Long
    sIconPath = gsOutputFilePath & "images\"
    If Not FolderExists(sIconPath) Then MkDir sIconPath
    sIconSpec = sIconPath & gsOutputFileBase & "_Icon" & CStr(miIconCount) & ".ico"
    If FileExists(sIconSpec) Then Kill sIconSpec
    hOut = FreeFile
    Open sIconSpec For Binary As hOut
    Put hOut, 1&, bbData()
    Close hOut
    '
    ' And return its name.  We put it back together in python code so it's OS agnostic.
    FrxIcon = Mid$(sIconSpec, InStrRev(sIconSpec, "\") + 1&)
    FrxIcon = Replace(FrxIcon, "'", "\'") ' Escape any single-quote values.
End Function

Public Function FrxCursor(sHexOffset As String) As String
    ' We'll do this one later if we need it.
End Function

Public Function FrxImage(sHexOffset As String) As String
    ' We'll only deal with BMP, GIF, JPG, WMF, or EMF herein.
    ' Anything else (such as icons), we'll just ignore.
    ' And we'll save them all as PNG files, so Python can easily handle them.
    '
    Const BI_BITFIELDS As Long = 3&
    '
    If ghFrx = 0& Then Exit Function ' Return vbNullString if we don't have an FRX file.
    '
    Dim iOffsetIn As Long
    iOffsetIn = OffsetFromLine(sHexOffset) + 1& ' We add +1 because of the way VB6 "Get" works.
    If iOffsetIn <= 0& Then Exit Function
    '
    Dim iLength As Long
    Get #ghFrx, iOffsetIn, iLength
    If iLength <= 0& Then Exit Function
    '
    ' Make sure our image tags are correct.
    Dim bbTags(3&) As Byte
    Get #ghFrx, iOffsetIn + 4&, bbTags()
    If bbTags(0&) <> &H6C Or bbTags(1&) <> &H74 Or bbTags(2&) <> 0 Or bbTags(3&) <> 0 Then Exit Function
    '
    ' Get data size, and validate it.
    Dim iDataSize As Long
    Get #ghFrx, iOffsetIn + 8&, iDataSize
    If iLength - 8& <> iDataSize Then Exit Function
    '
    ' Make sure it's one of: BMP, GIF, JPG, WMF, or EMF.  If not, get out.
    ' Only other possible types are icons or cursors.
    Dim iType As Long
    Get #ghFrx, iOffsetIn + 12&, iType
    If iType = &H10000 Or iType = &H20000 Then Exit Function
    '
    ' Get our image data, based on the type.
    Dim bbData() As Byte
    ReDim bbData(iDataSize - 1&)
    Get ghFrx, iOffsetIn + 12&, bbData()
    '
    ' Make a StdPicture out of it.
    Dim StdPic As StdPicture
    Set StdPic = ArrayToPicture(VarPtr(bbData(0&)), UBound(bbData) + 1&)
    If StdPic Is Nothing Then Exit Function
    If StdPic.Handle = 0& Then Exit Function
    '
    ' Build our output file specification.
    miPngCount = miPngCount + 1&
    Dim sPngSpec As String, sPngPath As String, hOut As Long
    sPngPath = gsOutputFilePath & "images\"
    If Not FolderExists(sPngPath) Then MkDir sPngPath
    sPngSpec = sPngPath & gsOutputFileBase & "_Image" & CStr(miPngCount) & ".png"
    If FileExists(sPngSpec) Then Kill sPngSpec
    '
    ' Save our StdPic as a PNG.
    SavePng StdPic, sPngSpec
    '
    ' And return its name.  We put it back together in python code so it's OS agnostic.
    FrxImage = Mid$(sPngSpec, InStrRev(sPngSpec, "\") + 1&)
    FrxImage = Replace(FrxImage, "'", "\'") ' Escape any single-quote values.
End Function

Public Function FrxList(sHexOffset As String) As String
    ' These are an array of ANSI strings (not Unicode) in the FRX.
    ' 4 byte header, first two bytes are array count, second two are length of longest string.
    ' Each string is prefaced with a two byte length, followed by the string, no terminator.
    ' A vbNullChar delimited string is returned.
    '
    If ghFrx = 0& Then Exit Function ' Return vbNullString if we don't have an FRX file.
    '
    Dim iOffsetIn As Long
    iOffsetIn = OffsetFromLine(sHexOffset) + 1& ' We add +1 because of the way VB6 "Get" works.
    If iOffsetIn <= 0& Then Exit Function
    '
    ' Get the count of array items.
    Dim iCount As Long
    Get #ghFrx, iOffsetIn, iCount
    iCount = iCount And &H7FFF& ' Mask off max value, and make sure it's positive.
    If iCount = 0& Then Exit Function
    '
    ' Get the data.
    Dim iOff As Long
    iOff = iOffsetIn + 4&
    Dim iLen As Integer
    Dim iPtr As Long
    Dim sa() As String
    Dim bb() As Byte
    ReDim sa(iCount - 1&)
    For iPtr = 0& To iCount - 1&
        Get #ghFrx, iOff, iLen
        ReDim bb(iLen - 1&)
        iOff = iOff + 2&
        Get #ghFrx, iOff, bb()
        sa(iPtr) = StrConv(bb, vbUnicode)
        iOff = iOff + CLng(iLen)
    Next
    '
    ' Return our data.
    FrxList = Join(sa, vbNullChar)
End Function

' *********************************************
' Some private stuff.
' *********************************************

Private Function OffsetFromLine(ByVal sHexOffset As String) As Long
    ' Returns -1 if there's any problem.
    '
    OffsetFromLine = -1&
    '
    Dim i As Long, iOffsetIn As Long
    i = InStrRev(sHexOffset, ":")
    If i Then sHexOffset = Mid$(sHexOffset, i + 1&)
    '
    Dim iErr As Long
    On Error Resume Next
        i = CLng("&h" & sHexOffset)
        iErr = Err.Number
    On Error GoTo 0
    If iErr Then Exit Function
    '
    OffsetFromLine = i
End Function

Private Sub SavePng(StdPicture As StdPicture, FileName As String)
    Dim gdipBitmap As Long
    Dim Stream As IUnknown
    Dim pParams As Long
    Const STGM_WRITE As Long = &H1&
    Const STGM_SHARE_EXCLUSIVE As Long = &H10&
    Const STGM_CREATE As Long = &H1000&
    '
    If StdPicture Is Nothing Then Err.Raise 5&
    '
    Call GdipCreateBitmapFromHBITMAP(StdPicture.Handle, StdPicture.hpal, gdipBitmap)
    Call SHCreateStreamOnFile(StrPtr(FileName), STGM_CREATE Or STGM_WRITE Or STGM_SHARE_EXCLUSIVE, Stream)
    Call GdipSaveImageToStream(gdipBitmap, Stream, ImageFormatPNG, pParams)
    GdipDisposeImage gdipBitmap
End Sub

Private Function ArrayToPicture(ArrayVarPtr As Long, iSize As Long) As IPicture
    ' function creates a stdPicture from the passed array
    '
    Dim aGUID(3&) As Long
    Dim IIStream As IUnknown, hMem As Long, lpMem As Long
    '
    On Error GoTo ExitRoutine
    hMem = GlobalAlloc(&H2&, iSize)
    If hMem <> 0& Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0& Then
            CopyMemory ByVal lpMem, ByVal ArrayVarPtr, iSize
            Call GlobalUnlock(hMem)
            Call CreateStreamOnHGlobal(hMem, 1&, IIStream)
        End If
    End If
    '
    If Not IIStream Is Nothing Then
        aGUID(0&) = &H7BF80980    ' GUID for stdPicture
        aGUID(1&) = &H101ABF32
        aGUID(2&) = &HAA00BB8B
        aGUID(3&) = &HAB0C3000
        Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0&), ArrayToPicture)
    End If
    '
ExitRoutine:
End Function


