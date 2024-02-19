Attribute VB_Name = "mod_Gen_Misc"
Option Explicit
'
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'
Private Declare Function CryptBinaryToString Lib "Crypt32" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Sub SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef psaInOut As Long)
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (a() As Any) As Long
'



Public Sub MakeZeroToNegOneArray(pArray As Long, Optional cDims As Long = 1&)
    ' Works with all array types (numbers, UDTs, objects (early or late), fixed-length strings), except:
    '       BSTR (String) arrays.  Use s() = Split(vbNullString) instead.
    '       Probably shouldn't be used with arrays IN varants, but a Varant array should be just fine.
    '       Non-dynamic arrays.  This is obvious, but still important to remember.
    '
    ' WARNING:  Before calling this, make SURE the array is ERASED.
    ' WARNING:  You can't use REDIM PRESERVE on these things.  Just use REDIM TheArray(Low, High)
    '
    ' If cDims > 1 then all the dimensions are made 0 to -1.
    '
    ' Example:  Erase SomeArray
    '           MakeZeroToNegOneArray ArrPtr(SomeArray)
    '
    SafeArrayAllocDescriptor cDims, ByVal pArray
End Sub

Public Function Base64Decode(sBase64Buf As String) As Byte()
    Const CRYPT_STRING_BASE64 As Long = 1&
    ' Get output buffer length.
    Dim lLen As Long
    Dim dwActualUsed As Long
    Call CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen, 0&, dwActualUsed)
    ' Convert Base64 to binary.
    Dim bb() As Byte: ReDim bb(lLen - 1&)
    Call CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, VarPtr(bb(0&)), lLen, 0&, dwActualUsed)
    ' Return results.
    Base64Decode = bb
End Function

Public Function Base64Encode(bbData() As Byte) As String
    Const CRYPT_STRING_BASE64 As Long = 1&
    ' Determine Base64 output String length required.
    Dim lLen As Long
    Call CryptBinaryToString(bbData(0&), UBound(bbData) + 1&, CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen)
    ' Convert binary to Base64.
    Base64Encode = String$(lLen - 1&, vbNullChar)
    Call CryptBinaryToString(bbData(0&), UBound(bbData) + 1&, CRYPT_STRING_BASE64, StrPtr(Base64Encode), lLen)
End Function

Public Function SelectedOptIndex(opt As Object) As Integer
    Dim o As Object
    For Each o In opt
        If o.Value Then
            SelectedOptIndex = o.Index
            Exit Function
        End If
    Next
    SelectedOptIndex = -1 ' This only happens if none are selected.
End Function

Public Sub nop()
    ' Just a dummy procedure for formatting VB code lines with the : statement separator.
End Sub

Public Function StrArrayHasMatch(sHay() As String, sNeedle As String) As Boolean
    Dim i As Long
    For i = LBound(sHay) To UBound(sHay)
        If sHay(i) = sNeedle Then
            StrArrayHasMatch = True
            Exit Function
        End If
    Next
    ' Just fall out if no match.
End Function

Public Function StrArrayHasRightMatch(sHay() As String, sNeedle As String) As Boolean
    Dim i As Long
    For i = LBound(sHay) To UBound(sHay)
        If Right$(sHay(i), Len(sNeedle)) = sNeedle Then
            StrArrayHasRightMatch = True
            Exit Function
        End If
    Next
    ' Just fall out if no match.
End Function

Public Function LeftMatch(sHay As String, sNeedle As String) As Boolean
    ' Case sensitive.
    LeftMatch = Left$(sHay, Len(sNeedle)) = sNeedle
End Function

Public Function UniqueFileSpec(sFileSpec As String) As String
    ' This add (#), finding a unique file.
    ' The full filespec should be provided.
    ' It MUST have an extension.
    '
    Dim sExtWithDot As String
    Dim sBase As String
    Dim i As Long
    Dim inc As Long
    '
    If Not FileExists(sFileSpec) Then
        UniqueFileSpec = sFileSpec
        Exit Function
    End If
    '
    i = InStrRev(sFileSpec, ".")
    sExtWithDot = Mid$(sFileSpec, i)
    sBase = Left$(sFileSpec, i - 1&)
    inc = 2& ' May be overwritten below.
    If InStr(sBase, " ") Then
        Dim sParen As String
        sParen = Mid$(sBase, InStrRev(sBase, " ") + 1&)
        If sParen Like "(###)" Or sParen Like "(##)" Or sParen Like "(#)" Then
            sBase = Left$(sBase, Len(sBase) - Len(sParen) - 1&)
            inc = CLng(Mid$(sParen, 2&, Len(sParen) - 2&)) + 1&
        End If
    End If
    Do ' No exit until we find a unique file name.
        UniqueFileSpec = sBase & " (" & Format$(inc) & ")" & sExtWithDot
        If Not FileExists(UniqueFileSpec) Then Exit Function
        inc = inc + 1&
    Loop
End Function

Public Function FileExists(sFileSpec As String) As Boolean
    On Error GoTo ExistsError
    ' If no error then something existed.
    FileExists = (GetAttr(sFileSpec) And vbDirectory) = 0&
    Exit Function
ExistsError:
    FileExists = False
End Function

Public Function FolderExists(ByVal sFolder As String) As Boolean
    If Right$(sFolder, 1&) = "\" Then sFolder = Left$(sFolder, Len(sFolder) - 1&)
    On Error GoTo ExistsError
    ' If no error then something existed.
    FolderExists = ((GetAttr(sFolder) And vbDirectory) = vbDirectory)
    Exit Function
ExistsError:
    FolderExists = False
End Function

Public Function RgbHex(ByVal iColor As Long) As String
    ' Returns string formatted as: #rrggbb (not like VB6, but like python).
    ' Includes the # sign upon return.
    ' Converts Windows system colors, if needed.
    Dim r As Long, g As Long, b As Long
    If iColor < 0& Then iColor = GetSysColor(iColor And &HFF)
    r = iColor And &HFF
    g = (iColor \ &H100&) And &HFF
    b = (iColor \ &H10000) And &HFF
    RgbHex = "#" & Right$("0" & Hex$(r), 2&) & Right$("0" & Hex$(g), 2&) & Right$("0" & Hex$(b), 2&)
End Function


