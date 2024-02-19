Attribute VB_Name = "mod_Frm2Py_Misc_Support"
Option Explicit
'


Public Sub PrintWidgetFontLines(uFont As FontType)
    Dim sFont As String
    '
    Print #ghPy, "            font = QFont('"; uFont.Name; "', "; CStr(CLng(uFont.Size)); ")"
    If uFont.Bold Then nop:           sFont = sFont & "font.setBold(True); "
    If uFont.Italic Then nop:         sFont = sFont & "font.setItalic(True); "
    If uFont.Underline Then nop:      sFont = sFont & "font.setUnderline(True); "
    If uFont.Strikethrough Then nop:  sFont = sFont & "font.setStrikeOut(True); "
  If Len(sFont) Then
    sFont = Left$(sFont, Len(sFont) - 2&) ' Clean it up.
    Print #ghPy, "            "; sFont
  End If
End Sub




Public Function TrueFalse(b As Boolean) As String
    If b Then
        TrueFalse = "True"
    Else
        TrueFalse = "False"
    End If
End Function

Public Function PythonListFromFrxList(sFrxList As String) As String
    '
    ' Escape any single-quotes in the list.  Double-quotes will actually be ok.
    sFrxList = Replace(sFrxList, "'", "\'")
    '
    ' Parse into an array.
    Dim sa() As String
    sa = Split(sFrxList, vbNullChar)
    '
    ' Build Python list.
    Select Case UBound(sa)
    Case -1&
        PythonListFromFrxList = "[]"
    Case 0&
        PythonListFromFrxList = "['" & sa(0&) & "']"
    Case 1&
        PythonListFromFrxList = "['" & sa(0&) & "', '" & sa(1&) & "']"
    Case Else
        Dim i As Long
        PythonListFromFrxList = "['" & sa(0&) & "', " & vbCrLf
        For i = 1& To UBound(sa) - 1&
            PythonListFromFrxList = PythonListFromFrxList & Space$(21&) & "'" & sa(i) & "', " & vbCrLf
        Next
        PythonListFromFrxList = PythonListFromFrxList & Space$(21&) & "'" & sa(i) & "']"
    End Select
End Function

Public Function GetStringValue(sLine As String) As String
    ' Gets the value after the equal sign, trimmed.
    ' Removes any leading and following " characters.
    ' Escapes any ' characters (python style).
    ' This assumes the ' character will be used to delineate the string in python.
    ' This procedure also gets the string from the FRX file, if necessary.
    GetStringValue = AfterEqual(sLine)
    If InStr(GetStringValue, ".frx"":") Then ' Get from FRX file.
        If ghFrx Then
            If Left$(GetStringValue, 1&) = "$" Then
                GetStringValue = FrxString(sLine)
            Else
                GetStringValue = FrxMultiLineText(sLine)
            End If
        Else
            GetStringValue = vbNullString
        End If
    Else
        ' Trim the quote marks.
        If Left$(GetStringValue, 1&) = """" Then GetStringValue = Mid$(GetStringValue, 2&)
        If Right$(GetStringValue, 1&) = """" Then GetStringValue = Left$(GetStringValue, Len(GetStringValue) - 1&)
    End If
    ' Escape any single-quote values and && values.
    GetStringValue = Replace(GetStringValue, "'", "\'")
    GetStringValue = Replace(GetStringValue, "&&", "&")
End Function

Public Function FixMultiString(s As String, iIndent As Long) As String
    ' NOTICE:  This INCLUDES the string's delineating ' symbols.
    '
    Dim i As Long
    i = InStr(s, vbCrLf)
    Select Case i
    Case 0&
        FixMultiString = "'" & s & "'"
    Case Len(s) - 1&    ' Only one vbCrLf, and it's at the end.
        FixMultiString = "'" & Replace$(s, vbCrLf, "\n") & "'"
    Case Else
        If Right$(s, 2&) = vbCrLf Then s = Left$(s, Len(s) - 2&) ' We'll add this back during processing.
        Dim sa() As String
        sa = Split(s, vbCrLf)
        FixMultiString = "'" & sa(0) & "\n' \"
        For i = 1& To UBound(sa) - 1&
            FixMultiString = FixMultiString & vbCrLf & Space$(iIndent) & "'" & sa(i) & "\n' \"
        Next
        FixMultiString = FixMultiString & vbCrLf & Space$(iIndent) & "'" & sa(i) & "\n'"
    End Select
End Function

Public Function AfterEqual(sLine As String) As String
    AfterEqual = Trim$(Mid$(sLine, InStr(sLine, "=") + 1&))
End Function

Public Function CLngEx(s As String) As Long
    ' Just handles a trailing & symbol on the input string (as seen in FRM files).
    If Right$(s, 1&) = "&" Then
        CLngEx = CLng(Left$(s, Len(s) - 1&))
    Else
        CLngEx = CLng(s)
    End If
End Function


