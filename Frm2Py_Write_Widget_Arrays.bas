Attribute VB_Name = "mod_Frm2Py_Write___Widgets_Arrays"
Option Explicit
'

Public Sub DoWidget_Arrays()
    ' This one is specifically to deal with VB6 control arrays.
    ' A separate DEF is built for each array to deal with them.
    ' This DEF, with an index, can be called much like VB6 does it.
    ' Or, alternatively, the individual widgets can be called.
    '
    ' First, build a collection of all the indexed control names.
    ' We do this down here because this actually gets done BEFORE DoWidget_Arrays.
    '
    Dim collIndexed As New Collection
    Dim pCtl As Long, sCtlName As String
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .IsIndexed Then
                On Error Resume Next
                    ' There will almost certainly be dupes, but we just want one name per array.
                    collIndexed.Add .OrigName, .OrigName
                On Error GoTo 0
            End If
        End With
    Next
    '
    ' Make sure we've got some.  If not, get out.
    If collIndexed.Count = 0& Then Exit Sub
    '
    ' Build a DEF for each array.
                    Print #ghPy, vbNullString
                    Print #ghPy, "    # Below are the functions to deal with control (widget) arrays."
                    Print #ghPy, "    # These allow us to address the widget arrays very similar to VB6 control arrays."
                    Print #ghPy, "    # Alternatively, we can address the individual instantiated widget objects, if we like."
    Dim v As Variant, sOrigName As String
    For Each v In collIndexed
        sOrigName = v
                    Print #ghPy, vbNullString
                    Print #ghPy, "    def "; sOrigName; "(self, Index):"
        For pCtl = 0& To UBound(guCtls)
            With guCtls(pCtl)
                If .OrigName = sOrigName Then
                    Print #ghPy, "        if Index == "; CStr(.Index); ": return self."; .Name
                End If
            End With
        Next
                    Print #ghPy, "        raise ValueError('Invalid index specified.')"
    Next
End Sub
