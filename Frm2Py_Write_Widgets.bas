Attribute VB_Name = "mod_Frm2Py_Write___Widgets"
Option Explicit
'

Public Sub DoAllWidgetsClasses()
    ' All the classes for the widgets are created here, within the form's class.
    ' As part of this, within each widget's class, all its methods, property, & internal events are created.
    ' In many cases, the internal events will check to see if an external event exists, and call it if it does.
    '
    ' Comment.
    Print #ghPy, vbNullString
    Print #ghPy, "    # ***********************************************************"
    Print #ghPy, "    # Widget classes with __init__ procs, methods, & properties."
    Print #ghPy, "    # ***********************************************************"
    '
    Dim pCtl As Long
    '
    ' Heavyweight ones first.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            Select Case .ClassName
                '
            Case "CommandButton":       DoCommandButtonClass guCtls(pCtl)
            Case "CheckBox":            DoCheckBoxClass guCtls(pCtl)
            Case "OptionButton":        DoOptionButtonClass guCtls(pCtl)
            Case "TextBox"
              If .MultiLine Then
                                        DoTextBoxMultiLineClass guCtls(pCtl)
              Else
                                        DoTextBoxSingleLineClass guCtls(pCtl)
              End If
            Case "Frame":               DoFrameClass guCtls(pCtl)
            Case "PictureBox":          DoPictureBoxClass guCtls(pCtl)
            Case "ListBox":             DoListBoxClass guCtls(pCtl)
            Case "ComboBox":            DoComboBoxClass guCtls(pCtl)
            End Select
        End With
    Next
    '
    ' And now the lightweight ones.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            Select Case .ClassName
            Case "Label":               DoLabelClass guCtls(pCtl)
            Case "Image":               DoImageClass guCtls(pCtl)
            Case "Line":                DoLineClass guCtls(pCtl)
            Case "Shape":               DoShapeClass guCtls(pCtl)
            End Select
        End With
    Next
End Sub

Public Sub InstantiateAllTheWidgets()
    Print #ghPy, vbNullString
    Print #ghPy, "        # Instantiate all the widgets.  Last one will have lowest z-order (first highest), as in VB6."
    Print #ghPy, vbNullString
    Dim sContainer As String
    Dim pCtl As Long
    '
    For pCtl = 0& To UBound(guCtls) ' We can't do it backwards because the containers must be instantiated before their contained widgets.
        With guCtls(pCtl)
            Select Case .ClassName
            '
            Case "CommandButton", "CheckBox", "OptionButton", "TextBox", "Frame", "PictureBox", "ListBox", "ComboBox"
                '
                GoSub SetContainerName ' Figure out container.
                Print #ghPy, "        self."; .Name; " = self.cls"; .Name; "("; sContainer; ", self); "; ' Write instantiation line.
                ' We put this next one on the same line as instantiation.
                Print #ghPy, "self."; .Name; ".lower()" ' And to bottom, because VB6 puts first on top, and Python puts last on top.
                .DoneInPython = True
            End Select
        End With
    Next
    '
    ' And now the lightweight ones.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            Select Case .ClassName
            '
            Case "Label", "Image"
                '
                GoSub SetContainerName ' Figure out container.
                Print #ghPy, "        self."; .Name; " = self.cls"; .Name; "("; sContainer; ", self); "; ' Write instantiation line.
                ' We put this next one on the same line as instantiation.
                Print #ghPy, "self."; .Name; ".lower()" ' And to bottom, because VB6 puts first on top, and Python puts last on top.
                .DoneInPython = True
            Case "Line", "Shape"
                '
                GoSub SetContainerName ' Figure out container.
                Print #ghPy, "        self."; .Name; " = self.cls"; .Name; "("; sContainer; ", self)" ' Write instantiation line.
                ' Line and Shape don't inherit anything, so there's no .lower() method to worry about.
                .DoneInPython = True
            End Select
        End With
    Next
    Exit Sub
    '
SetContainerName:
    With guCtls(pCtl)
        If .ContainerName = guForm.Name Then
            sContainer = "self"
        Else
            sContainer = "self." & .ContainerName
        End If
    End With
    Return
End Sub

Public Sub GetRidOfControlsNotProcessed()
    Dim pCtl As Long
    Dim iCtlsCount As Long
    Dim i As Long
    iCtlsCount = UBound(guCtls) + 1&
    Do While pCtl < iCtlsCount
        If guCtls(pCtl).DoneInPython = False Then
            For i = pCtl + 1& To iCtlsCount - 1&
                guCtls(i - 1&) = guCtls(i)
            Next
            iCtlsCount = iCtlsCount - 1&
        Else
            pCtl = pCtl + 1&
        End If
    Loop
    If iCtlsCount = 0& Then
        MakeZeroToNegOneArray ArrPtr(guCtls)
    Else
        ReDim Preserve guCtls(iCtlsCount - 1&)
    End If
End Sub

Public Sub SetWidgetTabOrders()
    '
    ' We will make a copy because not all widgets can participate in the tab order.
    If UBound(guCtls) = -1& Then Exit Sub ' No widgets to process.
    '
    ' First count widgets that CAN participate in tab order, copying as we go.
    ' We also count how many >=0 TabIndex values there are, and also find TabIndex_Max.
    Dim iTabsCount As Long, iPosCount As Long, iTabMax As Long
    Dim uTabs() As TabsType
    ReDim uTabs(UBound(guCtls))
    Dim pCtl As Long
    For pCtl = 0& To UBound(guCtls)
        Select Case guCtls(pCtl).ClassName
        Case "Label", "Frame", "Line", "Shape", "Image"       ' Can NOT participate.
        Case Else
            uTabs(iTabsCount).ClassName = guCtls(pCtl).ClassName
            uTabs(iTabsCount).Name = guCtls(pCtl).Name
            uTabs(iTabsCount).TabIndex = guCtls(pCtl).TabIndex
            uTabs(iTabsCount).TabStop = guCtls(pCtl).TabStop
            iTabsCount = iTabsCount + 1&
            If guCtls(pCtl).TabIndex >= 0& Then iPosCount = iPosCount + 1&
            If guCtls(pCtl).TabIndex > iTabMax Then iTabMax = guCtls(pCtl).TabIndex
        End Select
    Next
    If iTabsCount < 2& Then Exit Sub ' There's nothing to do with less than 2 widgets needing tab orders.
    ReDim Preserve uTabs(iTabsCount - 1&)
    '
    ' Push all the -1 TabIndex values out to the end.
    Dim pTab As Long
    For pTab = 0& To UBound(uTabs)
        If uTabs(pTab).TabIndex = -1& Then
            iTabMax = iTabMax + 1&
            uTabs(pTab).TabIndex = iTabMax
        End If
    Next
    '
    ' Sort them on TabIndex.  Holes (non-contiguous TabOrder) won't matter.
    Dim uTabSwap As TabsType
    Dim i As Long, j As Long
    For i = 0& To UBound(uTabs) - 1&
        For j = i + 1& To UBound(uTabs)
            If uTabs(i).TabIndex > uTabs(j).TabIndex Then
                uTabSwap = uTabs(i)
                uTabs(i) = uTabs(j)
                uTabs(j) = uTabSwap
            End If
        Next
    Next
    '
    ' We can now insert our setTabOrder statements.
    Print #ghPy, vbNullString
    Print #ghPy, "        # Tab orders for all the widgets, per VB6's TabIndex order."
    Print #ghPy, "        # Completely independent of z-order, which is also true in VB6."
    Print #ghPy, vbNullString
    For pTab = 0& To UBound(uTabs) - 1&
        Print #ghPy, "        self.setTabOrder(self."; uTabs(pTab).Name; ", self."; uTabs(pTab + 1&).Name; ")"
    Next
End Sub

Public Sub BuildWidgetsDictionary()
    Print #ghPy, vbNullString
    Print #ghPy, "        # Widgets collection (dictionary) similar to VB6's 'Controls' collection."
    Print #ghPy, vbNullString
    Select Case UBound(guCtls)
    Case -1&
        Print #ghPy, "        self.Widgets = []"
    Case 0&
        Print #ghPy, "        self.Widgets = [self."; guCtls(0&).Name; "]"
    Case 1&
        Print #ghPy, "        self.Widgets = [self."; guCtls(0&).Name; ", self."; guCtls(1&).Name; "]"
    Case Else
        Print #ghPy, "        self.Widgets = [self."; guCtls(0&).Name; ", "
        Dim pCtl As Long
      For pCtl = 1& To UBound(guCtls) - 1&
        Print #ghPy, "                        self."; guCtls(pCtl).Name; ", "
      Next
        Print #ghPy, "                        self."; guCtls(UBound(guCtls)).Name; "]"
    End Select
End Sub


' *****************************************************************************************
' *****************************************************************************************
' The following are called by ...Write_Main... after the _Events file is possibly created.
' *****************************************************************************************
' *****************************************************************************************

Public Sub DoExternalWidgetEventProcedures()
    ' We will loop through our control list in here.
    '
    ' We create a collection, so we don't create any duplicate event procedures.
    ' This covers us in the case of control (widget) arrays.
    Dim collProcsDone As New Collection
    '
    Print #ghPy, vbNullString
    Print #ghPy, "# ************************************************"
    Print #ghPy, "# Widget event procedures for coding."
    Print #ghPy, "# Delete them if you don't need them."
    Print #ghPy, "# ************************************************"
    '
    Dim pCtl As Long
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            ' Just include ALL possible events (for any/all widgets).
            ' Which widgets get which calls is tested inside the AddExternalWidgetEventProc procedure.
            AddExternalWidgetEventProc collProcsDone, .OrigName & "_Change", guCtls(pCtl)
            AddExternalWidgetEventProc collProcsDone, .OrigName & "_Click", guCtls(pCtl)
            AddExternalWidgetEventProc collProcsDone, .OrigName & "_DblClick", guCtls(pCtl)
        End With
    Next
End Sub

Private Sub AddExternalWidgetEventProc(collProcsDone As Collection, sProcName As String, uCtrl As CtrlType)
    ' Just support for DoExternalWidgetEventProcedures
    '
    ' Make sure we haven't already done it.  If we have, skip it.
    ' This allows us to handle control arrays.
    Dim iErr As Long
    On Error Resume Next
        collProcsDone.Add sProcName, sProcName
        iErr = Err.Number
    On Error GoTo 0
    If iErr Then Exit Sub ' We've already done this one.
    '
    ' Now, if it's a control array, add to argument list.
    Dim sIndex As String
    If uCtrl.IsIndexed Then sIndex = "Index, "
    '
    ' Used so we can easily figure out what type of event we're dealing with.
    Dim sSuffix As String
    sSuffix = Mid$(sProcName, InStrRev(sProcName, "_")) ' We keep the underscore.
    '
    ' And, the procedures.
    Select Case uCtrl.ClassName
    '
    ' Heavyweight ones.
    '
    Case "CommandButton"
        Select Case sSuffix
        Case "_Click"
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, event):"
            Print #ghPy, "    print('"; sProcName; "', self.objectName(), self.Container.objectName(), self.Form.objectName())"
        End Select
    Case "CheckBox"
        Select Case sSuffix
        Case "_Click"
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, state):"
            Print #ghPy, "    print('"; sProcName; "', self.objectName(), self.Container.objectName(), self.Form.objectName(), 'State:', state)"
        End Select
    Case "OptionButton"
        Select Case sSuffix
        Case "_Click"
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, state):"
            Print #ghPy, "    print('"; sProcName; "', self.objectName(), self.Container.objectName(), self.Form.objectName(), 'State:', state)"
        End Select
    Case "TextBox"
        Select Case sSuffix
        Case "_Change"
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, text):"
            Print #ghPy, "    print('"; sProcName; "', self.objectName(), self.Container.objectName(), self.Form.objectName(), 'Text:', text)"
        End Select
    Case "Frame"            ' None, at this time.
    Case "PictureBox"       ' None, at this time.
    Case "ListBox"
        Select Case sSuffix
        Case "_Click", "_DblClick" ' Other than the procedure name, these are identical, at least the inserted stub.
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, item):"
            Print #ghPy, "    print('"; sProcName; "', 'clicked:', item.text(), end=' selected: ')"
            Print #ghPy, "    selected_items = self.selectedItems()"
            Print #ghPy, "    for item in selected_items: print(item.text(), end=' ')"
            Print #ghPy, "    print('')"
        End Select
    Case "ComboBox"
        Select Case sSuffix
        Case "_Change"
            Print #ghPy, vbNullString
            Print #ghPy, "def "; sProcName; "("; sIndex; "self, text):"
            Print #ghPy, "    print('"; sProcName; "', self.objectName(), self.Container.objectName(), self.Form.objectName(), 'Text:', text)"
        End Select
    '
    ' Lightweight ones.
    '
    Case "Label"            ' None, at this time.
    Case "Image"            ' None, at this time.
    Case "Line"             ' None, at this time.
    Case "Shape"            ' None, at this time.
    End Select
End Sub

