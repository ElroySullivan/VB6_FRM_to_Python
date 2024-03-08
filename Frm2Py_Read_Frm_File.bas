Attribute VB_Name = "mod_Frm2Py_Read_Frm_File"
Option Explicit
'

Public Function GotFrmInputFileSpec() As Boolean
    '
    ' Get our file specification.
    Dim sPath As String
    sPath = GetSetting("VB6_FRM_to_PY", "Settings", "LastPath", App.Path & "\")
    ShowOpenFileDialog 0&, "Frm File (*.frm)" & vbNullChar & "*.frm" & vbNullChar, sPath, OFN_FILEMUSTEXIST, "VB6 FRM File to Convert"
    If Not FileDialogSuccessful Then Exit Function
    '
    ' Save our file specifications.
    gsInputFileSpec = FileDialogSpec
    gsInputFileName = FileDialogName
    gsInputFilePath = FileDialogFolder
    gsInputFileBase = Left$(gsInputFileName, Len(gsInputFileName) - 4&)
    SaveSetting "VB6_FRM_to_PY", "Settings", "LastPath", gsInputFilePath
    '
    ' Good to go.
    GotFrmInputFileSpec = True
End Function
    
Public Function LoadedFrmFileAndCleanedUp() As Boolean
    ' This populates the gsUiLines() and gsCodeLines() arrays, and cleans them up.
    '
    Dim i As Long, j As Long
    Dim sLine As String
    '
    ' Get our file open.
    Dim hFrm As Long
    hFrm = FreeFile
    Open gsInputFileSpec For Input As hFrm
    '
    ' Count UI lines.
    Dim iUiLineCount As Long
    Do
        If EOF(hFrm) Then
            MsgBox "ERROR: Not a well formatted FRM file.  Found EOF before finding ""Attribute"" line.", vbCritical
            Close hFrm
            Exit Function
        End If
        Line Input #hFrm, sLine
        If LeftMatch(sLine, "Attribute") Then Exit Do
        iUiLineCount = iUiLineCount + 1&
    Loop
    '
    ' Get UI section into memory.
    ReDim gsUiLines(iUiLineCount - 1&)
    Seek #hFrm, 1&
    For i = 0& To UBound(gsUiLines)
        Line Input #hFrm, gsUiLines(i)
    Next
    '
    ' Count code lines, for gathering used events.
    ' First spin to beginning of code.
    Dim iCodeLineCount As Long
    Seek #hFrm, 1&
    Do
        Line Input #hFrm, sLine
        If LeftMatch(sLine, "Attribute") Then Exit Do
    Loop
    Do
        If EOF(hFrm) Then
            MsgBox "ERROR: Not a well formatted FRM file.  Found EOF too soon in code section.", vbCritical
            Close hFrm
            Exit Function
        End If
        Line Input #hFrm, sLine
        If LeftMatch(sLine, "Attribute") = False Then Exit Do
    Loop
    ' Save this position.
    Dim iCodeSeek As Long
    iCodeSeek = Seek(hFrm)
    ' Now count lines to eof.
    Do While Not EOF(hFrm)
        Line Input #hFrm, sLine
        iCodeLineCount = iCodeLineCount + 1&
    Loop
    '
    ' Make sure we've got some.
    If iCodeLineCount = 0& Then
        gsCodeLines = Split(vbNullString)
        GoTo NoCodeLines
    End If
    '
    ' Make space for code lines.
    ReDim gsCodeLines(iCodeLineCount - 1&)
    '
    ' Now we can read them.
    Seek #hFrm, iCodeSeek
    For i = 0& To UBound(gsCodeLines)
        Line Input #hFrm, gsCodeLines(i)
    Next
    '
    ' We're now done with the FRM file.
    Close hFrm
    '
    ' We'll clean up our code lines first, just to have it done.
    ' Trim them all, as we can't depend on any indentation like we can with the UI section.
    For i = 0& To UBound(gsCodeLines)
        gsCodeLines(i) = Trim$(gsCodeLines(i))
    Next
    ' Toss comment only lines.  We will be looking for things, and this will speed that up.
    For i = 0& To UBound(gsCodeLines)
        If Left$(gsCodeLines(i), 1&) = "'" Then gsCodeLines(i) = vbNullString
    Next
    ' Toss "Private", "Friend", "Public"
    For i = 0& To UBound(gsCodeLines)
        If LeftMatch(gsCodeLines(i), "Private ") Then gsCodeLines(i) = Mid$(gsCodeLines(i), 9&)
        If LeftMatch(gsCodeLines(i), "Friend ") Then gsCodeLines(i) = Mid$(gsCodeLines(i), 8&)
        If LeftMatch(gsCodeLines(i), "Public ") Then gsCodeLines(i) = Mid$(gsCodeLines(i), 8&)
    Next
    ' Toss all lines that don't start with "Sub ", as they can't be events,
    ' and delete "Sub " from lines that do have it.
    For i = 0& To UBound(gsCodeLines)
        If LeftMatch(gsCodeLines(i), "Sub ") Then
            gsCodeLines(i) = Mid$(gsCodeLines(i), 5&)
        Else
            gsCodeLines(i) = vbNullString
        End If
    Next
    ' Toss all lines that don't have an underscore, as they can't be events.
    For i = 0& To UBound(gsCodeLines)
        If InStr(gsCodeLines(i), "_") = 0& Then gsCodeLines(i) = vbNullString
    Next
    '
    ' Now let's pack and redim our code lines array.
    i = 0&
    Do While i < iCodeLineCount
        If Len(gsCodeLines(i)) = 0& Then
            For j = i + 1& To iCodeLineCount - 1&
                gsCodeLines(j - 1&) = gsCodeLines(j)
            Next
            iCodeLineCount = iCodeLineCount - 1&
        Else
            i = i + 1&
        End If
    Loop
    '
    ' Make sure we've still got some.
    If iCodeLineCount = 0& Then
        gsCodeLines = Split(vbNullString)
        GoTo NoCodeLines
    End If
    '
    ' Tighten up space used.
    ReDim Preserve gsCodeLines(iCodeLineCount - 1&)
    '
    ' And delete the arguments.
    For i = 0& To UBound(gsCodeLines)
        gsCodeLines(i) = Left$(gsCodeLines(i), InStr(gsCodeLines(i), "(") - 1&)
    Next
    ' At this point, we may have a few extra lines,
    ' but it will be a pretty tight list of event procedures.
    ' To check:
    '       For i = 0 To 11: Debug.Print gsCodeLines(i): Next
    '
NoCodeLines:
    '
    ' And now to process the UI lines.
    '
    ' We must remove any comments off the ends of the UI lines.
    ' I'm going to assume there are no cases where quoted string lines have comments.
    ' If there are, this will need to be smartened up.
    ' Also, it's assumed that a single quote (') is ONLY used for comments.
    For i = 0& To UBound(gsUiLines)
        If InStr(gsUiLines(i), """") = 0& Then
            j = InStr(gsUiLines(i), "'")
            If j Then gsUiLines(i) = Left$(gsUiLines(i), j - 1&)
        End If
    Next
    '
    ' Best to RTrim our lines.  The IDE sometimes put spaces out there.
    ' We leave (and use) the left indenting though (which is three spaces per indent).
    For i = 0& To UBound(gsUiLines)
        gsUiLines(i) = RTrim$(gsUiLines(i))
    Next
    '
    ' Certain control libraries, we're going to treat like they're regular VB controls.
    ' We'll just join the array to make things easier.
    ' DON'T foul up the 3-space indentation, as it's used later to figure out nesting.
    Dim sHay As String
    sHay = Join(gsUiLines, vbCrLf)
    '
    ' First do Krool's controls.  Who knows how many versions he'll eventually have.
    For j = 14& To 25& ' Versions 14 thru 15 covered.
        Dim sNeedle As String
        sNeedle = " BEGIN VBCCR" & CStr(j) & "."
        If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, "BEGIN VB.")
    Next
    ' Get the "W" off Krool's class names.
    sNeedle = ".CommandButtonW ":   If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".CommandButton ")
    sNeedle = ".CheckBoxW ":        If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".CheckBox ")
    sNeedle = ".ComboBoxW ":        If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".ComboBox ")
    sNeedle = ".FrameW ":           If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".Frame ")
    sNeedle = ".LabelW ":           If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".Label ")
    sNeedle = ".ListBoxW ":         If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".ListBox ")
    sNeedle = ".OptionButtonW ":    If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".OptionButton ")
    sNeedle = ".TextBoxW ":         If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, ".TextBox ")
    '
    ' The MSComctlLib and ComctlLib libraries.  We'll eventually pick on some of these controls.
    ' Regarding Krool, I believe his names are already the same, so we should have covered that above.
    sNeedle = " BEGIN MSComctlLib.":        If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, " BEGIN VB.")
    sNeedle = " BEGIN ComctlLib.":          If InStr(sHay, sNeedle) Then sHay = Replace(sHay, sNeedle, " BEGIN VB.")
    '
    ' Put back into array.
    gsUiLines = Split(sHay, vbCrLf)
    '
    ' Good to go.
    LoadedFrmFileAndCleanedUp = True
End Function
    
Public Function PopulatedFormUdt() As Boolean
    With guForm
        '
        ' There are a few values that may not be found, but have defaults.
        ' See far indentation ----------------->
                                                                                .BorderStyle = vbSizable            ' If it sizable, the VB6 IDE doesn't put the property in the frm file.
                                                                                .BackColor = &H8000000F             ' It may not appear in FRM file, and this is the default.
                                                                                .Enabled = True
                                                                                .ForeColor = &H80000012
                                                                                .Visible = True                     ' It may not appear in FRM file, and default is True.
        '
        ' Parse form into its UDT.
        Dim bInFont As Boolean
        Dim i As Long
        For i = 0& To UBound(gsUiLines)
            Select Case True
            Case LeftMatch(gsUiLines(i), "Begin VB.Form"):                      .Name = Mid$(gsUiLines(i), InStrRev(gsUiLines(i), " ") + 1&) ' No ' nor " allowed by VB6.
            Case LeftMatch(gsUiLines(i), "   BackColor"):                       .BackColor = CLngEx(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ForeColor"):                       .ForeColor = CLngEx(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   BorderStyle"):                     .BorderStyle = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   Caption"):                         .Caption = GetStringValue(gsUiLines(i))
            Case LeftMatch(gsUiLines(i), "   ClientHeight"):                    .ClientHeight = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ClientLeft"):                      .ClientLeft = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ClientTop"):                       .ClientTop = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ClientWidth"):                     .ClientWidth = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   Icon"):                            .Icon = FrxIcon(gsUiLines(i))
            Case LeftMatch(gsUiLines(i), "   Picture"):                         .Picture = FrxImage(gsUiLines(i))
            '
            Case LeftMatch(gsUiLines(i), "   BeginProperty Font"):          bInFont = True
            Case bInFont And LeftMatch(gsUiLines(i), "      Name"):             .Font.Name = GetStringValue(gsUiLines(i))
            Case bInFont And LeftMatch(gsUiLines(i), "      Size"):             .Font.Size = CSng(AfterEqual(gsUiLines(i)))
            Case bInFont And LeftMatch(gsUiLines(i), "      Weight"):           .Font.Weight = CLng(AfterEqual(gsUiLines(i)))
            Case bInFont And LeftMatch(gsUiLines(i), "      Underline"):        .Font.Underline = CBool(AfterEqual(gsUiLines(i)))
            Case bInFont And LeftMatch(gsUiLines(i), "      Italic"):           .Font.Italic = CBool(AfterEqual(gsUiLines(i)))
            Case bInFont And LeftMatch(gsUiLines(i), "      Strikethrough"):    .Font.Strikethrough = CBool(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   EndProperty") And bInFont:     bInFont = False
            '
            Case LeftMatch(gsUiLines(i), "   ControlBox"):                      .ControlBox = CLng(AfterEqual(gsUiLines(i))) <> 0&
            Case LeftMatch(gsUiLines(i), "   MaxButton"):                       .MaxButton = CLng(AfterEqual(gsUiLines(i))) <> 0&
            Case LeftMatch(gsUiLines(i), "   MinButton"):                       .MaxButton = CLng(AfterEqual(gsUiLines(i))) <> 0&
            Case LeftMatch(gsUiLines(i), "   MDIChild"):                        .MDIChild = CBool(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ScaleHeight"):                     .ScaleHeight = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   ScaleWidth"):                      .ScaleWidth = CLng(AfterEqual(gsUiLines(i)))
            Case LeftMatch(gsUiLines(i), "   StartUpPosition"):                 .StartUpPosition = CLng(AfterEqual(gsUiLines(i)))
            ' It's possible that Tag's value may have spilled into the FRX file.
            Case LeftMatch(gsUiLines(i), "   Tag"):                             .Tag = GetStringValue(gsUiLines(i))
            Case LeftMatch(gsUiLines(i), "   Enabled"):                         .Enabled = CLng(AfterEqual(gsUiLines(i))) <> 0&
            Case LeftMatch(gsUiLines(i), "   Visible"):                         .Visible = CLng(AfterEqual(gsUiLines(i))) <> 0&
            Case LeftMatch(gsUiLines(i), "   WindowState"):                     .WindowState = CLng(AfterEqual(gsUiLines(i)))
            End Select
        Next
        '
        If .Font.Weight >= 700& Then nop:                                       .Font.Bold = True
        '
        ' If we defaulted the font (or used the old MS Sans Serif font), let's use Segoe UI.
        If .Font.Name = vbNullString Or .Font.Name = "MS Sans Serif" Then nop:  .Font.Name = "Segoe UI Semibold"
        If .Font.Size = 0! Then nop:                                            .Font.Size = 9!
        If .Font.Weight = 0& Then nop:                                          .Font.Weight = 600&
        ' Border styles forced to none, fixed single, or sizable.
        Select Case .BorderStyle
        Case vbFixedDialog: .BorderStyle = vbFixedSingle
        End Select  ' All others, good to go.
    End With ' guForm.
    '
    ' Good to go.
    PopulatedFormUdt = True
End Function

Public Function ValidatedFormUdt() As Boolean
    ' A bit of error checking, but not all bad conditions are tested.
    ' Hopefully, there's be no Notepad editing of the FRM file.
    '
    ' Make sure it's actually a VB.Form (and not a VB.MDIForm).
    If guForm.Name = vbNullString Then
        MsgBox "ERROR: No form name was found.  A typical reason for this is that it's an MDI form, which can't be converted.", vbCritical
        Exit Function
    End If
    '
    ' Make sure it's not an MDI child form.
    If guForm.MDIChild Then
        MsgBox "ERROR: This form is an MDI child form, which can't be converted.  Only standard non-MDI forms can be converted.", vbCritical
        Exit Function
    End If
    '
    ' Are Client... and Scale... the same?  If not, we can't proceed.
    If guForm.ClientHeight <> guForm.ScaleHeight Or guForm.ClientWidth <> guForm.ScaleWidth Then
        MsgBox "ERROR: The form's client dimensions don't match the form's scale dimensions.  To do these conversion, ""twips"" must be used for the scale mode.", vbCritical
        Exit Function
    End If
    '
    ' Check startup position.
    If guForm.StartUpPosition = vbStartUpOwner Then
        MsgBox "ERROR: The form's startup position was set to ""Startup Owner"" which has no meaning without MDI.  We can't convert.", vbCritical
        Exit Function
    End If
    '
    ' Good to go.
    ValidatedFormUdt = True
End Function

Public Function PopulatedCtlsUi() As Boolean
    Dim i As Long
    Dim sLine As String
    '
    ' Make sure we start out empty.
    guForm.NestLevelMax = 0&
    Erase guCtls
    '
    ' Make a "working" controls array for controls to make comparisons easier.
    Dim sWorkCtls() As String
    ReDim sWorkCtls(UBound(gsValidCtls))
    For i = 0& To UBound(gsValidCtls)
        sWorkCtls(i) = "Begin " & gsValidCtls(i) & " "
    Next
    '
    ' Make a "working" properties array for comparisons.
    Dim sWorkProps() As String
    ReDim sWorkProps(UBound(gsCtrlProps))
    For i = 0& To UBound(gsCtrlProps)
        If gsCtrlProps(i) <> "BeginProperty Font" Then ' Fonts must be handled a bit differently.
            sWorkProps(i) = gsCtrlProps(i) & " "
        Else
            sWorkProps(i) = gsCtrlProps(i)
        End If
    Next
    '
    ' Count controls, all of them, even ones we don't recognize.
    Dim iCtlsCount As Long
    Dim pLine As Long
    For pLine = 0& To UBound(gsUiLines)
        If Not LeftMatch(gsUiLines(pLine), "Begin VB.Form") Then  ' Don't include the form.
            If LeftMatch(LTrim$(gsUiLines(pLine)), "Begin ") Then ' We gather them all.
                iCtlsCount = iCtlsCount + 1&
            End If
        End If
    Next
    ' Make sure we've got controls before proceeding.
    If iCtlsCount = 0& Then
        MakeZeroToNegOneArray ArrPtr(guCtls)
        GoTo JumpForNoControlsFound
    End If
    '
    ' Now gather them and fill in all their properties.
    ReDim guCtls(iCtlsCount - 1&)
    Dim pCtl As Long
    pCtl = -1& ' We'll increment first so we start at -1.
    For pLine = 0& To UBound(gsUiLines)
        If Not LeftMatch(gsUiLines(pLine), "Begin VB.Form") Then  ' Don't include the form.
            sLine = LTrim$(gsUiLines(pLine))
            If LeftMatch(sLine, "Begin ") Then ' We gather them all then delete bad ones later.
                '
                ' Increment control pointer.
                pCtl = pCtl + 1&
                With guCtls(pCtl)
                    '
                    ' Is it a control we know about.
                    .FullClassName = "Unknown.Unknown"
                    .ClassName = "Unknown"
                    .ContainerName = "Unknown"
                    For i = 0& To UBound(sWorkCtls)
                        If LeftMatch(sLine, sWorkCtls(i)) Then
                            .Good = True
                            .FullClassName = gsValidCtls(i)
                            .ClassName = Mid$(.FullClassName, InStrRev(.FullClassName, ".") + 1&)
                            Exit For
                        End If
                    Next
                    '
                    ' Use the indentation to tell the nesting level.
                    .NestLevel = (Len(gsUiLines(pLine)) - Len(sLine) - 3&) \ 3&
                    .Name = Mid$(sLine, InStrRev(sLine, " ") + 1&)
                    '
                    ' All the rest can be skipped if it's a bad control.
                    If .Good Then
                        '
                        ' Container information.  We go backwards until we find one that's nested one less.
                        If .NestLevel = 0& Then
                            .GoodContainer = True
                            .ContainerName = guForm.Name
                        Else
                            Dim pCont As Long
                            For pCont = pCtl - 1& To 0& Step -1&
                                If .NestLevel - 1& = guCtls(pCont).NestLevel Then
                                    .GoodContainer = guCtls(pCont).Good
                                    .ContainerName = guCtls(pCont).Name
                                    .ContainerIsIndexed = guCtls(pCont).IsIndexed
                                    .ContainerIndex = guCtls(pCont).Index
                                    '
                                    guCtls(pCont).HasChild = True
                                    Exit For
                                End If
                            Next
                        End If
                        '
                        ' And now we get to fill in the rest of the properties.
                        ' If it's not a good container, don't bother.
                        If .GoodContainer Then
                            '
                            ' Preset various "default" values.
                            PresetCtrlPropDefaults guCtls(pCtl)
                            '
                            ' Build control end string.
                            Dim sEnd As String
                            sEnd = Space$(.NestLevel * 3& + 3&) & "End"
                            '
                            ' Patch up our sWorkProps() array for nesting indentation.
                            Dim sNestSpace As String
                            sNestSpace = Space$(.NestLevel * 3& + 6&) ' It's +6 because properties are indented one more than the control.
                            For i = 0& To UBound(sWorkProps)
                                sWorkProps(i) = sNestSpace & LTrim$(sWorkProps(i))
                            Next
                            '
                            ' Loop to get the rest of the properties, peeking ahead within the lines.
                            Dim pLine2 As Long
                            For pLine2 = pLine + 1& To UBound(gsUiLines)
                                If gsUiLines(pLine2) = sEnd Then Exit For   ' Don't spin to end.
                                Dim pPropType As Long
                                For pPropType = 0& To UBound(sWorkProps)
                                    If LeftMatch(gsUiLines(pLine2), sWorkProps(pPropType)) Then
                                        '
                                        ' We have to handle fonts a bit differently.
                                        If gsCtrlProps(pPropType) = "BeginProperty Font" Then
                                            '
                                            ' Peek ahead to get font properties.
                                            Dim pLine3 As Long, sLine3 As String
                                            For pLine3 = pLine2 + 1& To UBound(gsUiLines)
                                                sLine3 = Trim$(gsUiLines(pLine3))
                                                If sLine3 = "EndProperty" Then Exit For ' Don't spin to end.
                                                Select Case True
                                                Case LeftMatch(sLine3, "Name "):            .Font.Name = GetStringValue(sLine3)
                                                Case LeftMatch(sLine3, "Size "):            .Font.Size = CSng(AfterEqual(sLine3))
                                                Case LeftMatch(sLine3, "Weight "):          .Font.Weight = CLng(AfterEqual(sLine3))
                                                Case LeftMatch(sLine3, "Underline "):       .Font.Underline = CBool(AfterEqual(sLine3))
                                                Case LeftMatch(sLine3, "Italic "):          .Font.Italic = CBool(AfterEqual(sLine3))
                                                Case LeftMatch(sLine3, "Strikethrough "):   .Font.Strikethrough = CBool(AfterEqual(sLine3))
                                                End Select
                                            Next
                                        Else
                                            '
                                            ' This sets all the other individual properties.
                                            SetCtrlProp guCtls(pCtl), gsUiLines(pLine2), pPropType
                                        End If
                                    End If
                                Next
                            Next
                        End If ' Good container.
                    End If ' Good control.
                    '
                    ' Just for some testing as the controls get processed.
                    'Debug.Print .Name, .NestLevel,
                    'Debug.Print .ContainerName, .ContainerIsIndexed, .ContainerIndex, .GoodContainer
                    '
                End With
            End If ' Processing a "Begin " (control) line.
        End If ' Excluding the actual form.
    Next
    '
    ' Knock out "Image" controls with no Picture.
    For pCtl = 0& To UBound(guCtls)
        If guCtls(pCtl).ClassName = "Image" Then
            If Len(guCtls(pCtl).Picture) = 0& Then guCtls(pCtl).Good = False
        End If
    Next
    '
    ' Rename the control arrays, and also set .OrigName.
    ' We'll also create a separate object to manage these.
    ' Just as a note, this does have a small chance of a conflict,
    ' where a non-indexed control is already named ..._1 etc,
    ' but we ignore this possibility.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            .OrigName = .Name
            If .IsIndexed Then .Name = .Name & "_" & CStr(.Index)
        End With
    Next
    '
    ' Get rid of controls that are bad or have bad containers.
    pCtl = 0&
    Do While pCtl < iCtlsCount
        If guCtls(pCtl).Good = False Or guCtls(pCtl).GoodContainer = False Then
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
        GoTo JumpForNoControlsFound
    End If
    ReDim Preserve guCtls(iCtlsCount - 1&)
    '
    ' Set the "Bold" for the fonts where appropriate.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .Font.Weight >= 700& Then .Font.Bold = True
        End With
    Next
    '
    ' Make sure TabStop=False if TabIndex=-1.  This is important for Python's setFocusPolicy.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .TabIndex = -1& Then .TabStop = False
        End With
    Next
    '
    ' Just for peeking/testing.
    'For pCtl = 0& To UBound(guCtls)
    '    With guCtls(pCtl)
    '        Debug.Print .Name, .NestLevel, .Caption, .ContainerName
    '        'Debug.Print Hex$(.BackColor), .Font.Name, .Font.Size
    '    End With
    'Next
    '
    ' Some stuff populated from the guCtls array.
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .NestLevel > guForm.NestLevelMax Then guForm.NestLevelMax = .NestLevel
            If .ClassName = "Menu" Then guForm.HasMenu = True
        End With
    Next
    '
JumpForNoControlsFound: ' UBound(guCtls) set to -1.
    '
    ' Good to go.
    PopulatedCtlsUi = True
End Function

Public Function ValidatedCtlsUi() As Boolean
    '
    ' The containers must have scale mode of twips.
    ' For now, we just examine PictureBoxes, as it's the only thing that has ScaleMode (other than forms).  Frames don't.
    Dim i As Long
    For i = 0& To UBound(guCtls)
        If guCtls(i).ClassName = "PictureBox" Then
            If guCtls(i).ScaleMode <> vbTwips Then
                ' This picturebox may be a control array, but we don't worry about that, and let the user figure out the problem.
                MsgBox "ERROR: There is a PictureBox (" & guCtls(i).Name & ") with a ScaleMode other than twips.  We can't proceed.", vbCritical
                Exit Function
            End If
        End If
    Next
    ValidatedCtlsUi = True
End Function

Private Sub PresetCtrlPropDefaults(uCtrl As CtrlType)
    ' Many default to 0, so we don't bother with those.
    '
    With uCtrl
        ' Font.  It's special so we handle it first and get it out of the way.
        ' It always defaults to the form's font properties (even if/when the control is in a container).
        .Font = guForm.Font
        .BorderWidth = 1&       ' Pixels.
        .Enabled = True
        .FillColor = &H0&
        .FillStyle = 1&         ' Transparent.
        .LargeChange = 1&       ' Just scrollbars.
        .Max = 32767&           ' Just scrollbars.
        .Min = 0&               ' Just scrollbars.
        .SmallChange = 1&       ' Just scrollbars.
        .ScaleMode = vbTwips    ' Just a picturebox, but we set it so we can easily check it.
        .TabIndex = -1&         ' We'll use this to shove them to the end, because there will be a ZERO.
        .Visible = True         ' Only timer isn't visible by default, and we don't mess with that one.
        '
        ' Appearance.
        Select Case .ClassName
        Case "Shape", "Line":                   .Appearance = vbFlat
        Case Else:                              .Appearance = vb3D
        End Select
        '
        ' Backcolor.
        Select Case .ClassName
        Case "ComboBox", "ListBox", "TextBox":  .BackColor = &H80000005
        Case "Line":                            .BackColor = &H80000008
        Case Else:                              .BackColor = &H8000000F
        End Select
        '
        ' BackStyle.
        Select Case .ClassName
        Case "Shape":                           .BackStyle = vbTransparent
        Case Else:                              .BackStyle = vbOpaque
        End Select
        '
        ' BorderStyle.
        Select Case .ClassName
        Case "Label", "OptionButton", "HScrollBar", "VScrollBar", "Image": _
                                                .BorderStyle = 0& ' None.
        Case Else:                              .BorderStyle = 1& ' Fixed single (or Solid for line).
        End Select
        '
        ' ForeColor.
        Select Case .ClassName
        Case "ComboBox", "ListBox", "TextBox", "Line": _
                                                .ForeColor = &H80000008
        Case Else:                              .ForeColor = &H80000012
        End Select
        '
        ' TabStop, it depends.
        Select Case .ClassName
        Case "ComboBox", "ListBox", "TextBox", _
             "CommandButton", "CheckBox", "HScrollBar", "VScrollBar", _
             "PictureBox", "OptionButton": _
                                                .TabStop = True
        ' All others default to 0 (False).
        End Select
    End With
End Sub

Private Sub SetCtrlProp(uCtrl As CtrlType, sLine As String, pPropType As Long)
    With uCtrl
        Select Case gsCtrlProps(pPropType)
        Case "Alignment":           .Alignment = CLng(AfterEqual(sLine))
        Case "Appearance":          .Appearance = CLng(AfterEqual(sLine))
        Case "BackColor":           .BackColor = CLngEx(AfterEqual(sLine))
        Case "BackStyle":           .BackStyle = CLng(AfterEqual(sLine))
        Case "BorderColor":         .BorderColor = CLngEx(AfterEqual(sLine))
        Case "BorderStyle"
                                    .BorderStyle = CLng(AfterEqual(sLine))
            Select Case .ClassName
            Case "Line"
                If .BorderStyle < 1& Or .BorderStyle > 5& Then ' Force to solid border, if it's one we don't recognize.
                                    .BorderStyle = 1&
                End If
            End Select
        Case "BorderWidth":         .BorderWidth = CLng(AfterEqual(sLine))
        Case "Cancel":              .Cancel = CBool(AfterEqual(sLine))         ' Boolean
        Case "Caption":             .Caption = GetStringValue(sLine)           ' String
        Case "Default":             .Default = CBool(AfterEqual(sLine))        ' Boolean
        Case "Enabled":             .Enabled = CBool(AfterEqual(sLine))        ' Boolean
        Case "FillColor":           .FillColor = CLngEx(AfterEqual(sLine))
        Case "FillStyle":           .FillStyle = CLng(AfterEqual(sLine))
        'Case "Font"                ' We must handle this one differently (see above caller).
        Case "ForeColor":           .ForeColor = CLngEx(AfterEqual(sLine))
        Case "Height":              .Height = CLng(AfterEqual(sLine))
        Case "Index"                ' Has two separate settings:
                                    .Index = CLng(AfterEqual(sLine))
                                    .IsIndexed = True
        Case "Interval":            .Interval = CLng(AfterEqual(sLine))
        Case "LargeChange":         .LargeChange = CLng(AfterEqual(sLine))
        Case "Left":                .Left = CLng(AfterEqual(sLine))
        Case "List":                .List = FrxList(sLine)                      ' vbNullChar delimited list of strings.
        Case "Locked":              .Locked = CBool(AfterEqual(sLine))          ' Boolean
        Case "Max":                 .Max = CLng(AfterEqual(sLine))
        Case "MaxLength":           .MaxLength = CLng(AfterEqual(sLine))
        Case "Min":                 .Min = CLng(AfterEqual(sLine))
        Case "MultiLine":           .MultiLine = CBool(AfterEqual(sLine))       ' Boolean
        Case "MultiSelect":         .MultiSelect = CLng(AfterEqual(sLine))
        Case "Picture" ' Do this only for controls we know we'll consider for pictures (to prevent saving superfluous pictures):
            Select Case .ClassName
            Case "Image", "PictureBox"
                                    .Picture = FrxImage(sLine)
            End Select
        Case "ScaleMode":           .ScaleMode = CLng(AfterEqual(sLine))
        Case "ScrollBars":          .ScrollBars = CLng(AfterEqual(sLine))
        Case "Shape":               .Shape = CLng(AfterEqual(sLine))
        Case "SmallChange":         .SmallChange = CLng(AfterEqual(sLine))
        Case "Sorted":              .Sorted = CBool(AfterEqual(sLine))          ' Boolean
        Case "Stretch":             .Stretch = CBool(AfterEqual(sLine))         ' Boolean
        Case "Style":               .Style = CLng(AfterEqual(sLine))
        Case "TabIndex":            .TabIndex = CLng(AfterEqual(sLine))
        Case "TabStop":             .TabStop = CBool(AfterEqual(sLine))         ' Boolean
        Case "Tag":                 .Tag = GetStringValue(sLine)                ' String
        Case "Text":                .Text = GetStringValue(sLine)               ' String
        Case "ToolTipText":         .ToolTipText = GetStringValue(sLine)        ' String
        Case "Top":                 .Top = CLng(AfterEqual(sLine))
        Case "Value":               .Value = CVar(AfterEqual(sLine))            ' Variant
        Case "Visible":             .Visible = CBool(AfterEqual(sLine))         ' Boolean
        Case "Width":               .Width = CLng(AfterEqual(sLine))
        Case "WordWrap":            .WordWrap = CBool(AfterEqual(sLine))        ' Boolean
        Case "X1":                  .X1 = CLng(AfterEqual(sLine))
        Case "X2":                  .X2 = CLng(AfterEqual(sLine))
        Case "Y1":                  .Y1 = CLng(AfterEqual(sLine))
        Case "Y2":                  .Y2 = CLng(AfterEqual(sLine))
        '
        End Select
    End With
End Sub

Public Sub MakeEverythingPixels()
    Const TwipsPerPixel As Long = 15&
    '
    ' Things on the form.
    With guForm
        .ClientWidth = .ClientWidth \ TwipsPerPixel
        .ClientHeight = .ClientHeight \ TwipsPerPixel
        .ClientLeft = .ClientLeft \ TwipsPerPixel
        .ClientTop = .ClientTop \ TwipsPerPixel
    End With
    '
    ' And now all the controls.
    Dim i As Long
    For i = 0& To UBound(guCtls)
        With guCtls(i)
            .Width = .Width \ TwipsPerPixel
            .Height = .Height \ TwipsPerPixel
            .Left = .Left \ TwipsPerPixel
            .Top = .Top \ TwipsPerPixel
            .X1 = .X1 \ TwipsPerPixel
            .X2 = .X2 \ TwipsPerPixel
            .Y1 = .Y1 \ TwipsPerPixel
            .Y2 = .Y2 \ TwipsPerPixel
        End With
    Next
End Sub

