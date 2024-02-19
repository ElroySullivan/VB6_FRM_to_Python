Attribute VB_Name = "mod_Gen_CommonDialog"
Option Explicit
'
' These are used to get information about how the dialog went.
Public FileDialogSpec As String             ' Used for both Open & Save.
Public FileDialogFolder As String           ' Used for both Open & Save.
Public FileDialogName As String             ' Used for both Open & Save.
Public FileDialogSuccessful As Boolean      ' Used for both Open & Save.
'
Public ColorDialogSuccessful As Boolean
Public ColorDialogColor As Long
'
Private Type OpenSaveType
    lStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type
Private Type ChooseColorType
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    rgbResult           As Long
    lpCustColors        As Long
    flags               As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
Private Enum ChooseColorFlagsEnum
    CC_RGBINIT = &H1                  ' Make the color specified by rgbResult be the initially selected color.
    CC_FULLOPEN = &H2                 ' Automatically display the Define Custom Colors half of the dialog box.
    CC_PREVENTFULLOPEN = &H4          ' Disable the button that displays the Define Custom Colors half of the dialog box.
    CC_SHOWHELP = &H8                 ' Display the Help button.
    CC_ENABLEHOOK = &H10              ' Use the hook function specified by lpfnHook to process the Choose Color box's messages.
    CC_ENABLETEMPLATE = &H20          ' Use the dialog box template identified by hInstance and lpTemplateName.
    CC_ENABLETEMPLATEHANDLE = &H40    ' Use the preloaded dialog box template identified by hInstance, ignoring lpTemplateName.
    CC_SOLIDCOLOR = &H80              ' Only allow the user to select solid colors. If the user attempts to select a non-solid color, convert it to the closest solid color.
    CC_ANYCOLOR = &H100               ' Allow the user to select any color.
End Enum
#If False Then ' Intellisense fix.
    Public CC_RGBINIT, CC_FULLOPEN, CC_PREVENTFULLOPEN, CC_SHOWHELP, CC_ENABLEHOOK, CC_ENABLETEMPLATE, CC_ENABLETEMPLATEHANDLE, CC_SOLIDCOLOR, CC_ANYCOLOR
#End If
Private Type KeyboardInput        '
    dwType As Long                ' Set to INPUT_KEYBOARD.
    wVK As Integer                ' shift, ctrl, menukey, or the key itself.
    wScan As Integer              ' Not being used.
    dwFlags As Long               '            HARDWAREINPUT hi;
    dwTime As Long                ' Not being used.
    dwExtraInfo As Long           ' Not being used.
    dwPadding As Currency         ' Not being used.
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const WM_LBUTTONDBLCLK  As Long = 515&
Private Const WM_SHOWWINDOW     As Long = 24&
Private Const WM_SETTEXT        As Long = &HC&
Private Const INPUT_KEYBOARD    As Long = 1&
Private Const KEYEVENTF_KEYUP   As Long = 2&
Private Const KEYEVENTF_KEYDOWN As Long = 0&
'
Private muEvents(1) As KeyboardInput    ' Just used to emulate "Enter" key.
Private pt32 As POINTAPI
Private msColorTitle As String
'
Public Enum FileDialogFlags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000         '  force no long names for 4.x modules
    OFN_EXPLORER = &H80000            '  new look commdlg
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000          '  force long names for 3.x modules
End Enum
#If False Then ' Intellisense fix.
    Public OFN_READONLY, OFN_OVERWRITEPROMPT, OFN_HIDEREADONLY, OFN_NOCHANGEDIR, OFN_SHOWHELP, OFN_ENABLEHOOK, OFN_ENABLETEMPLATE, OFN_ENABLETEMPLATEHANDLE, OFN_NOVALIDATE, OFN_ALLOWMULTISELECT, OFN_EXTENSIONDIFFERENT, OFN_PATHMUSTEXIST, OFN_FILEMUSTEXIST
    Public OFN_CREATEPROMPT, OFN_SHAREAWARE, OFN_NOREADONLYRETURN, OFN_NOTESTFILECREATE, OFN_NONETWORKBUTTON, OFN_NOLONGNAMES, OFN_EXPLORER, OFN_NODEREFERENCELINKS, OFN_LONGNAMES
#End If
'
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Private Declare Function SetFocusTo Lib "user32" Alias "SetFocus" (Optional ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal uFlags As Long) As Long
Private Declare Function SendMessageWLong Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
Private Const MAX_PATH_W As Long = 260&
'

Public Function ShowColorDialog(hWndOwner As Long, lDefaultColor As Long, Optional NewColor As Long, Optional Title As String = "Select Color", Optional CustomColorsHex As String) As Boolean
    ' You can optionally use ColorDialogSuccessful & ColorDialogColor or the return of ShowColorDialog and NewColor.  They will be the same.
    '
    ' CustomColorHex is a comma separated hex string of 16 custom colors.  It's best to just let the user specify these, starting out with all black.
    ' If this CustomColorHex string doesn't separate into precisely 16 values, it's ignored, resulting with all black custom colors.
    ' The string is returned, and it's up to you to save it if you wish to save your user-specified custom colors.
    ' These will be specific to this program, because this is your CustomColorsHex string.
    '
    Dim uChooseColor As ChooseColorType
    Dim CustomColors(15&) As Long
    Dim sArray() As String
    Dim i As Long
    '
    msColorTitle = Title
    '
    ' Setup custom colors.
    sArray = Split(CustomColorsHex, ",")
    If UBound(sArray) = 15& Then
        For i = 0& To 15&
            CustomColors(i) = Val("&h" & sArray(i))
        Next
    End If
    '
    uChooseColor.hWndOwner = hWndOwner
    uChooseColor.lpCustColors = VarPtr(CustomColors(0&))
    uChooseColor.flags = CC_ENABLEHOOK Or CC_FULLOPEN Or CC_RGBINIT
    uChooseColor.hInstance = App.hInstance
    uChooseColor.lStructSize = LenB(uChooseColor)
    uChooseColor.lpfnHook = ProcedureAddress(AddressOf ColorHookProc)
    uChooseColor.rgbResult = lDefaultColor
    '
    ColorDialogSuccessful = False
    If ChooseColorAPI(uChooseColor) = 0& Then
        Exit Function
    End If
    If uChooseColor.rgbResult > &HFFFFFF Then Exit Function
    '
    ColorDialogColor = uChooseColor.rgbResult
    NewColor = uChooseColor.rgbResult
    ColorDialogSuccessful = True
    ShowColorDialog = True
    '
    ' Return custom colors.
    ReDim sArray(15&)
    For i = 0& To 15&
        sArray(i) = Hex$(CustomColors(i))
    Next
    CustomColorsHex = Join(sArray, ",")
End Function

Private Function ColorHookProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_SHOWWINDOW Then
        SetWindowText hWnd, msColorTitle
        ColorHookProc = 1&
    End If
    '
    If uMsg = WM_LBUTTONDBLCLK Then
        '
        ' If we're on a hWnd with text, we probably should ignore the double-click.
        GetCursorPos pt32
        ScreenToClient hWnd, pt32
        '
        If WindowText(ChildWindowFromPointEx(hWnd, pt32.X, pt32.Y, 0&)) = vbNullString Then
            ' For some reason, this SetFocus is necessary for the dialog to receive keyboard input under certain circumstances.
            SetFocusTo hWnd
            ' Build EnterKeyDown & EnterKeyDown events.
            muEvents(0&).wVK = vbKeyReturn: muEvents(0&).dwFlags = KEYEVENTF_KEYDOWN: muEvents(0&).dwType = INPUT_KEYBOARD
            muEvents(1&).wVK = vbKeyReturn: muEvents(1&).dwFlags = KEYEVENTF_KEYUP:   muEvents(1&).dwType = INPUT_KEYBOARD
            ' Put it on buffer.
            SendInput 2&, muEvents(0&), Len(muEvents(0&))
            ColorHookProc = 1&
        End If
    End If
End Function

Public Sub ShowOpenFileDialog(hWndOwner As Long, FileFilter As String, Optional InitialFolder As String, Optional flags As FileDialogFlags, Optional Title As String = "Open File")
    ' See "Public" variables for what is set upon return.
    '
    ' Example:
    '    ShowOpenFileDialog hWnd, "Access Databases (*.mdb, *.mde)" & vbNullChar & "*.mdb; *.mde" & vbNullChar, "C:\"
    '    If FileDialogSuccessful = True Then
    '        MsgBox "You selected file: " & FileDialogSpec
    '    Else
    '        MsgBox "No file selected."
    '    End If
    Dim OpenFile As OpenSaveType
    Dim l As Long
    '
    OpenFile.lStructSize = LenB(OpenFile)
    OpenFile.hWndOwner = hWndOwner
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = FileFilter
    OpenFile.nFilterIndex = 1&
    OpenFile.lpstrFile = String$(MAX_PATH_W, vbNullChar)
    OpenFile.nMaxFile = MAX_PATH_W
    OpenFile.lpstrFileTitle = vbNullString
    OpenFile.nMaxFileTitle = 0&
    OpenFile.lpstrInitialDir = InitialFolder
    OpenFile.lpstrTitle = Title
    OpenFile.flags = flags
    '
    l = GetOpenFileNameW(VarPtr(OpenFile))
    If l = 0& Then
        FileDialogSpec = "none"
        FileDialogFolder = vbNullString
        FileDialogName = vbNullString
        FileDialogSuccessful = False
    Else
        FileDialogSpec = RTrimNull(OpenFile.lpstrFile)
        FileDialogFolder = Left$(FileDialogSpec, InStrRev(FileDialogSpec, "\"))
        FileDialogName = Mid$(FileDialogSpec, InStrRev(FileDialogSpec, "\") + 1&)
        FileDialogSuccessful = True
    End If
End Sub

Public Sub ShowSaveFileDialog(hWndOwner As Long, FileFilter As String, Optional FilterIndex As Long = 1&, _
                              Optional InitialFolder As String, Optional flags As FileDialogFlags, Optional Title As String = "Save File", _
                              Optional InitialFile As String, Optional DefaultExt As String)
    ' See "Public" variables for what is set upon return.
    Dim SaveFile As OpenSaveType
    Dim l As Long
    '
    SaveFile.lStructSize = LenB(SaveFile)
    SaveFile.hWndOwner = hWndOwner
    SaveFile.hInstance = App.hInstance
    SaveFile.lpstrFilter = FileFilter
    SaveFile.nFilterIndex = FilterIndex
    SaveFile.lpstrFile = InitialFile & String$(MAX_PATH_W - Len(InitialFile), vbNullChar)
    SaveFile.nMaxFile = MAX_PATH_W
    SaveFile.lpstrFileTitle = vbNullString
    SaveFile.nMaxFileTitle = 0&
    SaveFile.lpstrInitialDir = InitialFolder
    SaveFile.lpstrTitle = Title
    SaveFile.lpstrDefExt = DefaultExt
    SaveFile.flags = flags
    '
    l = GetSaveFileNameW(VarPtr(SaveFile))
    If l = 0& Then
        FileDialogSpec = "none"
        FileDialogFolder = vbNullString
        FileDialogName = vbNullString
        FileDialogSuccessful = False
    Else
        FileDialogSpec = RTrimNull(SaveFile.lpstrFile)
        FileDialogFolder = Left$(FileDialogSpec, InStrRev(FileDialogSpec, "\"))
        FileDialogName = Mid$(FileDialogSpec, InStrRev(FileDialogSpec, "\") + 1&)
        FileDialogSuccessful = True
    End If
End Sub

Private Function ProcedureAddress(AddressOf_TheProc As Long) As Long
    ProcedureAddress = AddressOf_TheProc
End Function

Private Function WindowText(hWnd As Long) As String
    WindowText = Space$(GetWindowTextLength(hWnd) + 1&)
    WindowText = Left$(WindowText, GetWindowText(hWnd, WindowText, Len(WindowText)))
End Function

Public Sub SetWindowText(hWnd As Long, sText As String)
    SendMessageWLong hWnd, WM_SETTEXT, 0&, StrPtr(sText)
End Sub

Private Function RTrimNull(s As String) As String
    Dim i As Long
    i = InStr(s, vbNullChar)
    If i Then
        RTrimNull = Left$(s, i - 1&)
    Else
        RTrimNull = s
    End If
End Function

