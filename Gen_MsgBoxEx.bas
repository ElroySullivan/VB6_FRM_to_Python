Attribute VB_Name = "mod_Gen_MsgBoxEx"
Option Explicit
'
Private Type MSGBOX_HOOK_PARAMS
   hWndOwner   As Long
   hHook       As Long
End Type
'
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'
Private miStyle As Long
Private msTitle As String
'
Private TimerID As Long
Private TimedOut As Boolean
'
Private msBut1 As String
Private msBut2 As String
Private msBut3 As String
'
Private mhWndMsgBox As Long
Private MSGHOOK As MSGBOX_HOOK_PARAMS
'

Public Function MsgBoxEx(hWndOwner As Long, ButtonsAndIcon As VBA.VbMsgBoxStyle, Message As String, _
                         ButA As String, Optional ButB As String, Optional ButC As String, _
                         Optional MilliSeconds As Long, Optional Title As String) As String
    ' This function sets your custom parameters and returns which button was pressed as a string.
    ' If a message box is a style with a "Cancel" button, the "X" on the form will be enabled.
    ' If the "X" of the form is pressed, the text of the last button (corresponding to cancel) will be returned.
    ' On an mbOkOnly textbox, the "X" will return the OK button.
    ' MilliSeconds is a timer.  If it times out, the function returns "TimedOut" string.  Can't be used with (mbAbortRetryIgnore and mbYesNo styles).
    '
    Dim mReturn As Long
    Dim hInstance As Long
    Dim hThreadId As Long
    '
    Const WH_CBT As Long = 5&
    Const GWL_HINSTANCE As Long = -6&
    '
    If ButtonsAndIcon = 0& Then ButtonsAndIcon = vbOK
    '
    miStyle = ButtonsAndIcon And &H7&       ' This isolates the buttons.
    If Len(Title) Then msTitle = Title Else msTitle = App.Title
    msBut1 = ButA
    msBut2 = ButB
    msBut3 = ButC
    '
    hInstance = App.hInstance
    hThreadId = GetCurrentThreadId()
    MSGHOOK.hWndOwner = GetDesktopWindow()
    MSGHOOK.hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, hInstance, hThreadId)
    '
    ' No default value is defined for mbAbortRetryIgnore and mbYesNo message box styles.
    ' In other words, the close (x) button is disabled, and the timer can NOT close the box.
    ' Therefore, timer can't be used.
    If miStyle = vbAbortRetryIgnore Or miStyle = vbYesNo Then MilliSeconds = 0&
    If MilliSeconds <> 0& Then TimerID = SetTimer(0&, 0&, MilliSeconds, AddressOf MsgBoxTimerProc)
    '
    Const MB_TASKMODAL As Long = &H2000&
    'mReturn = MessageBox(hWndOwner, Message, Space$(120), ButtonsAndIcon)
    mReturn = MessageBox(0&, Message, Space$(120), ButtonsAndIcon Or MB_TASKMODAL)
    '
    If TimerID <> 0& Then KillTimer 0&, TimerID: TimerID = 0&
    If TimedOut Then
        MsgBoxEx = "TimedOut"
        TimedOut = False
    Else
        Select Case mReturn
        Case vbOK:            MsgBoxEx = msBut1
        Case vbAbort:         MsgBoxEx = msBut1
        Case vbRetry:         MsgBoxEx = msBut2
        Case vbIgnore:        MsgBoxEx = msBut3
        Case vbYes:           MsgBoxEx = msBut1
        Case vbNo:            MsgBoxEx = msBut2
        Case vbCancel ' This may be the second or third button.
            If miStyle = vbYesNoCancel Then
                MsgBoxEx = msBut3
            Else
                MsgBoxEx = msBut2
            End If
        End Select
    End If
End Function

Public Function MsgBoxHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' This function catches the messagebox before it opens
    ' and changes the text of the buttons - then removes the hook.
    Const IDPROMPT As Long = &HFFFF&
    Const IDOK As Long = 1&
    Const IDCANCEL As Long = 2&
    Const IDAbort As Long = 3&
    Const IDRETRY As Long = 4&
    Const IDIGNORE As Long = 5&
    Const IDYES As Long = 6&
    Const IDNO As Long = 7&
    Const HCBT_ACTIVATE As Long = 5&
    '
    If uMsg = HCBT_ACTIVATE Then
        mhWndMsgBox = wParam
        '
        SetWindowText wParam, msTitle
        'SetDlgItemText wParam, IDPROMPT, mPrompt
        Select Case miStyle
        Case vbAbortRetryIgnore
            SetDlgItemText wParam, IDAbort, msBut1
            SetDlgItemText wParam, IDRETRY, msBut2
            SetDlgItemText wParam, IDIGNORE, msBut3
        Case vbYesNoCancel
            SetDlgItemText wParam, IDYES, msBut1
            SetDlgItemText wParam, IDNO, msBut2
            SetDlgItemText wParam, IDCANCEL, msBut3
        Case vbOKOnly
            SetDlgItemText wParam, IDOK, msBut1
        Case vbRetryCancel
            SetDlgItemText wParam, IDRETRY, msBut1
            SetDlgItemText wParam, IDCANCEL, msBut2
        Case vbYesNo
            SetDlgItemText wParam, IDYES, msBut1
            SetDlgItemText wParam, IDNO, msBut2
        Case vbOKCancel
            SetDlgItemText wParam, IDOK, msBut1
            SetDlgItemText wParam, IDCANCEL, msBut2
        End Select
        UnhookWindowsHookEx MSGHOOK.hHook
    End If
    MsgBoxHookProc = False
End Function

Public Function MsgBoxTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
    Const WM_CLOSE As Long = &H10&
    If IsWindow(mhWndMsgBox) Then
        PostMessage mhWndMsgBox, WM_CLOSE, 0&, ByVal 0&
        TimedOut = True
    End If
End Function

