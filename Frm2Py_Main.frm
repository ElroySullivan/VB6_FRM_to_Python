VERSION 5.00
Begin VB.Form frm_Frm2Py_Main 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "version info here, use Project's Title to change what goes here"
   ClientHeight    =   6795
   ClientLeft      =   18735
   ClientTop       =   6585
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   12
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm2Py_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7290
   Begin VB.CheckBox chkQuiet 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Run quietly, and exit app when done.  It will still show errors."
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   420
      TabIndex        =   5
      Top             =   6000
      Width           =   4455
   End
   Begin VB.CheckBox chkOverwriteEvents 
      BackColor       =   &H00F0F0F0&
      Caption         =   $"Frm2Py_Main.frx":0442
      ForeColor       =   &H00800000&
      Height          =   1995
      Left            =   780
      TabIndex        =   4
      Top             =   2820
      Width           =   6255
   End
   Begin VB.CheckBox chkSeparateEventsFile 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Put the form and control events in a separate Python file, using an ""_events"" suffix on this events file (recommended)."
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   420
      TabIndex        =   3
      Top             =   1740
      Value           =   1  'Checked
      Width           =   6375
   End
   Begin VB.CheckBox chkSameCoreName 
      BackColor       =   &H00F0F0F0&
      Caption         =   $"Frm2Py_Main.frx":0593
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   420
      TabIndex        =   2
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Begin a Conversion"
      Default         =   -1  'True
      Height          =   675
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   2895
   End
   Begin VB.TextBox txtWorking 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1230
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Frm2Py_Main.frx":065B
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frm_Frm2Py_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = App.Title
    Me.Left = (Screen.Width - Me.Width) / 2!
    Me.Top = (Screen.Height - Me.Height) / 3!
    '
    ' Set settings to same as last time.
    ' If nothing in registry, use VB6 form designer defaults.
    On Error Resume Next ' In case some garbage got into the registry.
        Me.chkSameCoreName.Value = -CInt(GetSetting("VB6_FRM_to_PY", "Settings", "SameCoreName", Me.chkSameCoreName.Value = vbChecked) = True)
        Me.chkSeparateEventsFile.Value = -CInt(GetSetting("VB6_FRM_to_PY", "Settings", "SeparateEvents", Me.chkSeparateEventsFile.Value = vbChecked) = True)
        Me.chkOverwriteEvents.Value = -CInt(GetSetting("VB6_FRM_to_PY", "Settings", "OverwriteEvents", Me.chkOverwriteEvents.Value = vbChecked) = True)
        Me.chkQuiet.Value = -CInt(GetSetting("VB6_FRM_to_PY", "Settings", "RunQuietly", Me.chkQuiet.Value = vbChecked) = True)
    On Error GoTo 0
    Me.chkOverwriteEvents.Enabled = Me.chkSeparateEventsFile.Value = vbChecked
    '
    ' Initializations.
    InitializeControlsArrays
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_Frm2Py_Main = Nothing
End Sub

Private Sub chkSeparateEventsFile_Click()
    Me.chkOverwriteEvents.Enabled = Me.chkSeparateEventsFile.Value = vbChecked
End Sub

Private Sub cmdBegin_Click()
    '
    ' Save settings for next time.
    SaveSetting "VB6_FRM_to_PY", "Settings", "SameCoreName", Me.chkSameCoreName.Value = vbChecked
    SaveSetting "VB6_FRM_to_PY", "Settings", "SeparateEvents", Me.chkSeparateEventsFile.Value = vbChecked
    SaveSetting "VB6_FRM_to_PY", "Settings", "OverwriteEvents", Me.chkOverwriteEvents.Value = vbChecked
    SaveSetting "VB6_FRM_to_PY", "Settings", "RunQuietly", Me.chkQuiet.Value = vbChecked
    '
    gbSeparateEventsFile = Me.chkSeparateEventsFile.Value = vbChecked
    gbOverwriteEvents = Me.chkOverwriteEvents.Value = vbChecked
    If gbSeparateEventsFile = False Then gbOverwriteEvents = True
    gbSameCoreFileName = Me.chkSameCoreName.Value = vbChecked
    gbRunQuietly = Me.chkQuiet.Value = vbChecked
    '
    ' And, let's do it.
    ' There are lots of places where errors could be found, or the user may have cancelled.
    ' That's why most things are functions (returning Booleans, for success).
    ' It MUST all be done in the following order.
    Me.Enabled = False
    Me.txtWorking.ZOrder
    Me.txtWorking.Left = 0
    Me.txtWorking.Top = 0
    Me.txtWorking.Width = Me.ScaleWidth
    Me.txtWorking.Height = Me.ScaleHeight
    Me.txtWorking.Visible = True
    
    If Not GotFrmInputFileSpec Then GoTo GetOut
    If Not GotPythonOutputFileSpec Then GoTo GetOut ' We've got to do this early to get the file names.
    OpenFrxFile
    If Not LoadedFrmFileAndCleanedUp Then GoTo GetOut
    If Not PopulatedFormUdt Then GoTo GetOut
    If Not ValidatedFormUdt Then GoTo GetOut
    If Not PopulatedCtlsUi Then GoTo GetOut
    If Not ValidatedCtlsUi Then GoTo GetOut
    MakeEverythingPixels
    '
    OpenPythonFile
    WritePythonHeader
    WritePythonFormAndWidgetClasses
    WriteModuleLevelProcsAndClasses
    '
    OpenPythonEventsFile ' If necessary, this will close the main Python file.
    WritePythonEventsCode
    WriteTestingCode
    ClosePythonFile
    CloseFrxFile
    '
    ' Done, tell user.
    Dim sRet As String
    If gbRunQuietly Then
        sRet = "Exit App"
    Else
        sRet = MsgBoxEx(Me.hWnd, vbOKCancel Or vbInformation, "Done with conversion of FRM to PY.", "Exit App", "OK")
    End If
    If sRet = "Exit App" Then
        Unload Me
        Exit Sub
    End If
    '
    ' Re-enable form and get out.
GetOut:
    ClosePythonFile ' These occur twice in case we jump out early.
    CloseFrxFile    ' These occur twice in case we jump out early.
    Me.Enabled = True
    Me.txtWorking.Visible = False
    Me.SetFocus
End Sub
