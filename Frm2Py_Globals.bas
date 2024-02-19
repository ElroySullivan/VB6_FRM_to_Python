Attribute VB_Name = "mod_Frm2Py__Globals"
Option Explicit
'
Public Enum EventProcsEnum
    NoEvents = 0&
    OnlyUsedEvents = 1&
    AllEvents = 2&
End Enum
#If False Then ' Intellisense fix.
    Dim NoEvents, OnlyUsedEvents, AllEvents
#End If
'
Public Type FontType
    Name            As String
    Size            As Single
    Weight          As Long
    Bold            As Boolean
    Underline       As Boolean
    Italic          As Boolean
    Strikethrough   As Boolean
End Type
'
Public Type FormType
    Name            As String       ' Internal name, not the file name.
    NestLevelMax    As Long         ' The maximum nesting level found for controls (starting at 0 for form).
    HasMenu         As Boolean      ' Has a menu on the form.
    BackColor       As Long
    BorderStyle     As FormBorderStyleConstants
    Caption         As String
    ClientHeight    As Long
    ClientLeft      As Long
    ClientTop       As Long
    ClientWidth     As Long
    ControlBox      As Boolean
    Enabled         As Boolean
    Font            As FontType     ' This is used only in the cases where controls don't have font specification.
    Icon            As String       ' After processing, this will be the file's name (no path).
    MaxButton       As Boolean
    MDIChild        As Boolean      ' Just used to check to make sure it's NOT an MDI child.
    MinButton       As Boolean
    Picture         As String       ' After FRM reading, this will be the file's name (no path).
    ScaleHeight     As Long
    ScaleWidth      As Long
    StartUpPosition As StartUpPositionConstants
    Tag             As String
    Visible         As Boolean
    WindowState     As FormWindowStateConstants
End Type
'
Public Type CtrlType
    Good            As Boolean  ' Whether or not it's a recognized control.
    DoneInPython    As Boolean  ' Mostly for handling tab order.
    ClassName       As String
    FullClassName   As String
    NestLevel       As Long     ' 0 = directly on the form.
    HasChild        As Boolean  ' This container (or menu) has child(ren).
    Name            As String
    IsIndexed       As Boolean
    '
    GoodContainer   As Boolean  ' These are deleted before processing starts.
    ContainerName   As String   ' If on form, the form's name.
    ContainerIsIndexed  As Boolean  ' Form is never indexed.
    ContainerIndex  As Long
    '
    ' These are the ones in the string array of properties.
    ' And also the ones indented under the control in the FRM file.
    Alignment       As AlignmentConstants
    Appearance      As MSComctlLib.AppearanceConstants ' ccFlat or cc3D
    BackColor       As Long
    BackStyle       As Long
    BorderColor     As Long
    BorderStyle     As Long
    BorderWidth     As Long
    Cancel          As Boolean
    Caption         As String
    Default         As Boolean
    Enabled         As Boolean  ' Default is True.
    FillColor       As Long
    FillStyle       As Long
    Font            As FontType
    ForeColor       As Long
    Height          As Long
    Index           As Long     ' Be sure to check IsIndexed before using.
    Interval        As Long
    LargeChange     As Long
    Left            As Long
    List            As String
    Locked          As Boolean
    Max             As Long
    MaxLength       As Long
    Min             As Long
    MultiLine       As Boolean
    MultiSelect     As Long
    Picture         As String   ' After FRM reading, this will be the file's name (no path).
    ScaleMode       As Long     ' Be sure to set to 1 (twips) when created, and don't allow anything else.
    ScrollBars      As Long
    Shape           As Long
    SmallChange     As Long
    Sorted          As Boolean
    Stretch         As Boolean
    Style           As Long
    TabIndex        As Long
    TabStop         As Boolean
    Tag             As String
    Text            As String
    ToolTipText     As String
    Top             As Long
    Value           As Variant
    Visible         As Boolean  ' Default is True.
    Width           As Long
    WordWrap        As Boolean
    X1              As Long
    X2              As Long
    Y1              As Long
    Y2              As Long
End Type
'
Public Type TabsType ' Controls, but just for getting taborder figured out.
    ClassName       As String
    Name            As String
    TabIndex        As Long
    TabStop         As Boolean
End Type
'
' Lists of controls and their properties we'll actually process.
Public gsValidCtls() As String  ' We build a list of the controls we actually process.
Public gsCtrlProps() As String  ' List of control property names we look for.
'
' Variables set when we start converting.
Public gbSameCoreFileName       As Boolean
Public gbSeparateEventsFile     As Boolean
Public gbOverwriteEvents        As Boolean
Public gbRunQuietly             As Boolean
'
' File names.
Public gsInputFileSpec As String
Public gsInputFileName As String
Public gsInputFilePath As String ' Always includes terminating \.
Public gsInputFileBase As String ' Just the core name of the file without extension.
'
Public gsOutputFileSpec As String
Public gsOutputFileName As String
Public gsOutputFilePath As String ' Always includes terminating \.
Public gsOutputFileBase As String ' Just the core name of the file without extension.
'
Public gsOutputEventsSpec As String
Public gsOutputEventsName As String
Public gsOutputEventsPath As String ' Always includes terminating \.
Public gsOutputEventsBase As String ' Just the core name of the file without extension.
Public gsOutputEventsAltSpec As String
'
' Info from the VB6 FRM file.
Public ghFrx            As Long     ' It's ZERO if we don't have one.
Public ghPy             As Long
Public gsUiLines()      As String
Public gsCodeLines()    As String
'
Public guForm As FormType
Public guCtls() As CtrlType
'

Public Sub InitializeControlsArrays()
    Dim i As Long
    '
    ' Known controls.  Just add them here if/when there are more.
    ReDim gsValidCtls(200&)
    i = -1&
    i = i + 1&: gsValidCtls(i) = "VB.Menu"          ' Not exactly a control, but treated like one.
    '
    i = i + 1&: gsValidCtls(i) = "VB.CommandButton"
    i = i + 1&: gsValidCtls(i) = "VB.CheckBox"
    i = i + 1&: gsValidCtls(i) = "VB.ComboBox"
    i = i + 1&: gsValidCtls(i) = "VB.Frame"
    i = i + 1&: gsValidCtls(i) = "VB.HScrollBar"
    i = i + 1&: gsValidCtls(i) = "VB.Image"
    i = i + 1&: gsValidCtls(i) = "VB.Label"
    i = i + 1&: gsValidCtls(i) = "VB.Line"
    i = i + 1&: gsValidCtls(i) = "VB.ListBox"
    i = i + 1&: gsValidCtls(i) = "VB.PictureBox"
    i = i + 1&: gsValidCtls(i) = "VB.OptionButton"
    i = i + 1&: gsValidCtls(i) = "VB.Shape"
    i = i + 1&: gsValidCtls(i) = "VB.TextBox"
    i = i + 1&: gsValidCtls(i) = "VB.VScrollBar"
    i = i + 1&: gsValidCtls(i) = "VB.TreeView"
    ' Now fix the ReDim.
    ReDim Preserve gsValidCtls(i)
    '
    ' And now the list of control properties.
    ' These are only the ones AFTER we take care of the primary ones.
    ReDim gsCtrlProps(200&)
    i = -1&
    i = i + 1&: gsCtrlProps(i) = "Alignment"
    i = i + 1&: gsCtrlProps(i) = "Appearance"
    i = i + 1&: gsCtrlProps(i) = "BackColor"
    i = i + 1&: gsCtrlProps(i) = "BackStyle"
    i = i + 1&: gsCtrlProps(i) = "BorderColor"
    i = i + 1&: gsCtrlProps(i) = "BorderStyle"
    i = i + 1&: gsCtrlProps(i) = "BorderWidth"
    i = i + 1&: gsCtrlProps(i) = "Cancel"
    i = i + 1&: gsCtrlProps(i) = "Caption"
    i = i + 1&: gsCtrlProps(i) = "Default"
    i = i + 1&: gsCtrlProps(i) = "Enabled"
    i = i + 1&: gsCtrlProps(i) = "FillColor"
    i = i + 1&: gsCtrlProps(i) = "FillStyle"
    i = i + 1&: gsCtrlProps(i) = "BeginProperty Font" ' This is handled differently.
    i = i + 1&: gsCtrlProps(i) = "ForeColor"
    i = i + 1&: gsCtrlProps(i) = "Height"
    i = i + 1&: gsCtrlProps(i) = "Index"
    i = i + 1&: gsCtrlProps(i) = "Interval"
    i = i + 1&: gsCtrlProps(i) = "LargeChange"
    i = i + 1&: gsCtrlProps(i) = "Left"
    i = i + 1&: gsCtrlProps(i) = "List"
    i = i + 1&: gsCtrlProps(i) = "Locked"
    i = i + 1&: gsCtrlProps(i) = "Max"
    i = i + 1&: gsCtrlProps(i) = "MaxLength"
    i = i + 1&: gsCtrlProps(i) = "Min"
    i = i + 1&: gsCtrlProps(i) = "MultiLine"
    i = i + 1&: gsCtrlProps(i) = "MultiSelect"
    i = i + 1&: gsCtrlProps(i) = "Picture"
    i = i + 1&: gsCtrlProps(i) = "ScaleMode"
    i = i + 1&: gsCtrlProps(i) = "ScrollBars"
    i = i + 1&: gsCtrlProps(i) = "Shape"
    i = i + 1&: gsCtrlProps(i) = "SmallChange"
    i = i + 1&: gsCtrlProps(i) = "Sorted"
    i = i + 1&: gsCtrlProps(i) = "Stretch"
    i = i + 1&: gsCtrlProps(i) = "Style"
    i = i + 1&: gsCtrlProps(i) = "TabIndex"
    i = i + 1&: gsCtrlProps(i) = "TabStop"
    i = i + 1&: gsCtrlProps(i) = "Tag"
    i = i + 1&: gsCtrlProps(i) = "Text"
    i = i + 1&: gsCtrlProps(i) = "ToolTipText"
    i = i + 1&: gsCtrlProps(i) = "Top"
    i = i + 1&: gsCtrlProps(i) = "Value"
    i = i + 1&: gsCtrlProps(i) = "Visible"
    i = i + 1&: gsCtrlProps(i) = "Width"
    i = i + 1&: gsCtrlProps(i) = "WordWrap"
    i = i + 1&: gsCtrlProps(i) = "X1"
    i = i + 1&: gsCtrlProps(i) = "X2"
    i = i + 1&: gsCtrlProps(i) = "Y1"
    i = i + 1&: gsCtrlProps(i) = "Y2"
    ' And fix the ReDim.
    ReDim Preserve gsCtrlProps(i)
    '
    ' All good.
End Sub
