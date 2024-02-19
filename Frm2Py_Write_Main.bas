Attribute VB_Name = "mod_Frm2Py_Write___Main"
Option Explicit
'
    
Public Sub WritePythonFormAndWidgetClasses()
    '
    ' Get our form's class started.
    With guForm
        Print #ghPy, vbNullString
        Print #ghPy, "# To import the following class, using the following line."
        Print #ghPy, "#import cls"; .Name; " from "; gsOutputFileBase
        '
        ' Function that initializes.
        Print #ghPy, "class cls"; .Name; "(QMainWindow):"
        Print #ghPy, "    paint_event_raised = pyqtSignal() # So we can 'emit()' our paintEvent to other widgets for this container."
        Print #ghPy, "    def __init__(self, parent=None):"
        Print #ghPy, "        super().__init__(parent)"
        Print #ghPy, "        self.Name = '"; .Name; "'"
        Print #ghPy, "        self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
        '
        ' Setup the form's properties from the VB6 FRM info.
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        # Form's initial properties, from VB6 FRM file."
        Print #ghPy, vbNullString
        '
        ' The Tag property.
        Print #ghPy, "        self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        '
        ' Border style and width/height.
      Select Case .BorderStyle
      Case vbBSNone         ' resize sets interior.
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)"
        Print #ghPy, "        self.resize("; CStr(.ClientWidth); ", "; CStr(.ClientHeight); ")"
      Case vbSizable        ' resize sets interior.
        Print #ghPy, "        self.resize("; CStr(.ClientWidth); ", "; CStr(.ClientHeight); ")"
      Case vbFixedSingle    ' setFixedSize sets interior.
        Print #ghPy, "        self.setFixedSize("; CStr(.ClientWidth); ", "; CStr(.ClientHeight); ")"
      Case vbSizableToolWindow
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.Tool)"
        Print #ghPy, "        self.resize("; CStr(.ClientWidth); ", "; CStr(.ClientHeight); ")"
      Case vbFixedToolWindow
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.Tool)"
        Print #ghPy, "        self.setFixedSize("; CStr(.ClientWidth); ", "; CStr(.ClientHeight); ")"
      End Select
        '
        ' Startup position and left/top.
      Select Case .StartUpPosition
      Case vbStartUpManual
        Print #ghPy, "        self.move("; CStr(.ClientLeft - 4&); ", "; CStr(.ClientTop - 27&); ")" ' Adjustments for non-client area.
      Case vbStartUpScreen
        Print #ghPy, "        centerPoint = QDesktopWidget().availableGeometry().center()"
        Print #ghPy, "        self.move(centerPoint.x() - self.width() // 2, centerPoint.y() - self.height() // 2)"
      Case Else ' vbStartUpWindowsDefault
        Print #ghPy, "        self.move(0, 0)"
      End Select
        '
        ' Form's icon.
      If Len(.Icon) Then
        Print #ghPy, "        self.icon_spec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Icon; "')"
        Print #ghPy, "        self.setWindowIcon(QIcon(self.icon_spec))"
      Else
        Print #ghPy, "        self.icon_spec = None"
      End If
        '
        ' Caption.
        Print #ghPy, "        self.setWindowTitle('"; .Caption; "')"
        '
        ' Window state.
      Select Case .WindowState
      Case vbMinimized
        Print #ghPy, "        self.setWindowState(self.windowState() | Qt.WindowMinimized)"
      Case vbMaximized
        Print #ghPy, "        self.setWindowState(self.windowState() | Qt.WindowMaximized)"
      End Select ' Otherwise, it defaults to normalized.
        '
        ' Min/Max/Close buttons.
      Select Case True
      Case .ControlBox ' This disables everything.
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMinimizeButtonHint & ~Qt.WindowMaximizeButtonHint & ~Qt.WindowCloseButtonHint)"
      Case .MinButton And .MaxButton
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMinimizeButtonHint & ~Qt.WindowMaximizeButtonHint)"
      Case .MinButton
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMinimizeButtonHint)"
      Case .MaxButton
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)"
      End Select ' Else just leave the defaults.
        '
        ' Note on visibility.  This is just controlled with .Show() or .Hide(), so there's no need to address it here.
        '
        ' Enabled.
        Print #ghPy, "        self.setEnabled("; TrueFalse(.Enabled); ")"
        '
        ' BackColor.
        Print #ghPy, "        self.setAutoFillBackground(True)"
        Print #ghPy, "        palette = self.palette()"
        Print #ghPy, "        palette.setColor(QPalette.Window, QColor('"; RgbHex(.BackColor); "'))"
        Print #ghPy, "        self.setPalette(palette)"
        '
        ' Picture.  And we fit it to the form's client area in the paintEvent.
      If Len(.Picture) Then
        Print #ghPy, "        self.image_spec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
        Print #ghPy, "        self.background_pixmap = QPixmap(self.image_spec)"
      Else
        Print #ghPy, "        self.image_spec = ''"
        Print #ghPy, "        self.background_pixmap = None"
      End If
        '
        ' Any menu on the form.
      If .HasMenu Then
        Print #ghPy, vbNullString
        Print #ghPy, "        # The form's menus."
        Print #ghPy, "        # The ones with clickable actions get an self.object so they can be disabled or hidden."
        Print #ghPy, vbNullString
        Print #ghPy, "        "; .Name; "_menu = self.menuBar()"
        Call AddTheMenus(.Name) ' This is recursive and adds them all.
      End If
        '
        ' Final form's constructor (__init__) things.
        InstantiateAllTheWidgets
        SetWidgetTabOrders
        '
        ' Form's internal event procedures.
        '
        Print #ghPy, vbNullString
        Print #ghPy, "    # ****************************************************************"
        Print #ghPy, "    # We're still in our form's class, but no longer in __init__."
        Print #ghPy, "    # ****************************************************************"
        Print #ghPy, vbNullString
        Print #ghPy, "    # Internal event procedures.  They'll try and call external ones, if found."
        '
        ' closeEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def closeEvent(self, event): # event.ignore()  # Prevents the window from closing."
        Print #ghPy, "        if '"; .Name; "_Close' in globals(): "; .Name; "_Close(self, event)"
        '
        ' focusInEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def focusInEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_FocusIn' in globals(): "; .Name; "_FocusIn(self, event)"
        '
        ' focusOutEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def focusOutEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_FocusOut' in globals(): "; .Name; "_FocusOut(self, event)"
        '
        ' hideEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def hideEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Hide' in globals(): "; .Name; "_Hide(self, event)"
        '
        ' keyPressEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def keyPressEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_KeyPress' in globals(): "; .Name; "_KeyPress(self, event)"
        '
        ' keyReleaseEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def keyReleaseEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_KeyRelease' in globals(): "; .Name; "_KeyRelease(self, event)"
        '
        ' leaveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def leaveEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Leave' in globals(): "; .Name; "_Leave(self, event)"
        '
        ' mouseDoubleClickEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseDoubleClickEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_DoubleClick' in globals(): "; .Name; "_DoubleClick(self, event)"
        '
        ' mouseMoveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseMoveEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_MouseMove' in globals(): "; .Name; "_MouseMove(self, event)"
        '
        ' mousePressEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mousePressEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_MousePress' in globals(): "; .Name; "_MousePress(self, event)"
        '
        ' mouseReleaseEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseReleaseEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_MouseRelease' in globals(): "; .Name; "_MouseRelease(self, event)"
        '
        ' moveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def moveEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Move' in globals(): "; .Name; "_Move(self, event)"
        '
        ' paintEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def paintEvent(self, event):"
        Print #ghPy, "        super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
        Print #ghPy, "        if self.background_pixmap: QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.background_pixmap)"
        Print #ghPy, "        if '"; .Name; "_Paint' in globals(): "; .Name; "_Paint(self, event)"
        Print #ghPy, "        # Be sure to do .emit() after any picture, so lines get drawn on top of the picture."
        Print #ghPy, "        self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
        '
        ' resizeEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def resizeEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Resize' in globals(): "; .Name; "_Resize(self, event)"
        '
        ' showEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def showEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Show' in globals(): "; .Name; "_Show(self, event)"
        '
        ' wheelEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def wheelEvent(self, event):"
        Print #ghPy, "        if '"; .Name; "_Wheel' in globals(): "; .Name; "_Wheel(self, event)"
        '
        ' Internal menu events.
      If .HasMenu Then
        Print #ghPy, vbNullString
        Print #ghPy, "    # Internal menu events."
        Dim pCtl As Long
        For pCtl = 0& To UBound(guCtls)
            With guCtls(pCtl)
                If .ClassName = "Menu" Then
                    If .Caption <> "-" Then
                        If .HasChild = False Then
                            Print #ghPy, vbNullString
                            Print #ghPy, "    def "; .Name; "_action(self):"
                            Print #ghPy, "        if '"; .Name; "_Click' in globals(): "; .Name; "_Click(self)"
                        End If
                    End If
                End If
            End With
        Next
      End If
    End With
    '
    'DoStandardFormMethodsAndProperties     ' We no longer do any of these.  Let the user use PyQt instead.
    DoWidgets_Class_Meth_Prop_IntEvt
End Sub
    
Private Sub AddTheMenus(sParent As String)
    ' Called from above.  This is recursive and adds them all.
    '
    Dim pCtl As Long
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .ClassName = "Menu" Then             ' Herein, we process only the menu.
                If .ContainerName = sParent Then    ' Only do our parent ... recursion will take care of the rest.
                    Select Case True
                    Case .Caption = "-"             ' It's just a separator.
                        Print #ghPy, vbNullString
                        Print #ghPy, "        "; sParent; "_menu.addSeparator()"
                    Case .HasChild                  ' It has sub-menu items, so do recursion.
                        Print #ghPy, vbNullString
                        Print #ghPy, "        "; .Name; "_menu = "; sParent; "_menu.addMenu('"; .Caption; "')"
                        Call AddTheMenus(.Name)
                    Case Else                       ' It's a clickable action menu item.
                        Print #ghPy, vbNullString
                        Print #ghPy, "        self."; .Name; " = QAction('"; .Caption; "', self)"
                        Print #ghPy, "        self."; .Name; ".triggered.connect(self."; .Name; "_action)"
                        Print #ghPy, "        "; sParent; "_menu.addAction(self."; .Name; ")"
                      If .Enabled = False Then
                        Print #ghPy, "        self."; .Name; ".setEnabled(False)"
                      End If
                      If .Visible = False Then
                        Print #ghPy, "        self."; .Name; ".setVisible(False)"
                      End If
                    End Select
                End If
            End If
        End With
    Next
End Sub
    
    
    
Public Sub WritePythonEventsAndTestCode()
    ' Called by main form.  Depending on option selected, this might go into the _EVENTS file.
    '
    ' If we don't overwrite, skip this ... but we changed it to write a (##) file instead.
    'If Not gbOverwriteEvents Then Exit Sub
    '
    ' ************* Back to module level (out of class) *************
    '
    ' And now we'll just put some event procedure stubs in at the module level.
    ' As with the bindings, we just comment out the ones not used, so we can see how to do them.
  If gbSeparateEventsFile Then
    Print #ghPy, vbNullString
    Print #ghPy, "# Initially created with "; App.Title; " written by Elroy Sullivan, PhD."
    Print #ghPy, vbNullString
    Print #ghPy, "from "; gsOutputFileBase; " import *"
    Print #ghPy, "from PyQt5.QtWidgets import QApplication"
  End If
    '
    Print #ghPy, vbNullString
    Print #ghPy, "# ****************************************************************************"
    Print #ghPy, "# All the form level, form's menu, & widget event procedures for coding."
    Print #ghPy, "# Any (or all) of these that aren't needed can be deleted without harm."
    Print #ghPy, "#"
    Print #ghPy, "# There are others, but you're left to your own devices to do those bindings"
    Print #ghPy, "# via the PyQt library.  Also, any other properties you wish to use can"
    Print #ghPy, "# also be done with the PyQt library.  The form inherits the QMainWindow"
    Print #ghPy, "# interface, and the widgets each inherit their appropriate PyQt interfaces,"
    Print #ghPy, "# so all of these PyQt methods & properties can be directly accessed via"
    Print #ghPy, "# the form and widget objects.  All the widget objects are nested within"
    Print #ghPy, "# the form object."
    Print #ghPy, "# ****************************************************************************"
    '
    DoFormEventProcedures
    If guForm.HasMenu Then DoFormMenuProcedures
    DoExternalWidgetEventProcedures
End Sub

Private Sub DoFormEventProcedures()
    ' Note that these should match the internal events.
    '
    With guForm
        Print #ghPy, vbNullString
        Print #ghPy, "# ****************************************"
        Print #ghPy, "# Form level events for coding."
        Print #ghPy, "# ****************************************"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Close(self, event): # event.ignore()  # Prevents the window from closing."
        Print #ghPy, "    print('"; .Name; "_Close', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_DoubleClick(self, event):"
        Print #ghPy, "    print('"; .Name; "_DoubleClick', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_FocusIn(self, event):"
        Print #ghPy, "    print('"; .Name; "_FocusIn', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_FocusOut(self, event):"
        Print #ghPy, "    print('"; .Name; "_FocusOut', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Hide(self, event):"
        Print #ghPy, "    print('"; .Name; "_Hide', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_KeyPress(self, event):"
        Print #ghPy, "#    print('"; .Name; "_KeyPress', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_KeyRelease(self, event):"
        Print #ghPy, "#    print('"; .Name; "_KeyRelease', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Leave(self, event):"
        Print #ghPy, "    print('"; .Name; "_Leave', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MouseMove(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MouseMove', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MousePress(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MousePress', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MouseRelease(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MouseRelease', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Move(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Move', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Paint(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Paint', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Resize(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Resize', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Show(self, event):"
        Print #ghPy, "    print('"; .Name; "_Show', self.Name)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Wheel(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Wheel', self.Name)"
    End With
End Sub

Private Sub DoFormMenuProcedures()
    Print #ghPy, vbNullString
    Print #ghPy, "# ****************************************"
    Print #ghPy, "# Menu events for coding."
    Print #ghPy, "# ****************************************"
    Dim pCtl As Long
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            If .ClassName = "Menu" Then
                If .Caption <> "-" Then
                    If .HasChild = False Then
                        Print #ghPy, vbNullString
                        Print #ghPy, "def "; .Name; "_Click(self):"
                        Print #ghPy, "    print('"; .Name; "_Click')"
                    End If
                End If
            End If
        End With
    Next
End Sub



' *********************************************************************************
' We keep file open/close down here because it's not frequently modified.
' *********************************************************************************

Public Function GotPythonOutputFileSpec() As Boolean
    '
    ' If it's the same core file name, it's easy.
    If gbSameCoreFileName Then
        gsOutputFileBase = gsInputFileBase
        gsOutputFilePath = gsInputFilePath
        gsOutputFileName = gsOutputFileBase & ".py"
        gsOutputFileSpec = gsOutputFilePath & gsOutputFileName
    Else
        ' Prompt for output python file.
        gsOutputFileName = gsInputFileBase & ".py"
        ShowSaveFileDialog 0, "Python (*.py)" & vbNullChar & "*.py" & vbNullChar, , gsInputFilePath, OFN_OVERWRITEPROMPT, "Converted Python File to Save", gsOutputFileName, "py"
        If FileDialogSuccessful = False Then Exit Function
        '
        ' Save our file specifications.
        gsOutputFileSpec = FileDialogSpec
        gsOutputFileName = FileDialogName
        gsOutputFilePath = FileDialogFolder
        gsOutputFileBase = Left$(gsOutputFileName, Len(gsOutputFileName) - 3&)
    End If
    '
    ' And the events file variables.
    gsOutputEventsPath = gsOutputFilePath
    If gbSeparateEventsFile Then
        gsOutputEventsBase = gsOutputFileBase & "_Events"
        gsOutputEventsName = gsOutputEventsBase & ".py"
        gsOutputEventsSpec = gsOutputEventsPath & gsOutputEventsName
        If gbOverwriteEvents = False Then
            gsOutputEventsAltSpec = UniqueFileSpec(gsOutputEventsSpec)
        Else
            gsOutputEventsAltSpec = gsOutputEventsSpec
        End If
    Else
        gsOutputEventsName = gsOutputFileName
        gsOutputEventsBase = gsOutputFileBase
        gsOutputEventsSpec = gsOutputFileSpec
        gsOutputEventsAltSpec = gsOutputEventsSpec
    End If
    '
    ' Good to go.
    GotPythonOutputFileSpec = True
End Function
    
Public Sub OpenPythonFile()
    '
    ' Kill any existing file.  We already said it was ok.
    If FileExists(gsOutputFileSpec) Then Kill gsOutputFileSpec
    '
    ghPy = FreeFile
    Open gsOutputFileSpec For Output As ghPy
End Sub
    
Public Sub OpenPythonEventsFile()
    '
    ' If it's not separate, just return and keep going.
    If gbSeparateEventsFile = False Then Exit Sub
    Close ghPy
    '
    ' We've already dealt with gbOverwriteEvents, so we can ignore it here.
    ' Kill any existing file.  We already said it was ok.
    If FileExists(gsOutputEventsAltSpec) Then Kill gsOutputEventsAltSpec
    ghPy = FreeFile
    Open gsOutputEventsAltSpec For Output As ghPy
End Sub

Public Sub ClosePythonFile()
    ' It may be the events file we're closing, if it's separate.
    If ghPy Then
        Close ghPy
        ghPy = 0&
    End If
End Sub
