Attribute VB_Name = "mod_Frm2Py_Write___MainForm"
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
        Print #ghPy, "        self.Vb6Class = 'Form'"
        Print #ghPy, "        self.setObjectName('"; .Name; "')"
        Print #ghPy, "        self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
        '
        ' Setup the form's properties from the VB6 FRM info.
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        # Form's initial properties, from VB6 FRM file."
        Print #ghPy, vbNullString
        ' Font.
        Dim sFont As String
        Print #ghPy, "        font = QFont('"; .Font.Name; "', "; CStr(CLng(.Font.Size)); ")"
        If .Font.Bold Then nop:           sFont = sFont & "font.setBold(True); "
        If .Font.Italic Then nop:         sFont = sFont & "font.setItalic(True); "
        If .Font.Underline Then nop:      sFont = sFont & "font.setUnderline(True); "
        If .Font.Strikethrough Then nop:  sFont = sFont & "font.setStrikeOut(True); "
      If Len(sFont) Then
        sFont = Left$(sFont, Len(sFont) - 2&) ' Clean it up.
        Print #ghPy, "        "; sFont
      End If
        Print #ghPy, "        self.setFont(font)"
        '
        ' The Tag property.
        Print #ghPy, "        self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        '
        ' Border style and width/height.
        Print #ghPy, "        self.__ClientWidth = "; CStr(.ClientWidth)
        Print #ghPy, "        self.__ClientHeight = "; CStr(.ClientHeight)
      Select Case .BorderStyle
      Case vbBSNone         ' resize sets interior.
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)"
        Print #ghPy, "        self.resize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.__FixedSize = False"
      Case vbSizable        ' resize sets interior.
        Print #ghPy, "        self.resize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.__FixedSize = False"
      Case vbFixedSingle    ' setFixedSize sets interior.
        Print #ghPy, "        self.setFixedSize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.__FixedSize = True"
      Case vbSizableToolWindow
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.Tool)"
        Print #ghPy, "        self.resize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.__FixedSize = False"
      Case vbFixedToolWindow
        Print #ghPy, "        self.setWindowFlags(self.windowFlags() | Qt.Tool)"
        Print #ghPy, "        self.setFixedSize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.__FixedSize = True"
      End Select
        '
        ' Startup position and left/top.
      Select Case .StartUpPosition
      Case vbStartUpManual
        Print #ghPy, "        self.__l = "; CStr(.ClientLeft - 4&) ' This works regardless of borderstyle because designer always has standard form decoration.
        Print #ghPy, "        self.__t = "; CStr(.ClientTop - 27&) ' This works regardless of borderstyle because designer always has standard form decoration.
      Case vbStartUpScreen
        Print #ghPy, "        centerPoint = QDesktopWidget().availableGeometry().center()"
        Print #ghPy, "        self.__l = centerPoint.x() - self.width() // 2"
        Print #ghPy, "        self.__t = centerPoint.y() - self.height() // 2"
      Case Else ' vbStartUpWindowsDefault
        Print #ghPy, "        self.__l = 0"
        Print #ghPy, "        self.__t = 0"
      End Select
        Print #ghPy, "        self.move(self.__l, self.__t)"
        '
        ' Form's icon.
      If Len(.Icon) Then
        Print #ghPy, "        self.__IconSpec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Icon; "')"
        Print #ghPy, "        self.setWindowIcon(QIcon(self.__IconSpec))"
      Else
        Print #ghPy, "        self.__IconSpec = None"
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
        ' Enabled & Visible.
        Print #ghPy, "        self.setEnabled("; TrueFalse(.Enabled); ")"
        ' Note on visibility.  This is just controlled with .Show() or .Hide(), so there's no need to address it here.
        'Print #ghPy, "        self.setVisible("; TrueFalse(.Visible); ")" ' We will still have a "Visible" property.
        '
        ' BackColor & ForeColor.
        Print #ghPy, "        self.setAutoFillBackground(True)"
        Print #ghPy, "        self.__BackColor = '"; RgbaHex(.BackColor); "'"
        ' Messing with a container's stylesheet fouls up its children, so we just don't do it.
        'Print #ghPy, "        self.__ForeColor = '"; RgbaHex(.ForeColor); "'"
        'Dim sStyle As String
        'sStyle = sStyle & "'color: ' + ToRgba(self.__ForeColor) + ';'"
        'Print #ghPy, "        self.setStyleSheet("; sStyle; ")"
        ' Do BackColor this way, to keep it from messing up any children.
        Print #ghPy, "        palette = self.palette()"
        Print #ghPy, "        palette.setColor(QPalette.Window, ToQColor(self.__BackColor))"
        Print #ghPy, "        self.setPalette(palette)"
        '
        ' Picture.  And we fit it to the form's client area in the paintEvent.
      If Len(.Picture) Then
        Print #ghPy, "        self.__ImageSpec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
        Print #ghPy, "        self.__BackPixmap = QPixmap(self.__ImageSpec)"
      Else
        Print #ghPy, "        self.__ImageSpec = ''"
        Print #ghPy, "        self.__BackPixmap = None"
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
        GetRidOfControlsNotProcessed
        SetWidgetTabOrders
        BuildWidgetsDictionary
        '
        ' Back to form's class, but out of __init__.
        '
        Print #ghPy, vbNullString
        Print #ghPy, "    # ****************************************************************"
        Print #ghPy, "    # We're still in our form's class, but no longer in __init__."
        Print #ghPy, "    # ****************************************************************"
        '
        ' Form's methods & properties.
        '
        ' Background picture.
        Print #ghPy, vbNullString
        Print #ghPy, "    # PNG (recommended), JPG, BMP, & SVG picture file formats are supported."
        Print #ghPy, "    # When in Windows OS, remember that '\' is escape character, so '\\' must be specified for back-slashes."
        Print #ghPy, "    # Or, the '/' also works when Python is running in Windows."
        Print #ghPy, "    # To see this in the return of PictureSpec, wrap return in repr(PictureSpec)"
        Print #ghPy, "    @property"
        Print #ghPy, "    def PictureSpec(self):"
        Print #ghPy, "        return self.__ImageSpec"
        Print #ghPy, "    @PictureSpec.setter"
        Print #ghPy, "    def PictureSpec(self, new_value: str):"
        Print #ghPy, "        if os.path.exists(new_value):"
        Print #ghPy, "            self.__ImageSpec = new_value"
        Print #ghPy, "            self.__BackPixmap = QPixmap(self.__ImageSpec)"
        Print #ghPy, "        else:"
        Print #ghPy, "            self.__ImageSpec = ''"
        Print #ghPy, "            self.__BackPixmap = None"
        Print #ghPy, "        self.repaint()"
        '
        ' BackColor.
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def BackColor(self):"
        Print #ghPy, "        return self.__BackColor"
        Print #ghPy, "    @BackColor.setter"
        Print #ghPy, "    def BackColor(self, new_hex_color: str):"
        Print #ghPy, "        if not isinstance(new_hex_color, (str)): return"
        Print #ghPy, "        if len(new_hex_color) == 7: new_hex_color += 'FF'"
        Print #ghPy, "        if not new_hex_color.startswith('#') or len(new_hex_color) != 9: return"
        Print #ghPy, "        self.__BackColor = new_hex_color"
        Print #ghPy, "        palette = self.palette()"
        Print #ghPy, "        palette.setColor(QPalette.Window, ToQColor(self.__BackColor))"
        Print #ghPy, "        self.setPalette(palette)"
        '
        ' Left, Top, ClientWidth, ClientHeight.
        Print #ghPy, vbNullString
        Print #ghPy, "    # Form's methods & properties (VB6 style).  For others, use PyQt members."
        Print #ghPy, vbNullString
        Print #ghPy, "    def Move(self, new_left: int, new_top: int, new_ClientWidth: int, new_ClientHeight: int):"
        Print #ghPy, "        self.__l = new_left"
        Print #ghPy, "        self.__t = new_top"
        Print #ghPy, "        self.__ClientWidth = new_ClientWidth"
        Print #ghPy, "        self.__ClientHeight = new_ClientHeight"
        Print #ghPy, "        if self.__FixedSize:"
        Print #ghPy, "            self.setFixedSize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        else:"
        Print #ghPy, "            self.resize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        self.move(self.__l, self.__t)"
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def Left(self):"
        Print #ghPy, "        return self.__l"
        Print #ghPy, "    @Left.setter"
        Print #ghPy, "    def Left(self, new_value: int):"
        Print #ghPy, "        self.__l = new_value"
        Print #ghPy, "        self.move(self.__l, self.__t)"
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def Top(self):"
        Print #ghPy, "        return self.__t"
        Print #ghPy, "    @Top.setter"
        Print #ghPy, "    def Top(self, new_value: int):"
        Print #ghPy, "        self.__t = new_value"
        Print #ghPy, "        self.move(self.__l, self.__t)"
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def ClientWidth(self):"
        Print #ghPy, "        return self.__ClientWidth"
        Print #ghPy, "    @ClientWidth.setter"
        Print #ghPy, "    def ClientWidth(self, new_value: int):"
        Print #ghPy, "        self.__ClientWidth = new_value"
        Print #ghPy, "        if self.__FixedSize:"
        Print #ghPy, "            self.setFixedSize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        else:"
        Print #ghPy, "            self.resize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def ClientHeight(self):"
        Print #ghPy, "        return self.__ClientHeight"
        Print #ghPy, "    @ClientHeight.setter"
        Print #ghPy, "    def ClientHeight(self, new_value: int):"
        Print #ghPy, "        self.__ClientHeight = new_value"
        Print #ghPy, "        if self.__FixedSize:"
        Print #ghPy, "            self.setFixedSize(self.__ClientWidth, self.__ClientHeight)"
        Print #ghPy, "        else:"
        Print #ghPy, "            self.resize(self.__ClientWidth, self.__ClientHeight)"
        '
        ' Font.
        Print #ghPy, vbNullString
        Print #ghPy, "    # Note that, for this main form, this font doesn't affect the caption, as that's controlled by the OS."
        Print #ghPy, "    @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "    def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "        return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        ' Visible.
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def Visible(self):"
        Print #ghPy, "        return self.isVisible()"
        Print #ghPy, "    @Visible.setter"
        Print #ghPy, "    def Visible(self, new_value: bool):"
        Print #ghPy, "        self.setVisible(new_value)"
        Print #ghPy, "        self.repaint()"
        ' Enabled.
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def Enabled(self):"
        Print #ghPy, "        return self.isEnabled()"
        Print #ghPy, "    @Enabled.setter"
        Print #ghPy, "    def Enabled(self, new_value: bool):"
        Print #ghPy, "        self.setEnabled(new_value)"
        ' Caption.
        Print #ghPy, vbNullString
        Print #ghPy, "    @property"
        Print #ghPy, "    def Caption(self):"
        Print #ghPy, "        return self.windowTitle()"
        Print #ghPy, "    @Caption.setter"
        Print #ghPy, "    def Caption(self, new_value: str):"
        Print #ghPy, "        self.setWindowTitle(new_value)"
        '
        ' Form's internal event procedures.
        '
        Print #ghPy, vbNullString
        Print #ghPy, "    # Internal event procedures.  They'll try and call external ones, if found."
        '
        ' closeEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def closeEvent(self, event): # event.ignore()  # Prevents the window from closing."
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Close' in globals(): "; .Name; "_Close(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' focusInEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def focusInEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_FocusIn' in globals(): "; .Name; "_FocusIn(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' focusOutEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def focusOutEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_FocusOut' in globals(): "; .Name; "_FocusOut(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' hideEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def hideEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Hide' in globals(): "; .Name; "_Hide(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' keyPressEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def keyPressEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_KeyPress' in globals(): "; .Name; "_KeyPress(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' keyReleaseEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def keyReleaseEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_KeyRelease' in globals(): "; .Name; "_KeyRelease(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' leaveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def leaveEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Leave' in globals(): "; .Name; "_Leave(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' mouseDoubleClickEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseDoubleClickEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_DoubleClick' in globals(): "; .Name; "_DoubleClick(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' mouseMoveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseMoveEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_MouseMove' in globals(): "; .Name; "_MouseMove(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' mousePressEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mousePressEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_MousePress' in globals(): "; .Name; "_MousePress(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' mouseReleaseEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def mouseReleaseEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_MouseRelease' in globals(): "; .Name; "_MouseRelease(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' moveEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def moveEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Move' in globals(): "; .Name; "_Move(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' paintEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def paintEvent(self, event):"
        Print #ghPy, "        super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
        Print #ghPy, "        if self.__BackPixmap: QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.__BackPixmap)"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Paint' in globals(): "; .Name; "_Paint(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        Print #ghPy, "        # Be sure to do .emit() after any picture, so lines get drawn on top of the picture."
        Print #ghPy, "        self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
        '
        ' resizeEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def resizeEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Resize' in globals(): "; .Name; "_Resize(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' showEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def showEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Show' in globals(): "; .Name; "_Show(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
        '
        ' wheelEvent
        Print #ghPy, vbNullString
        Print #ghPy, "    def wheelEvent(self, event):"
      If DoFormEvents Then
        Print #ghPy, "        if '"; .Name; "_Wheel' in globals(): "; .Name; "_Wheel(self, event)"
      Else
        Print #ghPy, "        pass"
      End If
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
    DoAllWidgetsClasses
    DoWidget_Arrays
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
                        Print #ghPy, "        self."; .Name; ".setEnabled("; TrueFalse(.Enabled); "}"
                        Print #ghPy, "        self."; .Name; ".setVisible("; TrueFalse(.Visible); "}"
                    End Select
                End If
            End If
        End With
    Next
End Sub
    
    
    
Public Sub WritePythonEventsCode()
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
        Print #ghPy, "# Delete them if you don't need them."
        Print #ghPy, "# ****************************************"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Close(self, event): # event.ignore()  # Prevents the window from closing."
        Print #ghPy, "    print('"; .Name; "_Close', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_DoubleClick(self, event):"
        Print #ghPy, "    print('"; .Name; "_DoubleClick', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_FocusIn(self, event):"
        Print #ghPy, "    print('"; .Name; "_FocusIn', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_FocusOut(self, event):"
        Print #ghPy, "    print('"; .Name; "_FocusOut', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Hide(self, event):"
        Print #ghPy, "    print('"; .Name; "_Hide', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_KeyPress(self, event):"
        Print #ghPy, "#    print('"; .Name; "_KeyPress', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_KeyRelease(self, event):"
        Print #ghPy, "#    print('"; .Name; "_KeyRelease', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Leave(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Leave', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MouseMove(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MouseMove', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MousePress(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MousePress', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_MouseRelease(self, event):"
        Print #ghPy, "#    print('"; .Name; "_MouseRelease', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Move(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Move', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Paint(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Paint', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Resize(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Resize', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "def "; .Name; "_Show(self, event):"
        Print #ghPy, "    print('"; .Name; "_Show', self.objectName())"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "# This one is noisy, so it's initially commented out."
        Print #ghPy, "#def "; .Name; "_Wheel(self, event):"
        Print #ghPy, "#    print('"; .Name; "_Wheel', self.objectName())"
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
