Attribute VB_Name = "mod_Frm2Py_Write___Widgets__Class_Specifics"
Option Explicit
'

Public Sub DoCommandButtonClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QPushButton): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font) # Also used for InternalCaption."
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
      If .Appearance = vbFlat Then
        sStyle = sStyle & "border: 1px solid black; "
      End If
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Deal with caption.
        Print #ghPy, "            self.InternalCaption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "', False)"
        Print #ghPy, "            self.InternalCaption.setGeometry(2, 2, self.__w-4, self.__h-4)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.clicked.connect(self.clickEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Caption(self):"
        Print #ghPy, "            return self.InternalCaption.text()"
        Print #ghPy, "        @Caption.setter"
        Print #ghPy, "        def Caption(self, new_value: str):"
        Print #ghPy, "            self.InternalCaption.setText(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def clickEvent(self, event):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self.Index, self, event)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self, event)"
      End If
    End With
End Sub



Public Sub DoCheckBoxClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QCheckBox): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font) # Also used for InternalCaption."
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
        sStyle = sStyle & "border: 0px; "
        ' PyQt checkbox doesn't have a 3D style for the check indicator.
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Deal with caption.
        Print #ghPy, "            self.InternalCaption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "')"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Initial value.
      Select Case .Value
      Case 2& ' Grayed.
        Print #ghPy, "            self.setTristate(True)"
        Print #ghPy, "            self.setCheckState(Qt.PartiallyChecked)"
      Case 0& ' Unchecked.
        Print #ghPy, "            self.setCheckState(Qt.Unchecked)"
      Case 1& ' Checked.
        Print #ghPy, "            self.setCheckState(Qt.Checked)"
      End Select
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.stateChanged.connect(self.clickEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Caption(self):"
        Print #ghPy, "            return self.InternalCaption.text()"
        Print #ghPy, "        @Caption.setter"
        Print #ghPy, "        def Caption(self, new_value: str):"
        Print #ghPy, "            self.InternalCaption.setText(new_value)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Value(self): # 0=unchecked, 1=grayed, 2=checked."
        Print #ghPy, "            return self.checkState()"
        Print #ghPy, "        @Value.setter # 0=unchecked, 1=grayed, 2=checked."
        Print #ghPy, "        def Value(self, new_value: int):"
        Print #ghPy, "            self.setCheckState(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def clickEvent(self, state):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self.Index, self, state)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self, state)"
      End If
    End With
End Sub



Public Sub DoOptionButtonClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QRadioButton): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font) # Also used for InternalCaption."
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
        sStyle = sStyle & "border: 0px; "
        ' PyQt checkbox doesn't have a 3D style for the check indicator.
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Deal with caption.
        Print #ghPy, "            self.InternalCaption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "')"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Initial value.
        Print #ghPy, "            self.setChecked("; TrueFalse(CBool(.Value)); ")"
        ' Grouping for radio button, done based on container copying VB6's scheme.
        ' We could have multiple groups per container though in Python.
        Print #ghPy, "            container.RadioGroup.addButton(self)"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.clicked.connect(self.clickEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.InternalCaption.setGeometry(16, 1, self.__w-17, self.__h-2)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Caption(self):"
        Print #ghPy, "            return self.InternalCaption.text()"
        Print #ghPy, "        @Caption.setter"
        Print #ghPy, "        def Caption(self, new_value: str):"
        Print #ghPy, "            self.InternalCaption.setText(new_value)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Value(self): # 0=unchecked, 1=grayed, 2=checked."
        Print #ghPy, "            return self.isChecked()"
        Print #ghPy, "        @Value.setter # 0=unchecked, 1=grayed, 2=checked."
        Print #ghPy, "        def Value(self, new_value: bool):"
        Print #ghPy, "            self.setChecked(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def clickEvent(self, state):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self.Index, self, state)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self, state)"
      End If
    End With
End Sub



Public Sub DoTextBoxMultiLineClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QPlainTextEdit): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "QPlainTextEdit{background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & ";} "
        sStyle = sStyle & "QScrollBar:vertical{background-color: #F0F0F0;} "
        sStyle = sStyle & "QScrollBar:horizontal{background-color: #F0F0F0;} "
      Select Case True
      Case .BorderStyle = vbBSNone
        sStyle = sStyle & "QPlainTextEdit{border: 0px;} "
      Case .Appearance = vbFlat
        sStyle = sStyle & "QPlainTextEdit{border: 1px solid black;} "
      Case Else ' 3D border.
        sStyle = sStyle & "QPlainTextEdit{border: 2px inset gray;} "
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Alignment. It's complicated for these QPlainTextEdit widgets, so we're going to skip it on a first pass.
        Print #ghPy, "            # We let alignment default to 'left' and may work on it more later."
        ' Scrollbars.
      If (.ScrollBars And 1&) = 1& Then ' Horizontal requested.
        Print #ghPy, "            self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn) # Qt.ScrollBarAsNeeded is another option."
        Print #ghPy, "            self.setLineWrapMode(QPlainTextEdit.NoWrap)"
      Else
        Print #ghPy, "            self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)"
        Print #ghPy, "            self.setLineWrapMode(QPlainTextEdit.WidgetWidth)"
      End If
      If (.ScrollBars And 2&) = 2& Then ' Vertical requested.
        Print #ghPy, "            self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn) # Qt.ScrollBarAsNeeded is another option."
      Else
        Print #ghPy, "            self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)"
      End If
        ' Set style, enabled, visible, locked.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        Print #ghPy, "            self.setReadOnly("; TrueFalse(.Locked); ")"
        ' Initial value.
        Print #ghPy, "            self.setPlainText("; FixMultiString(.Text, 25&); ")"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.textChanged.connect(self.changedEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Text(self):"
        Print #ghPy, "            return self.toPlainText()"
        Print #ghPy, "        @Text.setter"
        Print #ghPy, "        def Text(self, new_value: str):"
        Print #ghPy, "            self.setPlainText(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def changedEvent(self):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self.Index, self, self.toPlainText())"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self, self.toPlainText())"
      End If
    End With
End Sub



Public Sub DoTextBoxSingleLineClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QLineEdit): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
      Select Case True
      Case .BorderStyle = vbBSNone
        sStyle = sStyle & "border: 0px; "
      Case .Appearance = vbFlat
        sStyle = sStyle & "border: 1px solid black; "
      Case Else ' 3D border.
        sStyle = sStyle & "border: 2px inset gray; "
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Alignment. We always go vertical top, as that's what VB6 does.
      Select Case .Alignment
      Case vbRightJustify
        Print #ghPy, "            self.setAlignment(Qt.AlignRight | Qt.AlignTop)"
      Case vbCenter
        Print #ghPy, "            self.setAlignment(Qt.AlignHCenter | Qt.AlignTop)"
      Case Else ' Left justify.
        Print #ghPy, "            self.setAlignment(Qt.AlignLeft | Qt.AlignTop)"
      End Select
        ' Set style, enabled, visible, locked.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        Print #ghPy, "            self.setReadOnly("; TrueFalse(.Locked); ")"
        ' Initial value.
        Print #ghPy, "            self.setText('"; .Text; "')"
        Print #ghPy, "            self.setCursorPosition(0) # Make sure carat is all the way to the left."
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.textChanged.connect(self.changedEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Text(self):"
        Print #ghPy, "            return self.text()"
        Print #ghPy, "        @Text.setter"
        Print #ghPy, "        def Text(self, new_value: str):"
        Print #ghPy, "            self.setText(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def changedEvent(self, text):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self.Index, self, text)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self, text)"
      End If
    End With
End Sub



Public Sub DoFrameClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
        Print #ghPy, "        paint_event_raised = pyqtSignal() # So we can 'emit()' our paintEvent to other widgets for this container."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font) # Also used for InternalCaption."
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
        sStyle = sStyle & "border: 0px; "
        sStyle = Trim$(sStyle)
        ' For a frame, we deal with the border in a paint event.
      Select Case True
      Case .BorderStyle = vbBSNone
        Print #ghPy, "            self.__Border = 0"
      Case .Appearance = vbFlat
        Print #ghPy, "            self.__Border = 1"
      Case Else ' 3D border.
        Print #ghPy, "            self.__Border = 2"
      End Select
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Deal with caption.
        Print #ghPy, "            caption_text = '"; .Caption; "'"
        Print #ghPy, "            font_metrics = QFontMetrics(font)"
        Print #ghPy, "            caption_width = font_metrics.horizontalAdvance(caption_text)"
        Print #ghPy, "            caption_height = font_metrics.height()"
        Print #ghPy, "            self.InternalCaption = PassThruWrapLabel(self, caption_text, Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "', False)"
        Print #ghPy, "            self.InternalCaption.setGeometry(6, 0, caption_width+1, caption_height)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            # We could implement events similar to QMainWindow, but that's not presently done."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Caption(self):"
        Print #ghPy, "            return self.InternalCaption.text()"
        Print #ghPy, "        @Caption.setter"
        Print #ghPy, "        def Caption(self, new_value: str):"
        Print #ghPy, "            self.InternalCaption.setText(new_value)"
        Print #ghPy, "            self.repaint() # Needed so the InternalCaption widget gets resized and the border redrawn."
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def paintEvent(self, event):"
        Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
        Print #ghPy, "            self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
        Print #ghPy, "            font_metrics = QFontMetrics(self.InternalCaption.font())"
        Print #ghPy, "            caption_width = font_metrics.horizontalAdvance(self.InternalCaption.text())"
        Print #ghPy, "            caption_height = font_metrics.height()"
        Print #ghPy, "            self.InternalCaption.setGeometry(6, 0, caption_width+1, caption_height)"
        Print #ghPy, "            if self.__Border == 0:"
        Print #ghPy, "                return"
        Print #ghPy, "            if self.__Border == 1:"
        Print #ghPy, "                painter = QPainter(self)"
        Print #ghPy, "                painter.setBrush(QBrush(Qt.transparent))"
        Print #ghPy, "                painter.setPen(QPen(QColor('#000000'), 1))"
        Print #ghPy, "                painter.drawRect(0, caption_height//2, self.__w-1, self.__h-caption_height//2-1)"
        Print #ghPy, "                return"
        Print #ghPy, "            if self.__Border == 2:"
        Print #ghPy, "                painter = QPainter(self)"
        Print #ghPy, "                painter.setBrush(QBrush(Qt.transparent))"
        Print #ghPy, "                painter.setPen(QPen(QColor('#C0C0C0'), 2))" ' #C0C0C0 & #808080 is what 'border: 2px inset gray;' uses.
        Print #ghPy, "                painter.drawRect(1, caption_height//2+1, self.__w-2, self.__h-caption_height//2-2)"
        Print #ghPy, "                painter.setPen(QPen(QColor('#808080'), 1))"
        Print #ghPy, "                painter.drawRect(0, caption_height//2, self.__w-2, self.__h-caption_height//2-2)"
        Print #ghPy, "                return"
    End With
End Sub



Public Sub DoPictureBoxClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
        Print #ghPy, "        paint_event_raised = pyqtSignal() # So we can 'emit()' our paintEvent to other widgets for this container."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' Any picture.
      If Len(.Picture) Then
        Print #ghPy, "            self.__ImageSpec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
        Print #ghPy, "            self.__BackPixmap = QPixmap(self.__ImageSpec)"
      Else
        Print #ghPy, "            self.__ImageSpec = ''"
        Print #ghPy, "            self.__BackPixmap = None"
      End If
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
      Select Case True
      Case .BorderStyle = vbBSNone
        sStyle = sStyle & "border: 0px; "
        Print #ghPy, "            self.__Border = 0"
      Case .Appearance = vbFlat
        sStyle = sStyle & "border: 1px solid black; "
        Print #ghPy, "            self.__Border = 1"
      Case Else ' 3D border.
        sStyle = sStyle & "border: 2px inset gray; "
        Print #ghPy, "            self.__Border = 2"
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            # We could implement events similar to QMainWindow, but that's not presently done."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def paintEvent(self, event):"
        Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
        Print #ghPy, "            self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
        Print #ghPy, "            if self.__BackPixmap: "
        Print #ghPy, "                if self.__Border == 0:"
        Print #ghPy, "                    QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.__BackPixmap)"
        Print #ghPy, "                    return"
        Print #ghPy, "                if self.__Border == 1:"
        Print #ghPy, "                    QPainter(self).drawPixmap(1, 1, self.width()-2, self.height()-2, self.__BackPixmap)"
        Print #ghPy, "                    return"
        Print #ghPy, "                if self.__Border == 2:"
        Print #ghPy, "                    QPainter(self).drawPixmap(2, 2, self.width()-4, self.height()-4, self.__BackPixmap)"
        Print #ghPy, "                    return"
    End With
End Sub



Public Sub DoListBoxClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QListWidget): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "QListWidget{background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & ";} "
        sStyle = sStyle & "QScrollBar:vertical{background-color: #F0F0F0;} "
        sStyle = sStyle & "QScrollBar:horizontal{background-color: #F0F0F0;} "
      Select Case True
      ' There is no option in VB6 to turn off the border on these.
      Case .Appearance = vbFlat
        sStyle = sStyle & "QListWidget{border: 1px solid black;} "
      Case Else ' 3D border.
        sStyle = sStyle & "QListWidget{border: 2px inset gray;} "
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Sorting.
        Print #ghPy, "            self.setSortingEnabled("; TrueFalse(.Sorted); ")"
        ' Multiselect.
      Select Case .MultiSelect
      Case vbMultiSelectSimple
        Print #ghPy, "            self.setSelectionMode(QListWidget.MultiSelection)"
      Case vbMultiSelectExtended
        Print #ghPy, "            self.setSelectionMode(QListWidget.ExtendedSelection)"
      Case Else ' vbMultiSelectNone
        Print #ghPy, "            self.setSelectionMode(QListWidget.SingleSelection)"
      End Select
        ' Scrollbar policies.
        Print #ghPy, "            self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)"
        Print #ghPy, "            self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)"
        ' Initial values.
        Print #ghPy, "            items = "; PythonListFromFrxList(.List)
        Print #ghPy, "            for item in items: self.addItem(QListWidgetItem(item))"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.itemClicked.connect(self.clickEvent)"
        Print #ghPy, "            self.itemDoubleClicked.connect(self.doubleclickEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def clickEvent(self, item):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self.Index, self, item)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Click' in globals(): "; .OrigName; "_Click(self, item)"
      End If
        Print #ghPy, vbNullString
        Print #ghPy, "        def doubleclickEvent(self, item):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_DblClick' in globals(): "; .OrigName; "_DblClick(self.Index, self, item)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_DblClick' in globals(): "; .OrigName; "_DblClick(self, item)"
      End If
    End With
End Sub



Public Sub DoComboBoxClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QComboBox): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        ' An extra QLineEdit to serve as the QComboBox edit portion.
        Print #ghPy, "            self.InternalText = QLineEdit()"
        Print #ghPy, "            self.setLineEdit(self.InternalText)"
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font) # Also used for InternalText."
        Print #ghPy, "            self.InternalText.setFont(font)"
        ' There winds up being two stylesheets to get this correct.
        ' The one for the ComboBox itself, but don't change background or color, or bad things happen.
        Dim sStyle As String
        sStyle = sStyle & "QComboBox {border: 0px;} "
        sStyle = sStyle & "QComboBox QAbstractItemView {background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & "; border: 1px solid black;}"
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        ' Now we can work on the stylesheet for the edit box portion.
        sStyle = vbNullString
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & "; "
      Select Case True
      ' There is no option in VB6 to turn off the border on these.
      Case .Appearance = vbFlat
        sStyle = sStyle & "border: 1px solid black; "
      Case Else ' 3D border.
        sStyle = sStyle & "border: 2px inset gray; "
      End Select
        sStyle = Trim$(sStyle)
        Print #ghPy, "            self.InternalText.setStyleSheet('"; sStyle; "')"
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Set enabled, visible.
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Sorting.
      If .Sorted Then
        Print #ghPy, "            self.setInsertPolicy(QComboBox.InsertAlphabetically)"
      Else
        Print #ghPy, "            self.setInsertPolicy(QComboBox.NoInsert) # This is actually, insert-at-bottom"
      End If
        ' Is it a dropdown list or dropdown combo.
      If .Style = vbComboDropdownList Then
        Print #ghPy, "            self.InternalText.setReadOnly(False)"
      Else ' Treat it as true dropdown combo (ignoring "simple combo").
        Print #ghPy, "            self.InternalText.setReadOnly(True)"
      End If
        ' Initial values.
        Print #ghPy, "            items = "; PythonListFromFrxList(.List)
        Print #ghPy, "            for item in items: self.addItem(item)"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
      If .TabStop Then ' Just FYI, Python also has Qt.TabFocus(only) & Qt.NoFocus(not even click).
        Print #ghPy, "            self.setFocusPolicy(Qt.StrongFocus)"
      Else
        Print #ghPy, "            self.setFocusPolicy(Qt.ClickFocus)"
      End If
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.currentTextChanged.connect(self.changedEvent)"
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Text(self):"
        Print #ghPy, "            return self.currentText()"
        Print #ghPy, "        @Text.setter"
        Print #ghPy, "        def Text(self, new_value: str):"
        Print #ghPy, "            self.setCurrentText(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def changedEvent(self, text):"
      If .IsIndexed Then
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self.Index, self, text)"
      Else
        Print #ghPy, "            if '"; .OrigName; "_Change' in globals(): "; .OrigName; "_Change(self, text)"
      End If
    End With
End Sub



Public Sub DoLabelClass(uCtrl As CtrlType)
    With uCtrl
        '
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QLabel): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        ' Experimented with several things, such as the following, but the QLabel just always gives more spacing than VB.Label.
        Dim sStyle As String
        sStyle = sStyle & "margin-left: 1px; padding-left: -1px; margin-top: 0px; padding-top: -1px; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
      Select Case True
      Case .BackStyle = vbTransparent And .BorderStyle = vbBSNone   ' Transparent without border.
        sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
        sStyle = sStyle & "border: 0px; "
      Case .BackStyle = vbTransparent And .Appearance = vbFlat      ' Transparent with flat border.
        sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
        sStyle = sStyle & "border: 1px solid black; "
      Case .BackStyle = vbTransparent                               ' Transparent with 3D border.
        sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
        sStyle = sStyle & "border: 2px inset gray; "
      Case .BorderStyle = vbBSNone                                  ' Opaque without border.
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "border: 0px; "
      Case .Appearance = vbFlat                                     ' Opaque with flat border.
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "border: 1px solid black; "
      Case Else                                                     ' Opaque with 3D border.
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "border: 2px inset gray; "
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l-1, self.__t, self.__w+1, self.__h) # Small adjustment needed, bug in PyQt?"
        ' Alignment. We always go vertical top, as that's what VB6 does.
      Select Case .Alignment
      Case vbRightJustify
        Print #ghPy, "            self.setAlignment(Qt.AlignRight | Qt.AlignTop)"
      Case vbCenter
        Print #ghPy, "            self.setAlignment(Qt.AlignHCenter | Qt.AlignTop)"
      Case Else ' Left justify.
        Print #ghPy, "            self.setAlignment(Qt.AlignLeft | Qt.AlignTop)"
      End Select
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Caption.
        Print #ghPy, "            self.setWordWrap(True)"
        Print #ghPy, "            self.setText('"; .Caption; "')"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
        Print #ghPy, "            self.setFocusPolicy(Qt.NoFocus)" ' If we change this, we'll need to rework SetWidgetTabOrders procedure.
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            # None at the moment for the VB6 light-weight controls."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Caption(self):"
        Print #ghPy, "            return self.text()"
        Print #ghPy, "        @Caption.setter"
        Print #ghPy, "        def Caption(self, new_value: str):"
        Print #ghPy, "            self.setText(new_value)"
        '
        ' Internal events.
        ' None presently for light-weight controls.
    End With
End Sub



Public Sub DoImageClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
        Print #ghPy, "        def __init__(self, container, form):"
        Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Font.
        PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
        Print #ghPy, "            self.setFont(font)"
        ' Any picture.
      If Len(.Picture) Then
        Print #ghPy, "            self.__ImageSpec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
        Print #ghPy, "            self.__BackPixmap = QPixmap(self.__ImageSpec)"
      Else
        Print #ghPy, "            self.__ImageSpec = ''"
        Print #ghPy, "            self.__BackPixmap = None"
      End If
        ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
        Dim sStyle As String
        sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
        sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
      Select Case True
      Case .BorderStyle = vbBSNone
        sStyle = sStyle & "border: 0px; "
        Print #ghPy, "            self.__Border = 0"
      Case .Appearance = vbFlat
        sStyle = sStyle & "border: 1px solid black; "
        Print #ghPy, "            self.__Border = 1"
      Case Else ' 3D border.
        sStyle = sStyle & "border: 2px inset gray; "
        Print #ghPy, "            self.__Border = 2"
      End Select
        sStyle = Trim$(sStyle)
        ' Tag, Tooltip, and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.setToolTip('"; .ToolTipText; "') # These html tags work: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        ' Set style, enabled, visible.
        Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
        Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
        Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
        Print #ghPy, "            self.setFocusPolicy(Qt.NoFocus)" ' If we change this, we'll need to rework SetWidgetTabOrders procedure.
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            # None at the moment for the VB6 light-weight controls."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.setGeometry(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            self.repaint()"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property                   # No setter needed, as this is all handled by the clsVb6Font class."
        Print #ghPy, "        def Font(self):             # The return isn't meant to be saved as the widget stays attached to clsFont."
        Print #ghPy, "            return clsVb6Font(self) # Just use this to Get/Set the font's properties."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property # These html tags work in these tooltips: <b> <i> <u> <font> <br> <p> <a>, as well as \n for new lines."
        Print #ghPy, "        def ToolTipText(self):"
        Print #ghPy, "            return self.toolTip()"
        Print #ghPy, "        @ToolTipText.setter"
        Print #ghPy, "        def ToolTipText(self, new_value: str):"
        Print #ghPy, "            self.setToolTip(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.isVisible()"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            self.setVisible(new_value)"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.isEnabled()"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.setEnabled(new_value)"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def paintEvent(self, event):"
        Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
        Print #ghPy, "            if self.__BackPixmap: "
        ' Note: If .Stretch is False, then the VB6 Image is already sized to fit the image,
        '       so, we don't have to worry about .Stretch at all.
        Print #ghPy, "                if self.__Border == 0:"
        Print #ghPy, "                    QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.__BackPixmap)"
        Print #ghPy, "                    return"
        Print #ghPy, "                if self.__Border == 1:"
        Print #ghPy, "                    QPainter(self).drawPixmap(1, 1, self.width()-2, self.height()-2, self.__BackPixmap)"
        Print #ghPy, "                    return"
        Print #ghPy, "                if self.__Border == 2:"
        Print #ghPy, "                    QPainter(self).drawPixmap(2, 2, self.width()-4, self.height()-4, self.__BackPixmap)"
        Print #ghPy, "                    return"
    End With
End Sub



Public Sub DoLineClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(): # This is one that doesn't inherit anything."
        Print #ghPy, "        def __init__(self, container, form):"
        ' No inheritance, so no: super().__init__(container)"
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' Tag and geometry.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        Print #ghPy, "            self.__X1 = "; CStr(.X1); "; self.__Y1 = "; CStr(.Y1); "; self.__X2 = "; CStr(.X2); "; self.__Y2 = "; CStr(.Y2)
        ' Visible & enabled.
        Print #ghPy, "            self.__visible = "; TrueFalse(.Visible)
        Print #ghPy, "            self.__enabled = True # Just a dummy, lines have no actual enabled property."
        ' BorderColor.
        Print #ghPy, "            self.BorderColor = '"; RgbHex(.BorderColor); "'"
        ' BorderStyle.
      Select Case .BorderStyle ' 1=solid, 2=dash, 3=dot, 4=dash-dot, 5=dash-dot-dot.
      Case 2    ' Dash.
        Print #ghPy, "            self.BorderStyle = Qt.DashLine"
      Case 3    ' Dot.
        Print #ghPy, "            self.BorderStyle = Qt.DotLine"
      Case 4    ' Dash dot.
        Print #ghPy, "            self.BorderStyle = Qt.DashDotLine"
      Case 5    ' Dash dot dot.
        Print #ghPy, "            self.BorderStyle = Qt.DashDotDotLine"
      Case Else ' Solid.  We ignore the inside-solid option.
        Print #ghPy, "            self.BorderStyle = Qt.SolidLine"
      End Select
        ' BorderWidth.
        Print #ghPy, "            self.BorderWidth = "; CStr(.BorderWidth)
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
        ' No inheritance, so no: self.setFocusPolicy(Qt.NoFocus)"
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.Container.paint_event_raised.connect(self.container_paint_event) # Connect to container's paintEvent."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, "        def Move(self, new_X1: int, new_Y1: int, new_X2: int, new_Y2: int):"
        Print #ghPy, "            self.__X1 = new_X1"
        Print #ghPy, "            self.__Y1 = new_Y1"
        Print #ghPy, "            self.__X2 = new_X2"
        Print #ghPy, "            self.__Y2 = new_Y2"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def X1(self):"
        Print #ghPy, "            return self.__X1"
        Print #ghPy, "        @X1.setter"
        Print #ghPy, "        def X1(self, new_value: int):"
        Print #ghPy, "            self.__X1 = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Y1(self):"
        Print #ghPy, "            return self.__Y1"
        Print #ghPy, "        @Y1.setter"
        Print #ghPy, "        def Y1(self, new_value: int):"
        Print #ghPy, "            self.__Y1 = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def X2(self):"
        Print #ghPy, "            return self.__X2"
        Print #ghPy, "        @X2.setter"
        Print #ghPy, "        def X2(self, new_value: int):"
        Print #ghPy, "            self.__X2 = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Y2(self):"
        Print #ghPy, "            return self.__Y2"
        Print #ghPy, "        @Y2.setter"
        Print #ghPy, "        def Y2(self, new_value: int):"
        Print #ghPy, "            self.__Y2 = new_value"
        Print #ghPy, "            self.Container.repaint()"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.__visible"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            if self.__visible != new_value:"
        Print #ghPy, "                self.__visible = new_value"
        Print #ghPy, "                self.Container.repaint() # Drawn on container, so we must repaint it."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.__enabled"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.__enabled = new_value"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def container_paint_event(self):"
        Print #ghPy, "            if self.__visible == False: return # Just don't draw it if it's invisible."
        Print #ghPy, "            painter = QPainter(self.Container)"
        Print #ghPy, "            pen = QPen(QColor(self.BorderColor))"
        Print #ghPy, "            pen.setWidth(self.BorderWidth)"
        Print #ghPy, "            pen.setStyle(self.BorderStyle)"
        Print #ghPy, "            painter.setPen(pen)"
        Print #ghPy, "            painter.drawLine(self.__X1, self.__Y1, self.__X2, self.__Y2)"
    End With
End Sub



Public Sub DoShapeClass(uCtrl As CtrlType)
    With uCtrl
        Print #ghPy, vbNullString
        Print #ghPy, "    class cls"; .Name; "(): # This is one that doesn't inherit anything."
        Print #ghPy, "        def __init__(self, container, form):"
        ' No inheritance, so no: super().__init__(container)"
        Print #ghPy, "            self.Vb6Class = '"; .ClassName; "'"
        Print #ghPy, "            self.Name = '"; .Name; "'"
        Print #ghPy, "            self.Container = container"   ' Save our container object.
        Print #ghPy, "            self.Form = form"             ' Save our form object.
        Print #ghPy, "            # Properties (from VB6)."
        ' Control array stuff.
        Print #ghPy, "            self.IsIndexed = "; TrueFalse(.IsIndexed)
        Print #ghPy, "            self.Index = "; CStr(.Index)
        ' The type of shape.
        Print #ghPy, "            self.Shape = "; CStr(.Shape); " # 0=rect, 1=square, 2=oval, 3=circle, 4=rounded rect, 5=rounded square."
        ' BackColor & BackStyle.
        Print #ghPy, "            self.BackColor = '"; RgbHex(.BackColor); "'"
        Print #ghPy, "            self.BackOpaque = "; TrueFalse(CBool(.BackStyle))
        Print #ghPy, "            # As a note: VB6's FillColor/FillStyle are presently ignored."
        ' Tag.
        Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
        ' Geometry.  We may need to make some adjustments to accomodate circle & square.
        If .Shape = 1 Or .Shape = 3 Or .Shape = 5 Then
            Select Case True
            Case .Width > .Height
                .Left = .Left + (.Width - .Height) \ 2&
                .Width = .Height
            Case .Height > .Width
                .Top = .Top + (.Height - .Width) \ 2&
                .Height = .Width
            ' If they're equal, no adjustment needed.
            End Select
        End If
        Print #ghPy, "            self.__w = "; CStr(.Width); "; self.__h = "; CStr(.Height); "; self.__l = "; CStr(.Left); "; self.__t = "; CStr(.Top)
        ' Visible & enabled.
        Print #ghPy, "            self.__visible = "; TrueFalse(.Visible)
        Print #ghPy, "            self.__enabled = True # Just a dummy, shapes have no actual enabled property."
        ' BorderColor.
        Print #ghPy, "            self.BorderColor = '"; RgbHex(.BorderColor); "'"
        ' BorderStyle.
      Select Case .BorderStyle ' 1=solid, 2=dash, 3=dot, 4=dash-dot, 5=dash-dot-dot.
      Case 2    ' Dash.
        Print #ghPy, "            self.BorderStyle = Qt.DashLine"
      Case 3    ' Dot.
        Print #ghPy, "            self.BorderStyle = Qt.DotLine"
      Case 4    ' Dash dot.
        Print #ghPy, "            self.BorderStyle = Qt.DashDotLine"
      Case 5    ' Dash dot dot.
        Print #ghPy, "            self.BorderStyle = Qt.DashDotDotLine"
      Case Else ' Solid.  We ignore the inside-solid option.
        Print #ghPy, "            self.BorderStyle = Qt.SolidLine"
      End Select
        ' BorderWidth.
        Print #ghPy, "            self.BorderWidth = "; CStr(.BorderWidth)
        ' Focus policy (i.e., TabStop).  TabIndex is handled later.
        ' No inheritance, so no: self.setFocusPolicy(Qt.NoFocus)"
        ' Bindings.
        Print #ghPy, "            # Bindings."
        Print #ghPy, "            self.Container.paint_event_raised.connect(self.container_paint_event) # Connect to container's paintEvent."
        '
        ' Python properties & methods, VB6 style.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Widget custom properties (VB6 style).  Use PyQt members for all others."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def Move(self, new_left: int, new_top: int, new_width: int, new_height: int):"
        Print #ghPy, "            self.__l = new_left"
        Print #ghPy, "            self.__t = new_top"
        Print #ghPy, "            self.__w = new_width"
        Print #ghPy, "            self.__h = new_height"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Left(self):"
        Print #ghPy, "            return self.__l"
        Print #ghPy, "        @Left.setter"
        Print #ghPy, "        def Left(self, new_value: int):"
        Print #ghPy, "            self.__l = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Top(self):"
        Print #ghPy, "            return self.__t"
        Print #ghPy, "        @Top.setter"
        Print #ghPy, "        def Top(self, new_value: int):"
        Print #ghPy, "            self.__t = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Width(self):"
        Print #ghPy, "            return self.__w"
        Print #ghPy, "        @Width.setter"
        Print #ghPy, "        def Width(self, new_value: int):"
        Print #ghPy, "            self.__w = new_value"
        Print #ghPy, "            self.Container.repaint()"
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Height(self):"
        Print #ghPy, "            return self.__h"
        Print #ghPy, "        @Height.setter"
        Print #ghPy, "        def Height(self, new_value: int):"
        Print #ghPy, "            self.__h = new_value"
        Print #ghPy, "            self.Container.repaint()"
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Visible(self):"
        Print #ghPy, "            return self.__visible"
        Print #ghPy, "        @Visible.setter"
        Print #ghPy, "        def Visible(self, new_value: bool):"
        Print #ghPy, "            if self.__visible != new_value:"
        Print #ghPy, "                self.__visible = new_value"
        Print #ghPy, "                self.Container.repaint() # Drawn on container, so we must repaint it."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        @property"
        Print #ghPy, "        def Enabled(self):"
        Print #ghPy, "            return self.__enabled"
        Print #ghPy, "        @Enabled.setter"
        Print #ghPy, "        def Enabled(self, new_value: bool):"
        Print #ghPy, "            self.__enabled = new_value"
        '
        ' Internal events.
        Print #ghPy, vbNullString
        Print #ghPy, "        # Internal event(s) for widget."
        '
        Print #ghPy, vbNullString
        Print #ghPy, "        def container_paint_event(self):"
        Print #ghPy, "            if self.__visible == False: return # Just don't draw it if it's invisible."
        Print #ghPy, "            painter = QPainter(self.Container)"
        Print #ghPy, "            pen = QPen(QColor(self.BorderColor))"
        Print #ghPy, "            pen.setWidth(self.BorderWidth)"
        Print #ghPy, "            pen.setStyle(self.BorderStyle)"
        Print #ghPy, "            painter.setPen(pen)"
        Print #ghPy, "            if self.BackOpaque:"
        Print #ghPy, "                brush = QBrush(QColor(self.BackColor))"
        Print #ghPy, "            else:"
        Print #ghPy, "                brush = QBrush(Qt.transparent)"
        Print #ghPy, "            painter.setBrush(brush)"
        Print #ghPy, "            if self.Shape == 0 or self.Shape == 1: # Square or rectangle."
        Print #ghPy, "                painter.drawRect(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            elif self.Shape == 2 or self.Shape == 3: # Oval or circle."
        Print #ghPy, "                painter.setRenderHint(QPainter.Antialiasing)"
        Print #ghPy, "                painter.drawEllipse(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "            else: # self.Shape == 4 or self.Shape == 5: # Rounded square or rounded rectangle."
        Print #ghPy, "                painter.setRenderHint(QPainter.Antialiasing)"
        Print #ghPy, "                rect = QRect(self.__l, self.__t, self.__w, self.__h)"
        Print #ghPy, "                painter.drawRoundedRect(rect, 20, 20, mode=Qt.RelativeSize)"
    End With
End Sub

