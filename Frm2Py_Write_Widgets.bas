Attribute VB_Name = "mod_Frm2Py_Write___Widgets"
Option Explicit
'

Public Sub DoWidgets_Class_Meth_Prop_IntEvt()
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
    Dim sStyle As String
    '
    ' Heavyweight ones first.
    For pCtl = 0& To UBound(guCtls)
        sStyle = vbNullString
        With guCtls(pCtl)
            Select Case .ClassName
                '
            Case "CommandButton"
                '
                ' The constructor.
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QPushButton): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
              If .Appearance = ccFlat Then
                sStyle = sStyle & "border: 1px solid black; "
              End If
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
                ' Deal with caption.
                Print #ghPy, "            self.caption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "', False)"
                Print #ghPy, "            self.caption.setGeometry(2, 2, w-4, h-4)"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Caption(self):"
                Print #ghPy, "            return self.caption.text()"
                Print #ghPy, "        @Caption.setter"
                Print #ghPy, "        def Caption(self, new_value: str):"
                Print #ghPy, "            self.caption.setText(new_value)"
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def clickEvent(self, event):"
                Print #ghPy, "            if '"; .Name; "_Click' in globals(): "; .Name; "_Click(self, event)"
                '
            Case "CheckBox"
                '
                ' The constructor.
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QCheckBox): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
                sStyle = sStyle & "border: 0px; "
                ' PyQt checkbox doesn't have a 3D style for the check indicator.
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
                ' Deal with caption.
                Print #ghPy, "            self.caption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "')"
                Print #ghPy, "            self.caption.setGeometry(16, 1, w-17, h-2)"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Caption(self):"
                Print #ghPy, "            return self.caption.text()"
                Print #ghPy, "        @Caption.setter"
                Print #ghPy, "        def Caption(self, new_value: str):"
                Print #ghPy, "            self.caption.setText(new_value)"
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Value(self): # 0=unchecked, 1=grayed, 2=checked."
                Print #ghPy, "            return self.checkState()"
                Print #ghPy, "        @Value.setter # 0=unchecked, 1=grayed, 2=checked."
                Print #ghPy, "        def Value(self, new_value: int):"
                Print #ghPy, "            self.caption.setCheckState(new_value)"
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def clickEvent(self, state):"
                Print #ghPy, "            if '"; .Name; "_Click' in globals(): "; .Name; "_Click(self, state)"
                '
            Case "OptionButton"
                '
                ' The constructor.
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QRadioButton): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
                sStyle = sStyle & "border: 0px; "
                ' PyQt checkbox doesn't have a 3D style for the check indicator.
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
                ' Deal with caption.
                Print #ghPy, "            self.caption = PassThruWrapLabel(self, '"; .Caption; "', Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "')"
                Print #ghPy, "            self.caption.setGeometry(16, 1, w-17, h-2)"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Caption(self):"
                Print #ghPy, "            return self.caption.text()"
                Print #ghPy, "        @Caption.setter"
                Print #ghPy, "        def Caption(self, new_value: str):"
                Print #ghPy, "            self.caption.setText(new_value)"
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Value(self): # 0=unchecked, 1=grayed, 2=checked."
                Print #ghPy, "            return self.isChecked()"
                Print #ghPy, "        @Value.setter # 0=unchecked, 1=grayed, 2=checked."
                Print #ghPy, "        def Value(self, new_value: bool):"
                Print #ghPy, "            self.caption.setChecked(new_value)"
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def clickEvent(self, state):"
                Print #ghPy, "            if '"; .Name; "_Click' in globals(): "; .Name; "_Click(self, state)"
                '
            Case "TextBox"
                '
              If .MultiLine Then ' MULTI_LINE TextBox.
                '
                Print #ghPy, vbNullString '                  /-- that's the multiline textbox.
                Print #ghPy, "    class cls"; .Name; "(QPlainTextEdit): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                Print #ghPy, "            self.setFont(font)"
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "QPlainTextEdit{background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & ";} "
                sStyle = sStyle & "QScrollBar:vertical{background-color: #F0F0F0;} "
                sStyle = sStyle & "QScrollBar:horizontal{background-color: #F0F0F0;} "
              Select Case True
              Case .BorderStyle = vbBSNone
                sStyle = sStyle & "QPlainTextEdit{border: 0px;} "
              Case .Appearance = ccFlat
                sStyle = sStyle & "QPlainTextEdit{border: 1px solid black;} "
              Case Else ' 3D border.
                sStyle = sStyle & "QPlainTextEdit{border: 2px inset gray;} "
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                ' Set style, enabled, visible.
                Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
                Print #ghPy, "            self.setEnabled("; TrueFalse(.Enabled); ")"
                Print #ghPy, "            self.setVisible("; TrueFalse(.Visible); ")"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
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
                Print #ghPy, vbNullString
                Print #ghPy, "        def changedEvent(self):"
                Print #ghPy, "            if '"; .Name; "_Change' in globals(): "; .Name; "_Change(self, self.toPlainText())"
                '
              Else ' SINGLE_LINE TextBox. <--------------------------------------------------------------------------------------------
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QLineEdit): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                Print #ghPy, "            self.setFont(font)"
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
              Select Case True
              Case .BorderStyle = vbBSNone
                sStyle = sStyle & "border: 0px; "
              Case .Appearance = ccFlat
                sStyle = sStyle & "border: 1px solid black; "
              Case Else ' 3D border.
                sStyle = sStyle & "border: 2px inset gray; "
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
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
                Print #ghPy, vbNullString
                Print #ghPy, "        def changedEvent(self, text):"
                Print #ghPy, "            if '"; .Name; "_Change' in globals(): "; .Name; "_Change(self, text)"
              End If ' Single or multi line.
                '
            Case "Frame" ' Container.
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
                Print #ghPy, "        paint_event_raised = pyqtSignal() # So we can 'emit()' our paintEvent to other widgets for this container."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
                sStyle = sStyle & "border: 0px; "
                sStyle = Trim$(sStyle)
                ' For a frame, we deal with the border in a paint event.
              Select Case True
              Case .BorderStyle = vbBSNone
                Print #ghPy, "            self.border = 0"
              Case .Appearance = ccFlat
                Print #ghPy, "            self.border = 1"
              Case Else ' 3D border.
                Print #ghPy, "            self.border = 2"
              End Select
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
                ' Deal with caption.
                Print #ghPy, "            caption_text = '"; .Caption; "'"
                Print #ghPy, "            font_metrics = QFontMetrics(font)"
                Print #ghPy, "            self.caption_width = font_metrics.horizontalAdvance(caption_text)"
                Print #ghPy, "            self.caption_height = font_metrics.height()"
                Print #ghPy, "            self.caption = PassThruWrapLabel(self, caption_text, Qt.AlignLeft | Qt.AlignVCenter, font, '"; RgbHex(.BackColor); "', '"; RgbHex(.ForeColor); "', False)"
                Print #ghPy, "            self.caption.setGeometry(6, 0, self.caption_width+1, self.caption_height)"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
                Print #ghPy, vbNullString
                Print #ghPy, "        @property"
                Print #ghPy, "        def Caption(self):"
                Print #ghPy, "            return self.caption.text()"
                Print #ghPy, "        @Caption.setter"
                Print #ghPy, "        def Caption(self, new_value: str):"
                Print #ghPy, "            font_metrics = QFontMetrics(self.caption.font())"
                Print #ghPy, "            self.caption_width = font_metrics.horizontalAdvance(new_value)"
                Print #ghPy, "            self.caption_height = font_metrics.height()"
                Print #ghPy, "            self.caption.setGeometry(6, 0, self.caption_width+1, self.caption_height)"
                Print #ghPy, "            self.caption.setText(new_value)"
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def paintEvent(self, event):"
                Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
                Print #ghPy, "            self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
                Print #ghPy, "            if self.border == 0: return # Nothing to do."
                Print #ghPy, "            if self.border == 1:"
                Print #ghPy, "                painter = QPainter(self)"
                Print #ghPy, "                painter.setBrush(QBrush(Qt.transparent))"
                Print #ghPy, "                painter.setPen(QPen(QColor('#000000'), 1))"
                Print #ghPy, "                painter.drawRect(0, self.caption_height//2, self.width()-1, self.height()-self.caption_height//2-1)"
                Print #ghPy, "                return"
                Print #ghPy, "            if self.border == 2:"
                Print #ghPy, "                painter = QPainter(self)"
                Print #ghPy, "                painter.setBrush(QBrush(Qt.transparent))"
                Print #ghPy, "                painter.setPen(QPen(QColor('#C0C0C0'), 2))" ' #C0C0C0 & #808080 is what 'border: 2px inset gray;' uses.
                Print #ghPy, "                painter.drawRect(1, self.caption_height//2+1, self.width()-2, self.height()-self.caption_height//2-2)"
                Print #ghPy, "                painter.setPen(QPen(QColor('#808080'), 1))"
                Print #ghPy, "                painter.drawRect(0, self.caption_height//2, self.width()-2, self.height()-self.caption_height//2-2)"
                Print #ghPy, "                return"
                '
            Case "PictureBox" ' Container.
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
                Print #ghPy, "        paint_event_raised = pyqtSignal() # So we can 'emit()' our paintEvent to other widgets for this container."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            self.RadioGroup = QButtonGroup(self) # Option button group for this container (VB6 style)."
                Print #ghPy, "            # Properties (from VB6)."
                ' Any picture.
              If Len(.Picture) Then
                Print #ghPy, "            self.image_spec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
                Print #ghPy, "            self.background_pixmap = QPixmap(self.image_spec)"
              Else
                Print #ghPy, "            self.image_spec = ''"
                Print #ghPy, "            self.background_pixmap = None"
              End If
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
              Select Case True
              Case .BorderStyle = vbBSNone
                sStyle = sStyle & "border: 0px; "
                Print #ghPy, "            self.border = 0"
              Case .Appearance = ccFlat
                sStyle = sStyle & "border: 1px solid black; "
                Print #ghPy, "            self.border = 1"
              Case Else ' 3D border.
                sStyle = sStyle & "border: 2px inset gray; "
                Print #ghPy, "            self.border = 2"
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                ' Python properties & methods, if any.
                '           A SavePicture might be nice.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def paintEvent(self, event):"
                Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
                Print #ghPy, "            self.paint_event_raised.emit() # This allows other widgets to 'see' this event, with binding."
                Print #ghPy, "            if self.background_pixmap: "
                Print #ghPy, "                if self.border == 0:"
                Print #ghPy, "                    QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.background_pixmap)"
                Print #ghPy, "                    return"
                Print #ghPy, "                if self.border == 1:"
                Print #ghPy, "                    QPainter(self).drawPixmap(1, 1, self.width()-2, self.height()-2, self.background_pixmap)"
                Print #ghPy, "                    return"
                Print #ghPy, "                if self.border == 2:"
                Print #ghPy, "                    QPainter(self).drawPixmap(2, 2, self.width()-4, self.height()-4, self.background_pixmap)"
                Print #ghPy, "                    return"
                '
            Case "ListBox"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QListWidget): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                Print #ghPy, "            self.setFont(font)"
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "QListWidget{background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & ";} "
                sStyle = sStyle & "QScrollBar:vertical{background-color: #F0F0F0;} "
                sStyle = sStyle & "QScrollBar:horizontal{background-color: #F0F0F0;} "
              Select Case True
              ' There is no option in VB6 to turn off the border on these.
              Case .Appearance = ccFlat
                sStyle = sStyle & "QListWidget{border: 1px solid black;} "
              Case Else ' 3D border.
                sStyle = sStyle & "QListWidget{border: 2px inset gray;} "
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                ' Python properties & methods, if any.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def clickEvent(self, item):"
                Print #ghPy, "            if '"; .Name; "_Click' in globals(): "; .Name; "_Click(self, item)"
                Print #ghPy, vbNullString
                Print #ghPy, "        def doubleclickEvent(self, item):"
                Print #ghPy, "            if '"; .Name; "_DblClick' in globals(): "; .Name; "_DblClick(self, item)"
                '
            Case "ComboBox"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QComboBox): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' An extra QLineEdit to serve as the QComboBox edit portion.
                Print #ghPy, "            line_edit = QLineEdit()"
                Print #ghPy, "            self.setLineEdit(line_edit)"
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                Print #ghPy, "            self.setFont(font)"
                ' There winds up being two stylesheets to get this correct.
                ' The one for the ComboBox itself, but don't change background or color, or bad things happen.
                sStyle = sStyle & "QComboBox {border: 0px;} "
                sStyle = sStyle & "QComboBox QAbstractItemView {background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & "; border: 1px solid black;}"
                Print #ghPy, "            self.setStyleSheet('"; sStyle; "')"
                ' Now we can work on the stylesheet for the edit box portion.
                sStyle = vbNullString
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; color: " & RgbHex(.ForeColor) & "; "
              Select Case True
              ' There is no option in VB6 to turn off the border on these.
              Case .Appearance = ccFlat
                sStyle = sStyle & "border: 1px solid black; "
              Case Else ' 3D border.
                sStyle = sStyle & "border: 2px inset gray; "
              End Select
                sStyle = Trim$(sStyle)
                Print #ghPy, "            line_edit.setStyleSheet('"; sStyle; "')"
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                Print #ghPy, "            line_edit.setReadOnly(False)"
              Else ' Treat it as true dropdown combo (ignoring "simple combo").
                Print #ghPy, "            line_edit.setReadOnly(True)"
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
                ' Python properties & methods, if any.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def changedEvent(self, text):"
                Print #ghPy, "            if '"; .Name; "_Change' in globals(): "; .Name; "_Change(self, text)"
            End Select
        End With
    Next
    '
    ' And now the lightweight ones.
    For pCtl = 0& To UBound(guCtls)
        sStyle = vbNullString
        With guCtls(pCtl)
            Select Case .ClassName
                '
            Case "Label"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QLabel): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Font.
                PrintWidgetFontLines .Font ' Just creates a font object.  Does NOT set the font on the widget.
                Print #ghPy, "            self.setFont(font)"
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                ' Experimented with several things, such as the following, but the QLabel just always gives more spacing than VB.Label.
                sStyle = sStyle & "margin-left: 1px; padding-left: -1px; margin-top: 0px; padding-top: -1px; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
              Select Case True
              Case .BackStyle = vbTransparent And .BorderStyle = vbBSNone   ' Transparent without border.
                sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
                sStyle = sStyle & "border: 0px; "
              Case .BackStyle = vbTransparent And .Appearance = ccFlat      ' Transparent with flat border.
                sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
                sStyle = sStyle & "border: 1px solid black; "
              Case .BackStyle = vbTransparent                               ' Transparent with 3D border.
                sStyle = sStyle & "background-color: rgba(0, 0, 0, 0)" & "; "
                sStyle = sStyle & "border: 2px inset gray; "
              Case .BorderStyle = vbBSNone                                  ' Opaque without border.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "border: 0px; "
              Case .Appearance = ccFlat                                     ' Opaque with flat border.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "border: 1px solid black; "
              Case Else                                                     ' Opaque with 3D border.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "border: 2px inset gray; "
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l-1, t, w+1, h) # Small adjustment needed, bug in PyQt?"
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
                ' Python properties & methods, if any.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Widget custom properties.  Use PyQt members for all others."
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
                '
            Case "Image"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(QFrame): # We're inheriting the widget's class."
                Print #ghPy, "        def __init__(self, container, form):"
                Print #ghPy, "            super().__init__(container) # Initialize the inherited object."
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Any picture.
              If Len(.Picture) Then
                Print #ghPy, "            self.image_spec = os.path.join(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Images'), '"; .Picture; "')"
                Print #ghPy, "            self.background_pixmap = QPixmap(self.image_spec)"
              Else
                Print #ghPy, "            self.image_spec = ''"
                Print #ghPy, "            self.background_pixmap = None"
              End If
                ' BackColor, ForeColor, & Flat or 3D ... via style sheet.
                sStyle = sStyle & "background-color: " & RgbHex(.BackColor) & "; "
                sStyle = sStyle & "color: " & RgbHex(.ForeColor) & "; "
              Select Case True
              Case .BorderStyle = vbBSNone
                sStyle = sStyle & "border: 0px; "
                Print #ghPy, "            self.border = 0"
              Case .Appearance = ccFlat
                sStyle = sStyle & "border: 1px solid black; "
                Print #ghPy, "            self.border = 1"
              Case Else ' 3D border.
                sStyle = sStyle & "border: 2px inset gray; "
                Print #ghPy, "            self.border = 2"
              End Select
                sStyle = Trim$(sStyle)
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            w = "; CStr(.Width); "; h = "; CStr(.Height); "; l = "; CStr(.Left); "; t = "; CStr(.Top)
                Print #ghPy, "            self.setGeometry(l, t, w, h)"
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
                ' Python properties & methods, if any.
                '           A SavePicture might be nice.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def paintEvent(self, event):"
                Print #ghPy, "            super().paintEvent(event) # Call the base class paintEvent to ensure default painting."
                Print #ghPy, "            if self.background_pixmap: "
                ' Note: If .Stretch is False, then the VB6 Image is already sized to fit the image,
                '       so, we don't have to worry about .Stretch at all.
                Print #ghPy, "                if self.border == 0:"
                Print #ghPy, "                    QPainter(self).drawPixmap(0, 0, self.width(), self.height(), self.background_pixmap)"
                Print #ghPy, "                    return"
                Print #ghPy, "                if self.border == 1:"
                Print #ghPy, "                    QPainter(self).drawPixmap(1, 1, self.width()-2, self.height()-2, self.background_pixmap)"
                Print #ghPy, "                    return"
                Print #ghPy, "                if self.border == 2:"
                Print #ghPy, "                    QPainter(self).drawPixmap(2, 2, self.width()-4, self.height()-4, self.background_pixmap)"
                Print #ghPy, "                    return"
                '
            Case "Line"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(): # This is one that doesn't inherit anything."
                Print #ghPy, "        def __init__(self, container, form):"
                ' No inheritance, so no: super().__init__(container)"
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
                ' Tag and geometry.
                Print #ghPy, "            self.Tag = '"; .Tag; "' # VB6 style 'TAG' property."
                Print #ghPy, "            self.x1 = "; CStr(.X1); "; self.y1 = "; CStr(.Y1); "; self.x2 = "; CStr(.X2); "; self.y2 = "; CStr(.Y2)
                ' Visible.
                Print #ghPy, "            self.Visible = "; TrueFalse(.Visible)
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
                ' Python properties & methods, if any.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def container_paint_event(self):"
                Print #ghPy, "            if self.Visible == False: return # Just don't draw it if it's invisible."
                Print #ghPy, "            painter = QPainter(self.Container)"
                Print #ghPy, "            pen = QPen(QColor(self.BorderColor))"
                Print #ghPy, "            pen.setWidth(self.BorderWidth)"
                Print #ghPy, "            pen.setStyle(self.BorderStyle)"
                Print #ghPy, "            painter.setPen(pen)"
                Print #ghPy, "            painter.drawLine(self.x1, self.y1, self.x2, self.y2)"
                '
            Case "Shape"
                '
                Print #ghPy, vbNullString
                Print #ghPy, "    class cls"; .Name; "(): # This is one that doesn't inherit anything."
                Print #ghPy, "        def __init__(self, container, form):"
                ' No inheritance, so no: super().__init__(container)"
                Print #ghPy, "            self.Name = '"; .Name; "'"
                Print #ghPy, "            self.Container = container"   ' Save our container object.
                Print #ghPy, "            self.Form = form"             ' Save our form object.
                Print #ghPy, "            # Properties (from VB6)."
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
                Print #ghPy, "            self.Width = "; CStr(.Width); "; self.Height = "; CStr(.Height); "; self.Left = "; CStr(.Left); "; self.Top = "; CStr(.Top)
                ' Visible.
                Print #ghPy, "            self.Visible = "; TrueFalse(.Visible)
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
                ' Python properties & methods, if any.
                '
                ' Internal events.
                Print #ghPy, vbNullString
                Print #ghPy, "        # Internal event(s) for widget."
                Print #ghPy, vbNullString
                Print #ghPy, "        def container_paint_event(self):"
                Print #ghPy, "            if self.Visible == False: return # Just don't draw it if it's invisible."
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
                Print #ghPy, "                painter.drawRect(self.Left, self.Top, self.Width, self.Height)"
                Print #ghPy, "            elif self.Shape == 2 or self.Shape == 3: # Oval or circle."
                Print #ghPy, "                painter.setRenderHint(QPainter.Antialiasing)"
                Print #ghPy, "                painter.drawEllipse(self.Left, self.Top, self.Width, self.Height)"
                Print #ghPy, "            else: # self.Shape == 4 or self.Shape == 5: # Rounded square or rounded rectangle."
                Print #ghPy, "                painter.setRenderHint(QPainter.Antialiasing)"
                Print #ghPy, "                rect = QRect(self.Left, self.Top, self.Width, self.Height)"
                Print #ghPy, "                painter.drawRoundedRect(rect, 20, 20, mode=Qt.RelativeSize)"
                '
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

Public Sub SetWidgetTabOrders()
    Dim i As Long, j As Long
    '
    ' First, get rid of widgets we didn't process, per instantiations.
    Dim pCtl As Long, iCtlsCount As Long
    pCtl = 0&
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
        GoTo GetOut
    End If
    ReDim Preserve guCtls(iCtlsCount - 1&)
    '
    ' And now, we will make a copy because not all widgets can participate in the tab order.
    '
    ' First count widgets that CAN participate in tab order, copying as we go.
    ' We also count how many >=0 TabIndex values there are, and also find TabIndex_Max.
    Dim iTabsCount As Long, iPosCount As Long, iTabMax As Long
    Dim uTabs() As TabsType
    ReDim uTabs(UBound(guCtls))
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
    If iTabsCount < 2& Then GoTo GetOut ' There's nothing to do with less than 2 widgets needing tab orders.
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
    '
GetOut: ' Just in case no controls nor tabbed controls were found.
    '
End Sub

' *****************************************************************************************
' *****************************************************************************************
' The following are called by ...Write_Main... after the _Events file is possibly created.
' *****************************************************************************************
' *****************************************************************************************

Public Sub DoExternalWidgetEventProcedures()
    ' We will loop through our control list in here.
    '
    Print #ghPy, vbNullString
    Print #ghPy, "# ************************************************"
    Print #ghPy, "# Widget event procedures for coding."
    Print #ghPy, "# ************************************************"
    '
    Dim pCtl As Long
    For pCtl = 0& To UBound(guCtls)
        With guCtls(pCtl)
            ' Just include ALL possible events (for any/all widgets).
            ' Which widgets get which calls is tested inside the AddExternalWidgetEventProc procedure.
            AddExternalWidgetEventProc .Name & "_Change", guCtls(pCtl)
            AddExternalWidgetEventProc .Name & "_Click", guCtls(pCtl)
            AddExternalWidgetEventProc .Name & "_DblClick", guCtls(pCtl)
        End With
    Next
End Sub

Private Sub AddExternalWidgetEventProc(sProcName As String, uCtrl As CtrlType)
    ' Just support for DoExternalWidgetEventProcedures
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
            Print #ghPy, "def " & sProcName & "(self, event):"
            Print #ghPy, "    print('" & sProcName & "', self.Name, self.Container.Name, self.Form.Name)"
        End Select
    Case "CheckBox"
        Select Case sSuffix
        Case "_Click"
            Print #ghPy, vbNullString
            Print #ghPy, "def " & sProcName & "(self, state):"
            Print #ghPy, "    print('" & sProcName & "', self.Name, self.Container.Name, self.Form.Name, 'State:', state)"
        End Select
    Case "OptionButton"
        Select Case sSuffix
        Case "_Click"
            Print #ghPy, vbNullString
            Print #ghPy, "def " & sProcName & "(self, state):"
            Print #ghPy, "    print('" & sProcName & "', self.Name, self.Container.Name, self.Form.Name, 'State:', state)"
        End Select
    Case "TextBox"
        Select Case sSuffix
        Case "_Change"
            Print #ghPy, vbNullString
            Print #ghPy, "def " & sProcName & "(self, text):"
            Print #ghPy, "    print('" & sProcName & "', self.Name, self.Container.Name, self.Form.Name, 'Text:', text)"
        End Select
    Case "Frame"            ' None, at this time.
    Case "PictureBox"       ' None, at this time.
    Case "ListBox"
        Select Case sSuffix
        Case "_Click", "_DblClick" ' Other than the procedure name, these are identical, at least the inserted stub.
            Print #ghPy, vbNullString
            Print #ghPy, "def " & sProcName & "(self, item):"
            Print #ghPy, "    print('" & sProcName & "', 'clicked:', item.text(), end=' selected: ')"
            Print #ghPy, "    selected_items = self.selectedItems()"
            Print #ghPy, "    for item in selected_items: print(item.text(), end=' ')"
            Print #ghPy, "    print('')"
        End Select
    Case "ComboBox"
        Select Case sSuffix
        Case "_Change"
            Print #ghPy, vbNullString
            Print #ghPy, "def " & sProcName & "(self, text):"
            Print #ghPy, "    print('" & sProcName & "', self.Name, self.Container.Name, self.Form.Name, 'Text:', text)"
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

