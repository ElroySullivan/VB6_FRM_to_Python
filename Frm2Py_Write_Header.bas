Attribute VB_Name = "mod_Frm2Py_Write___Header_Etc"
Option Explicit
'

Public Sub WritePythonHeader()
    Print #ghPy, vbNullString
    Print #ghPy, "# Initially created with "; App.Title; " written by Elroy Sullivan, PhD."
  If gbSeparateEventsFile Then
    Print #ghPy, vbNullString
    Print #ghPy, "from "; gsOutputEventsBase; " import *"
  End If
    Print #ghPy, vbNullString
    Print #ghPy, "import sys"
    Print #ghPy, "import os"
    Print #ghPy, "from PyQt5.QtCore import Qt, QRect, pyqtSignal"
    Print #ghPy, "from PyQt5.QtGui import QIcon, QPixmap, QPalette, QColor, QTextCursor, QTextBlockFormat"
    Print #ghPy, "from PyQt5.QtGui import QFont, QFontMetrics, QPainter, QPen, QBrush"
    Print #ghPy, "from PyQt5.QtWidgets import QMainWindow, QAction, QDesktopWidget"
    Print #ghPy, "from PyQt5.QtWidgets import QLabel, QPushButton, QCheckBox, QRadioButton, QButtonGroup"
    Print #ghPy, "from PyQt5.QtWidgets import QFrame, QListWidget, QListWidgetItem, QComboBox"
    Print #ghPy, "#                          single line   multi line    like RTB"
    Print #ghPy, "from PyQt5.QtWidgets import QLineEdit, QPlainTextEdit, QTextEdit"
  If Not gbSeparateEventsFile Then
    Print #ghPy, "from PyQt5.QtWidgets import QApplication"
  End If
End Sub

Public Sub WriteModuleLevelProcsAndClasses()
    Print #ghPy, vbNullString
    Print #ghPy, vbNullString
    Print #ghPy, "# *******************************"
    Print #ghPy, "# We're now back at MODULE level."
    Print #ghPy, "# *******************************"
    Print #ghPy, vbNullString
    Print #ghPy, "# ************************************************"
    Print #ghPy, "# Some needed general procedures & classes."
    Print #ghPy, "# ************************************************"
    Print #ghPy, vbNullString
    Print #ghPy, "class PassThruWrapLabel(QLabel): # Handles alignment, wordwrap, font, backcolor, & forecolor."
    Print #ghPy, "    def __init__(self, parent, text, alignment, font, backcolor, forecolor, DoClickSpoof=True):"
    Print #ghPy, "        super().__init__(parent)"
    Print #ghPy, "        self.parent = parent"
    Print #ghPy, "        if DoClickSpoof: self.mouseReleaseEvent = self.ClickSpoof"
    Print #ghPy, "        self.setWordWrap(True)"
    Print #ghPy, "        self.setAlignment(alignment)"
    Print #ghPy, "        self.setFont(font)"
    Print #ghPy, "        self.setStyleSheet('background-color: ' + backcolor + '; color: ' + forecolor + '; border: 0px;')"
    Print #ghPy, "        self.setText(text)"
    Print #ghPy, "    #"
    Print #ghPy, "    def mousePressEvent(self, event):   self.parent.mousePressEvent(event)"
    Print #ghPy, "    def mouseReleaseEvent(self, event): self.parent.mouseReleaseEvent(event)"
    Print #ghPy, "    def mouseMoveEvent(self, event):    self.parent.mouseMoveEvent(event)"
    Print #ghPy, "    def wheelEvent(self, event):        self.parent.wheelEvent(event)"
    Print #ghPy, "    def keyPressEvent(self, event):     self.parent.keyPressEvent(event)"
    Print #ghPy, "    def keyReleaseEvent(self, event):   self.parent.keyReleaseEvent(event)"
    Print #ghPy, "    def focusInEvent(self, event):      self.parent.focusInEvent(event)"
    Print #ghPy, "    def focusOutEvent(self, event):     self.parent.focusOutEvent(event)"
    Print #ghPy, "    #"
    Print #ghPy, "    def ClickSpoof(self, event):   "
    Print #ghPy, "        if self.rect().contains(event.pos()):"
    Print #ghPy, "            self.parent.click()"
    Print #ghPy, vbNullString
    Print #ghPy, "def IntFromStrColor(strColor: str):"
    Print #ghPy, "    # This converts a '#rrggbb' string into an 0x00bbggrr integer."
    Print #ghPy, "    # The input string MUST be in '#rrggbb' or results will be unpredictable."
    Print #ghPy, "    r = int(strColor[1:3], 16)"
    Print #ghPy, "    g = int(strColor[3:5], 16)"
    Print #ghPy, "    b = int(strColor[5:7], 16)"
    Print #ghPy, "    return (b << 16) + (g << 8) + r"
    Print #ghPy, vbNullString
    Print #ghPy, "def StrColorFromInt(intColor: int):"
    Print #ghPy, "    # This converts an 0x00bbggrr integer into a '#rrggbb' string."
    Print #ghPy, "    # The eight high-order bits of the integer will be ignored."
    Print #ghPy, "    r = hex(intColor & 0xff)[2:].zfill(2)"
    Print #ghPy, "    g = hex(intColor >> 8 & 0xff)[2:].zfill(2)"
    Print #ghPy, "    b = hex(intColor >> 16 & 0xff)[2:].zfill(2)"
    Print #ghPy, "    return '#' + r + g + b"
    Print #ghPy, vbNullString
    Print #ghPy, "class clsVb6Font(): "
    Print #ghPy, "    # Used to make a QFont 'look like' a VB6 font."
    Print #ghPy, "    # Returned by the above clsWidgets from their Font property."
    Print #ghPy, "    # It should just be used to get/set the properties in this class, and not directly used."
    Print #ghPy, "    # The widget stays associated with this font object."
    Print #ghPy, "    # If the widget for the font is some nested widget, that must be what's passed in during initialization."
    Print #ghPy, "    def __init__(self, widget):  "
    Print #ghPy, "        self.widget = widget"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Name(self):                             "
    Print #ghPy, "        return self.widget.font().family()"
    Print #ghPy, "    @Name.setter"
    Print #ghPy, "    def Name(self, new_name: int):              "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setFamily(new_name)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Size(self):                             "
    Print #ghPy, "        return self.widget.font().pointSize()"
    Print #ghPy, "    @Size.setter"
    Print #ghPy, "    def Size(self, new_size: int):              "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setPointSize(new_size)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Bold(self):                             "
    Print #ghPy, "        return self.widget.font().bold()"
    Print #ghPy, "    @Bold.setter"
    Print #ghPy, "    def Bold(self, new_bold: bool):             "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setBold(new_bold)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Italic(self):                           "
    Print #ghPy, "        return self.widget.font().italic()"
    Print #ghPy, "    @Italic.setter"
    Print #ghPy, "    def Bold(self, new_italic: bool):           "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setItalic(new_italic)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Underline(self):                        "
    Print #ghPy, "        return self.widget.font().underline()"
    Print #ghPy, "    @Bold.setter"
    Print #ghPy, "    def Underline(self, new_underline: bool):   "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setUnderline(new_underline)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    @property"
    Print #ghPy, "    def Strikeout(self):                        "
    Print #ghPy, "        return self.widget.font().strikeOut()"
    Print #ghPy, "    @Strikeout.setter"
    Print #ghPy, "    def Strikeout(self, new_strikeout: bool):   "
    Print #ghPy, "        font = self.widget.font()"
    Print #ghPy, "        font.setStrikeOut(new_strikeout)"
    Print #ghPy, "        self.widget.setFont(font)"
    Print #ghPy, "        self.CheckAndSetSubFont(font)"
    Print #ghPy, "    def CheckAndSetSubFont(self, font):"
    Print #ghPy, "        if self.widget.Vb6Class == 'CommandButton':"
    Print #ghPy, "            self.widget.InternalCaption.setFont(font)"
    Print #ghPy, "            return"
    Print #ghPy, "        if self.widget.Vb6Class == 'CheckBox':"
    Print #ghPy, "            self.widget.InternalCaption.setFont(font)"
    Print #ghPy, "            return"
    Print #ghPy, "        if self.widget.Vb6Class == 'OptionButton':"
    Print #ghPy, "            self.widget.InternalCaption.setFont(font)"
    Print #ghPy, "            return"
    Print #ghPy, "        if self.widget.Vb6Class == 'ComboBox':"
    Print #ghPy, "            self.widget.InternalText.setFont(font)"
    Print #ghPy, "            return"
    Print #ghPy, "        if self.widget.Vb6Class == 'Frame':"
    Print #ghPy, "            self.widget.InternalCaption.setFont(font)"
    Print #ghPy, "            self.widget.repaint() # Necessary for frame to redraw caption and border."
    Print #ghPy, "            return"
End Sub

Public Sub WriteTestingCode()
    '
    ' Write out a test to run it if we're the "__main__".
    Print #ghPy, vbNullString
    Print #ghPy, vbNullString
    Print #ghPy, "# ******************************************************************"
    Print #ghPy, "# Some test code in case we want to directly run and test this file."
    Print #ghPy, "# ******************************************************************"
    Print #ghPy, vbNullString
    Print #ghPy, "if __name__ == '__main__':"
    '
    Print #ghPy, "    app = QApplication(sys.argv) # Get an application object."
    Print #ghPy, "    "; guForm.Name; " = cls"; guForm.Name; "() # Instantiate our form.  It inherits QMainWindow."
    Print #ghPy, "    "; guForm.Name; ".show() # Show our form."
    Print #ghPy, "    sys.exit(app.exec_()) # Loop waiting on events to be raised."
    '
    Print #ghPy, vbNullString
    Print #ghPy, vbNullString
    Print #ghPy, vbNullString
End Sub
