# -*- coding: utf-8 -*-
__title__   = "FFE-pyRevit \nv1.6.2"
__doc__ = """Version = v1.6.2
Date    = 10.29.2025
_____________________________________________________________________
Description:
pyRevit About Form
_____________________________________________________________________
How-To:
- Click the Button
_____________________________________________________________________
Last update:
- [03.06.2025] - v1.0 RELEASE
- [03.25.2025] - v1.1 RELEASE
- [06.11.2025] - v1.2 RELEASE
- [07.25.2025] - v1.3 RELEASE
- [08.06.2025] - v1.4 RELEASE
- [09.14.2025] - v1.5 RELEASE
- [09.22.2025] - v1.5.1 Updated version numbering
- [10.08.2025] - v1.6.0 Added hooks
- [10.29.2025] - v1.6.2 Added Tutorials links
_____________________________________________________________________
Author: Kyle Guggenheim from FFE Inc."""


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
from Autodesk.Revit.DB import *
from pyrevit import forms   # By importing forms you also get references to WPF package! Very IMPORTANT
import wpf, os, clr         # wpf can be imported only after pyrevit.forms!

# .NET Imports
clr.AddReference("System")
from System.Collections.Generic import List
from System.Windows import Application, Window, ResourceDictionary
from System.Windows.Controls import CheckBox, Button, TextBox, ListBoxItem
from System.Diagnostics.Process import Start
from System.Windows.Window import DragMove
from System.Windows.Input import MouseButtonState
from System import Uri

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
PATH_SCRIPT = os.path.dirname(__file__)
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document #type: Document

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝
#====================================================================================================
class ListItem:
    """Helper Class for defining items in the ListBox."""
    def __init__(self,  Name='Unnamed', element = None, checked = False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element

    def __str__(self):
        return self.Name

# ╔╦╗╔═╗╦╔╗╔  ╔═╗╔═╗╦═╗╔╦╗
# ║║║╠═╣║║║║  ╠╣ ║ ║╠╦╝║║║
# ╩ ╩╩ ╩╩╝╚╝  ╚  ╚═╝╩╚═╩ ╩ MAIN FORM
#====================================================================================================
class AboutForm(Window):

    def __init__(self):
        # Connect to .xaml File (in same folder)
        path_xaml_file = os.path.join(PATH_SCRIPT, 'AboutUI.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # Show Form
        self.ShowDialog()

    # ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝
    #====================================================================================================
    def button_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()

    def header_drag(self, sender, e):
        """Drag window by holding LeftButton on the header."""
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def Hyperlink_RequestNavigate(self, sender, e):
        """Forwarding for a Hyperlinks."""
        Start(e.Uri.AbsoluteUri)



# ╦ ╦╔═╗╔═╗  ╔═╗╔═╗╦═╗╔╦╗
# ║ ║╚═╗║╣   ╠╣ ║ ║╠╦╝║║║
# ╚═╝╚═╝╚═╝  ╚  ╚═╝╩╚═╩ ╩
#====================================================================================================

# Show form to the user
UI = AboutForm()
