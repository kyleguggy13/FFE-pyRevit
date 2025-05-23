# -*- coding: utf-8 -*-
__title__   = "CC Collections"
__doc__ = """Version = 1.0
Date    = 03.07.2025
_____________________________________________________________________
Description:
Autodesk Content Catalog Browser
_____________________________________________________________________
How-To:
- Click the Button
_____________________________________________________________________
Last update:
- [03.07.2025] - V1.0 RELEASE
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
        path_xaml_file = os.path.join(PATH_SCRIPT, 'TutorialsUI.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # Show Form
        self.ShowDialog()


    # ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ Events
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
# ╚═╝╚═╝╚═╝  ╚  ╚═╝╩╚═╩ ╩ Use Form
#====================================================================================================
# Show form to the user
UI = AboutForm()
