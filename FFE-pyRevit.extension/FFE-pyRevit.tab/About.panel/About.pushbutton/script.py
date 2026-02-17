# -*- coding: utf-8 -*-
__title__   = "FFE-pyRevit \nv1.11.0"
__doc__ = """Version = v1.11.0
Date    = 02.11.2026
__________________________________________________________________
Description:
pyRevit About Form
__________________________________________________________________
How-To:
- Click the Button
__________________________________________________________________
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
- [11.04.2025] - v1.7.0 Added Project Info Comparison Tool
- [11.07.2025] - v1.8.0 Added Revision Comparison Tool
- [11.14.2025] - RevisionCompare: v1.8.1 Updated to fully support "Per Sheet" revision numbering.
- [12.15.2025] - v1.9.0 Added Design Guide Model Splitter Tool
- [02.02.2026] - v1.10.0 Added Duct Network Summary to MEPTools
- [02.09.2026] - v1.10.1 Updated Duct Network Summary to work on all duct system types.
- [02.11.2026] - v1.11.0 Added Sheets_Counter tool.
__________________________________________________________________
Author: Kyle Guggenheim from FFE Inc."""
# ↑ This module docstring doubles as the default changelog text.

#____________________________________________________________________ IMPORTS
from Autodesk.Revit.DB import *
from pyrevit import forms   # Import forms to bootstrap WPF in pyRevit
import wpf, os, clr

# .NET Imports
clr.AddReference("System")
from System.Collections.Generic import List
from System.Windows import Application, Window, ResourceDictionary, Clipboard
from System.Windows.Controls import CheckBox, Button, TextBox, ListBoxItem
from System.Diagnostics.Process import Start
from System.Windows.Window import DragMove
from System.Windows.Input import MouseButtonState
from System import Uri

#____________________________________________________________________ VARIABLES
PATH_SCRIPT = os.path.dirname(__file__)
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document  # type: Document

#____________________________________________________________________ CLASSES
class ListItem:
    """Helper Class for defining items in the ListBox."""
    def __init__(self,  Name='Unnamed', element=None, checked=False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element

    def __str__(self):
        return self.Name

#____________________________________________________________________ MAIN FORM
class AboutForm(Window):
    def __init__(self):
        # Connect to .xaml File (in same folder)
        path_xaml_file = os.path.join(PATH_SCRIPT, 'AboutUI.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # Wire up changelog area (safe if controls are not present)
        self._init_changelog()

        # Show Form
        self.ShowDialog()

    # ---------------------- Changelog helpers ----------------------
    def _get_changelog_text(self):
        """Returns the changelog text to display."""
        # 1) CHANGELOG.md next to the script (preferred if present)
        changelog_md = os.path.join(PATH_SCRIPT, "CHANGELOG.md")
        if os.path.exists(changelog_md):
            try:
                with open(changelog_md, "r") as fh:
                    return fh.read()
            except Exception as ex:
                forms.alert("Could not read CHANGELOG.md:\n{}".format(ex), warn_icon=True)

        # 2) Fallback to this module's __doc__ (the header already tracks releases)
        if __doc__:
            return __doc__

        # 3) Last resort
        return "No changelog available."

    def _init_changelog(self):
        """Populate TextBox and wire buttons if present in XAML."""
        try:
            tb = self.FindName("ChangelogTextBox")
            if tb:
                tb.Text = self._get_changelog_text()

        except Exception as ex:
            # Non-fatal: the rest of the window still works
            forms.alert("Changelog UI wiring skipped:\n{}".format(ex), warn_icon=True)

    # ---------------------- Window events ----------------------
    def button_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()

    def header_drag(self, sender, e):
        """Drag window by holding LeftButton on the header."""
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def Hyperlink_RequestNavigate(self, sender, e):
        """Forwarding for a Hyperlink."""
        Start(e.Uri.AbsoluteUri)

#____________________________________________________________________ USE FORM
# Show form to the user
UI = AboutForm()
