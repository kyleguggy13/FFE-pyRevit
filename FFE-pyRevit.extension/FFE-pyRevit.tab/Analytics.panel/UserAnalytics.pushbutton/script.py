# -*- coding: utf-8 -*-
__title__     = "User Analytics"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 10.22.2025
________________________________________________________________
Tested Revit Versions: 2026
______________________________________________________________
Description:

______________________________________________________________
How-to:
 -> Click the button

______________________________________________________________
Last update:
 - [10.22.2025] - v0.1 Beta Release

______________________________________________________________
Author: Kyle Guggenheim"""

import sys

#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document


output_window = output.get_output()



#____________________________________________________________________ FUNCTIONS
def md_list(items):
    """Render a Python list as one markdown string (never pass lists to print_md)."""
    try:
        lines = [u"- {}".format(u"{}".format(it)) for it in items]
    except Exception:
        lines = [u"- {}".format(str(it)) for it in items]
    return u"\n".join(lines)


def sanitize(v):
    """Return a friendly string for table cells."""
    if v is True:  return "True"
    if v is False: return "False"
    if v is None:  return "N/A"
    return str(v)




#____________________________________________________________________ MAIN
# Host data
doc_title = doc.Title


# Header
output_window.print_md("# Phase Filter Comparison")
output_window.print_md("## **Host model:** `{}`".format(doc_title))





#_____________________________________________________________________ 🏃‍➡️ RUN 


