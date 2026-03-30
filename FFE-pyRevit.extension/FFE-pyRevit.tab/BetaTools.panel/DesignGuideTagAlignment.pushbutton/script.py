# -*- coding: utf-8 -*-
__title__     = "Design Guide Tag Alignment"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 03.30.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [03.30.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)



#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import ElementCategoryFilter, ElementId, FamilyInstance


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

log_status = ""
action = "Design Guide Tag Alignment"

output_window = output.get_output()
"""Output window for displaying results."""



#____________________________________________________________________ FUNCTIONS


