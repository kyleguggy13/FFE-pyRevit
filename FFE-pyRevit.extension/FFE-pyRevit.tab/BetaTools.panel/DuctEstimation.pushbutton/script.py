# -*- coding: utf-8 -*-
__title__     = "Duct Estimation"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 07.09.2025
# ______________________________________________________________
# Description:
# -> Estimates total surface area & volume of sheet metal.
#
# -> Estimates total surface area & volume of insulation.
# 
# ______________________________________________________________
# How-to:
#
# -> Click the button
#   
# ______________________________________________________________
# Last update:
# 
# ______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (AUTODESK)

import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, Transaction, Family, BuiltInParameter, ElementType

#____________________________________________________________________ IMPORTS (PYREVIT)

from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

#____________________________________________________________________ MAIN

# Set up the output panel
output_window = output.get_output()
output_window.set_title("Duct Estimation")
output_window.print_md("## 🛠 Duct & Insulation Estimation Tool")
output_window.print_md("### ⚠️ This tool estimates the total surface area and volume of sheet metal and insulation in the project.")


# Test arrowhead_collector
duct_collector = FilteredElementCollector(doc).OfClass(ElementType).WhereElementIsElementType().ToElements()
for duct_type in duct_collector:
    if duct_type.FamilyName == "Round Duct" or duct_type.FamilyName == "Rectangular Duct":
        ducttype_Name = DB.Element.Name.__get__(duct_type)
        ducttype_ID = DB.Element.Id.__get__(duct_type)
        output_window.print_md("### ✅ Found Duct Types: {}, {}".format(ducttype_Name, ducttype_ID))









#_____________________________________________________________________ 🏃‍➡️ RUN 


#==================================================
