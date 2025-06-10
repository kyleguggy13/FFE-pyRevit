# -*- coding: utf-8 -*-
__title__     = "Keynote Update"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.09.2025
# _____________________________________________________________________
# Description:
#
# 
# _____________________________________________________________________
# How-to:
#
# -> Click the button
# -> 
# _____________________________________________________________________
# Last update:
# - [06.09.2025] - 1.0 RELEASE
# _____________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

from logging import Filter
from requests import get
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, Transaction


#____________________________________________________________________ IMPORTS (PYREVIT)

from pyrevit import revit, DB
from pyrevit.script import output


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


#____________________________________________________________________ MAIN


# Set up the output panel
output_window = output.get_output()
output_window.set_title("Family Type Renamer")
output_window.print_md("## 🛠 Renaming Family Types Based on Parameters")

# Customize these parameter names
PARAM_NAME_1 = "Width"
PARAM_NAME_2 = "Height"

def get_param_value(symbol, param_name):
    param = symbol.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or str(param.AsDouble())
    return "None"

def rename_family_types(param1_name, param2_name):
    doc = revit.doc
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)

    renamed_types = []

    #________________________________________________________________ 🤖 Transaction
    with Transaction(doc, "Rename Family Types") as t:
        t.Start()
        for symbol in collector:
            val1 = get_param_value(symbol, param1_name)
            val2 = get_param_value(symbol, param2_name)
            if val1 and val2:
                new_name = "{} x {}".format(val1, val2)
                old_name = symbol.Name
                try:
                    if old_name != new_name:
                        symbol.Name = new_name
                        renamed_types.append([symbol.FamilyName, old_name, new_name])
                except Exception as e:
                    output.print_md("### ❌ Error Renaming Type: {} - {}".format(old_name, str(e)))
        t.Commit()


    #________________________________________________________________ 📊 Output Results
    if renamed_types:
        output.print_md("### ✅ Renamed Types")
        output.print_table(
            table_data=renamed_types,
            title="Family Type Renames",
            columns=["Family", "Old Name", "New Name"]
        )
    else:
        output.print_md("### ⚠️ No Types Renamed")

# Run it
rename_family_types(PARAM_NAME_1, PARAM_NAME_2)




#==================================================
