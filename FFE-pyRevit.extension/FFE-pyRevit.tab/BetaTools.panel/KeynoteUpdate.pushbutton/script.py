# -*- coding: utf-8 -*-
__title__     = "Keynote Update"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.09.2025
# _____________________________________________________________________
# Description:
#   This script renames selected Family Types in Revit based on two parameters:
#   - "Number" and "Text"
# 
# _____________________________________________________________________
# How-to:
#
# -> Click the button
# -> Select the Family Types you want to rename
# -> The script will rename them to "Number Text" format
#   
# _____________________________________________________________________
# Last update:
# - [06.09.2025] - 1.0 RELEASE
# _____________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

import select
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, Transaction


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
output_window.set_title("Family Type Renamer")
output_window.print_md("## 🛠 Rename Selected Family Types")

# Customize these parameter names
PARAM_NAME_1 = "Number"
PARAM_NAME_2 = "Text"

def get_param_value(symbol, param_name):
    param = symbol.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or str(param.AsDouble())
    return "None"


def select_family_types():
    doc = revit.doc
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    symbols = list(collector)

    # Build display list
    symbol_options = ["{} : {}".format(s.FamilyName, s.TypeName) for s in symbols]
    selected = forms.SelectFromList.show(
        symbol_options,
        multiselect=True,
        title="Select Family Types to Rename"
        )
    
    if not selected: 
        forms.alert("No types selected.")
        return []
    
    # Get selected FamilySymbol objects back from names
    selected_symbols = []
    for s in symbols:
        label = "{} : {}".format(s.FamilyName, s.Name)
        if label in selected:
            selected_symbols.append(s)
    
    return selected_symbols


def rename_selected_types(symbols, param1_name, param2_name):
    doc = revit.doc
    renamed_types = []

    #________________________________________________________________ 🤖 Transaction
    with Transaction(doc, "Rename Family Types") as t:
        t.Start()
        for symbol in symbols:
            val1 = get_param_value(symbol, param1_name)
            val2 = get_param_value(symbol, param2_name)
            if val1 and val2:
                new_name = "{} {}".format(val1, val2)
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

#_____________________________________________________________________ 🏃‍➡️ RUN 
selected_symbols = select_family_types()
if selected_symbols:
    rename_selected_types(selected_symbols, PARAM_NAME_1, PARAM_NAME_2)

#==================================================
