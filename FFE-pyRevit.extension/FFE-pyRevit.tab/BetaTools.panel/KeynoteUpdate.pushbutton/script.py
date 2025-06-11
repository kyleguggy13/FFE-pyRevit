# -*- coding: utf-8 -*-
__title__     = "Keynote Update"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.09.2025
# ______________________________________________________________
# Description:
# -> This script renames selected Family Types 
#    in Revit based on two parameters:
# -> "Number" and "Text"
# 
# ______________________________________________________________
# How-to:
#
# -> Click the button
# -> Select the Family Types you want to rename
# -> The script will rename them to "Number Text" format
#   
# ______________________________________________________________
# Last update:
# - [06.09.2025] - 1.0 RELEASE
# ______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

import select
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, Transaction
from Autodesk.Revit.DB import BuiltInParameter


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
output_window.print_md("## 🛠 Rename Selected Annotation Family Types")

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
    # Filter only Generic Annotation symbols
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    symbols = [s for s in collector if s.Category and s.Category.Id == ElementId(BuiltInCategory.OST_GenericAnnotation)]

    # Build display list
    symbol_options = []
    symbol_lookup = {}

    for s in symbols:
        try:
            type_name = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            label = "{} : {}".format(s.FamilyName, type_name)
            symbol_options.append(label)
            symbol_lookup[label] = s
        except:
            continue

    selected = forms.SelectFromList.show(
        symbol_options,
        multiselect=True,
        title="Select Annotation Family Types to Rename"
    )

    if not selected: 
        forms.alert("No types selected.")
        return []

    return [symbol_lookup[label] for label in selected if label in symbol_lookup]



def rename_selected_types(symbols, param1_name, param2_name):
    doc = revit.doc
    renamed_types = []

    #________________________________________________________________ 🤖 Transaction
    with Transaction(doc, "Rename Selected Family Types") as t:
        t.Start()
        for symbol in symbols:
            val1 = get_param_value(symbol, param1_name)
            val2 = get_param_value(symbol, param2_name)
            if val1 and val2:
                new_name = "{} {}".format(val1, val2)
                old_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                try:
                    if old_name != new_name:
                        symbol.Name = new_name
                        renamed_types.append([symbol.FamilyName, old_name, new_name])
                except Exception as e:
                    output_window.print_md("### ❌ Error Renaming Type: {} : {}".format(old_name, str(e)))
        t.Commit()


    #________________________________________________________________ 📊 Output Results
    if renamed_types:
        output_window.print_md("### ✅ Renamed Types")
        output_window.print_table(
            table_data=renamed_types,
            title="Family Type Renames",
            columns=["Family", "Old Name", "New Name"]
        )
    else:
        output_window.print_md("### ⚠️ No Types Renamed")

#_____________________________________________________________________ 🏃‍➡️ RUN 
selected_symbols = select_family_types()
if selected_symbols:
    rename_selected_types(selected_symbols, PARAM_NAME_1, PARAM_NAME_2)

#==================================================
