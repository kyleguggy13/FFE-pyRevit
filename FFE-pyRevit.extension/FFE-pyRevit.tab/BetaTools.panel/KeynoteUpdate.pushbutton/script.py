# -*- coding: utf-8 -*-
__title__     = "Keynote Update"
__version__   = 'Version = 1.2'
__doc__       = """Version = 1.2
Date    = 06.11.2025
# ______________________________________________________________
# Description:
# -> This script renames all Family Types within selected
#    Annotation Families based on two parameters:
# -> "Number" and "Text"
# -> It also sets the Leader Arrowhead to "Arrow Filled 20 Degree"
# 
# ______________________________________________________________
# How-to:
#
# -> Click the button
# -> Select the Annotation Families you want to rename
# -> All Family Types within those families will be 
#    renamed to "Number Text" format
# -> Leader Arrowhead will be set to "Arrow Filled 20 Degree"
#   
# ______________________________________________________________
# Last update:
# - [06.11.2025] - 1.2 Added Leader Arrowhead setting
# ______________________________________________________________
Inspiration: Marc Padros
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
output_window.set_title("Family Type Renamer")
output_window.print_md("## 🛠 Rename Types in Selected Annotation Families")
output_window.print_md("### ⚠️ Must have two parameters: **Number** and **Text**")


# Test arrowhead_collector
arrowhead_collector = FilteredElementCollector(doc).OfClass(ElementType).WhereElementIsElementType().ToElements()
for arrowhead_type in arrowhead_collector:
    if arrowhead_type.FamilyName == "Arrowhead":
        Arrowheadtype_Name = DB.Element.Name.__get__(arrowhead_type)
        Arrowheadtype_ID = DB.Element.Id.__get__(arrowhead_type)
        output_window.print_md("### ✅ Found Arrowhead Type: {}, {}".format(Arrowheadtype_Name, Arrowheadtype_ID))


# Customize these parameter names
PARAM_NAME_1 = "Number"
PARAM_NAME_2 = "Text"


# Function to get parameter value from a symbol
# Returns the value as a string or "None" if not set
def get_param_value(symbol, param_name):
    param = symbol.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or str(param.AsDouble())
    return "None"


# Function to select annotation families
# Returns a list of selected Family objects
def select_annotation_families():
    doc = revit.doc
    collector = FilteredElementCollector(doc).OfClass(Family)
    annotation_families = [f for f in collector if f.IsEditable and f.FamilyCategory and f.FamilyCategory.Id == ElementId(BuiltInCategory.OST_GenericAnnotation)]

    family_options = [f.Name for f in annotation_families]
    selected = forms.SelectFromList.show(
        family_options,
        multiselect=True,
        title="Select Annotation Families"
    )

    if not selected:
        forms.alert("No families selected.")
        return []

    return [f for f in annotation_families if f.Name in selected]


# Function to set the leader arrowhead for a symbol
# It sets the arrowhead to "Arrow Filled 20 Degree"
def set_leader_arrowhead(symbol):
    doc = revit.doc
    arrowhead_name = "Arrow Filled 20 Degree"
    
    # Find the arrowhead type in the document
    arrowhead_collector = FilteredElementCollector(doc).OfClass(ElementType).WhereElementIsElementType().ToElements()
    arrowhead_type = next((a for a in arrowhead_collector if a.Name == arrowhead_name), None)

    if not arrowhead_type:
        output_window.print_md("### ❌ Arrowhead Type Not Found: {}".format(arrowhead_name))
        return

    # Set the leader arrowhead type
    try:
        symbol.get_Parameter(BuiltInParameter.LEADER_ARROWHEAD_TYPE).Set(arrowhead_type.Id)
        output_window.print_md("### ✅ Leader Arrowhead Set to: {}".format(arrowhead_name))
    except Exception as e:
        output_window.print_md("### ❌ Error Setting Leader Arrowhead: {}".format(str(e)))


# Function to rename types in selected families
# It renames each type based on the values of the specified parameters
def rename_types_in_families(families, param1_name, param2_name):
    doc = revit.doc
    renamed_types = []

    with Transaction(doc, "Rename Types in Selected Families") as t:
        t.Start()
        for family in families:
            symbols = family.GetFamilySymbolIds()
            for sym_id in symbols:
                symbol = doc.GetElement(sym_id)
                # set_leader_arrowhead(symbol)  # Set the leader arrowhead
                val1 = get_param_value(symbol, param1_name)
                val2 = get_param_value(symbol, param2_name)
                if val1 and val2:
                    new_name = "{} {}".format(val1, val2)
                    old_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    try:
                        if old_name != new_name:
                            symbol.Name = new_name
                            renamed_types.append([family.Name, old_name, new_name])
                            
                    except Exception as e:
                        output_window.print_md("### ❌ Error Renaming Type: {} : {}".format(old_name, str(e)))
        t.Commit()

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
selected_families = select_annotation_families()
if selected_families:
    rename_types_in_families(selected_families, PARAM_NAME_1, PARAM_NAME_2)

#==================================================
