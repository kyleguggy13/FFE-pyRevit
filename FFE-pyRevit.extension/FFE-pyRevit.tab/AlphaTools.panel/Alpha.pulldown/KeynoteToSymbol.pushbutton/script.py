# -*- coding: utf-8 -*-
__title__     = "Keynote To Symbol"
__version__   = 'Version = 0.1'
__doc__       = """Version = 0.1
Date    = 09.22.2025
# ______________________________________________________________
# Description:
# -> This script converts Revit Keynotes into Generic 
#    Annotation Symbols.
# -> 
# 
# ______________________________________________________________
# How-to:
#
# -> Click the button
# 
#   
# ______________________________________________________________
# Last update:
# - [09.22.2025] - 0.10 Initialized
# ______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (AUTODESK)
# from hmac import new
# from math import e
# import re
# from unittest import result
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, Transaction, Family, BuiltInParameter, ElementType
# from Autodesk.Revit.DB import BuiltInCategory, ElementId, Category, View, ViewFamilyType



#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


#____________________________________________________________________ TO-DO
# - [ ] Add error handling for missing parameters
# - [ ] Add option to select different arrowhead types
# - [ ] Check for duplicate numbers or texts
# - [ ] Add element links to the output table for easier identification
# - [ ] Add a confirmation dialog before renaming types 

#____________________________________________________________________ MAIN

### Set up the output panel
output_window = output.get_output()
output_window.set_title("Keynote To Symbol")
output_window.print_md("# 🛠 Convert FFE_Keynote_Hexagon ➡️ FFE_Symbol_Keynote")


### Keynote Parameter Names
PARAM_NAME_KEYNOTE_1 = "Key Value"
PARAM_NAME_KEYNOTE_2 = "Keynote Text"

### Customize these parameter names
PARAM_NAME_1 = "Number"
PARAM_NAME_2 = "Text"



def get_all_elements_of_category_in_view():
    # Get the current view
    view = doc.ActiveView
    # Collect all elements in the specified category that are visible in the current view
    collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_KeynoteTags).WhereElementIsNotElementType()
    elements = collector.ToElements()
    
    return elements

output_window.print_md("### Found {} Keynote elements in the current view.".format(len(get_all_elements_of_category_in_view())))
output_window.print_md("---")

# for elem in get_all_elements_of_category_in_view():
#     keynote_param_1 = elem.LookupParameter(PARAM_NAME_KEYNOTE_1)
#     keynote_param_2 = elem.LookupParameter(PARAM_NAME_KEYNOTE_2)
    
#     if keynote_param_1 and keynote_param_1.HasValue:
#         val1 = keynote_param_1.AsString()
#     else:
#         val1 = "None"
    
#     if keynote_param_2 and keynote_param_2.HasValue:
#         val2 = keynote_param_2.AsString()
#     else:
#         val2 = "None"


# Get unique keynote values from the elements
keynote_values = set()
for elem in get_all_elements_of_category_in_view():
    keynote_param_1 = elem.LookupParameter(PARAM_NAME_KEYNOTE_1)
    keynote_param_2 = elem.LookupParameter(PARAM_NAME_KEYNOTE_2)
    
    if keynote_param_1 and keynote_param_1.HasValue:
        val1 = keynote_param_1.AsString()
    else:
        val1 = "None"
    
    if keynote_param_2 and keynote_param_2.HasValue:
        val2 = keynote_param_2.AsString()
    else:
        val2 = "None"
    
    keynote_values.add((val1, val2))

keynote_values = list(keynote_values)


# Create new types in the FFE_Symbol_Keynote family
# and rename them based on the keynote values
# Set the arrowhead to "Arrow Filled 20 Degree"
# Ensure the family is loaded in the document
# Check if the family is loaded


### Collect Families
family_name = "FFE_Symbol_Keynote"
doc = revit.doc
collector = FilteredElementCollector(doc).OfClass(Family)

family_collector = []
for f in collector:
    if f.IsEditable and f.FamilyCategory and f.FamilyCategory.Id == ElementId(BuiltInCategory.OST_GenericAnnotation):
        output_window.print_md("f: {}".format(f.Name))
        family_collector.append(f)

family = next((f for f in family_collector if f.Name == family_name), None)

output_window.print_md("family name: {}".format(family.Name))
output_window.print_md("family id: {}".format(family.Id))

if not family:
    output_window.print_md("### ❌ Family '{}' not found in the document.".format(family_name))
    output_window.print_md("### Please load the family and try again.")
    sys.exit()





# Create family types based on unique keynote values
def create_types_in_family(family, keynote_values):
    # doc = revit.doc
    created_types = []
    with Transaction(doc, "Create Keynote Symbol Types") as t:
        t.Start()
        for val1, val2 in keynote_values:
            # Create a new type name based on the keynote values
            new_type_name = "{} {}".format(val1, val2)

            keynote_param_1 = val1
            keynote_param_2 = val2
            
            output_window.print_md("keynote_param_1: {}".format(keynote_param_1))
            output_window.print_md("keynote_param_2: {}".format(keynote_param_2))
            output_window.print_md("---")

            # Check if the type already exists
            existing_type = None
            for sym_id in family.GetFamilySymbolIds():
                symbol = doc.GetElement(sym_id)
                symbol_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if symbol_name == new_type_name:
                    existing_type = symbol
                    break

            
            
            if not existing_type:
                # Create a new family symbol (type)
                try:
                    # Duplicate the first available symbol in the family
                    symbol_ids = family.GetFamilySymbolIds()
                    output_window.print_md("symbol_ids: {}".format(symbol_ids))

                    if symbol_ids is None or symbol_ids.Count == 0:
                        raise Exception("No FamilySymbols found in family '{}'.".format(family.Name))
                    
                    base_symbol_id = next(iter(symbol_ids))
                    base_symbol = doc.GetElement(base_symbol_id)
                    output_window.print_md("Base Symbol: {}".format(base_symbol))

                    # Duplicate the base symbol to create a new type
                    new_symbol_id = base_symbol.Duplicate(new_type_name)
                    output_window.print_md("new_symbol_id: {}".format(new_symbol_id))

                    new_symbol = doc.GetElement(new_symbol_id)
                    output_window.print_md("### ✅ Created Type: {}".format(new_symbol))

                    # Set the symbol parameters
                    param_number = new_symbol.LookupParameter(PARAM_NAME_1)                    
                    param_text = new_symbol.LookupParameter(PARAM_NAME_2)
                    
                    param_number.Set(keynote_param_1)
                    param_text.Set(keynote_param_2)

                    
                    created_types.append([family.Name, new_type_name])
                except Exception as e:
                    output_window.print_md("### ❌ Error Creating Type: {} : {}".format(new_type_name, str(e)))
        
        t.Commit()

    if created_types:
        output_window.print_md("### ✅ Created Types")
        output_window.print_table(
            table_data=created_types,
            title="Family Type Renames",
            columns=["Family", "Name"]
        )
    else:
        output_window.print_md("### ⚠️ No Types Created")


#_____________________________________________________________________ 🏃‍➡️ RUN 
create_types_in_family(family, keynote_values)





"""
### Function to get parameter value from a symbol
### Returns the value as a string or "None" if not set
def get_param_value(symbol, param_name):
    param = symbol.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or str(param.AsDouble())
    return "None"


### Function to select annotation families
### Returns a list of selected Family objects
def select_annotation_families():
    doc = revit.doc
    collector = FilteredElementCollector(doc).OfClass(Family)

    annotation_families = []
    for f in collector:
        if f.IsEditable and f.FamilyCategory and f.FamilyCategory.Id == ElementId(BuiltInCategory.OST_GenericAnnotation):
            annotation_families.append(f)

    family_options = [f.Name for f in annotation_families]
    selected = forms.SelectFromList.show(
        family_options,
        multiselect=True,
        title="Select Annotation Families (This will set the Leader Arrowhead)",
    )

    if not selected:
        forms.alert("No families selected.")
        return []

    return [f for f in annotation_families if f.Name in selected]



### Function to rename types in selected families
### It renames each type based on the values of the specified parameters
def rename_types_in_families(families, param1_name, param2_name):
    doc = revit.doc
    renamed_types = []

    with Transaction(doc, "Rename Types in Selected Families") as t:
        t.Start()
        for family in families:
            symbols = family.GetFamilySymbolIds()
            for sym_id in symbols:
                symbol = doc.GetElement(sym_id)

                set_leader_arrowhead(symbol)  # Set the leader arrowhead
                
                val1 = get_param_value(symbol, param1_name)  # Get value of first parameter
                val2 = get_param_value(symbol, param2_name)  # Get value of second parameter
                
                if val1 and val2:
                    new_name = "{} {}".format(val1, val2)
                    old_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    try:
                        if old_name != new_name:
                            symbol.Name = new_name  # Set the new name
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
# """
