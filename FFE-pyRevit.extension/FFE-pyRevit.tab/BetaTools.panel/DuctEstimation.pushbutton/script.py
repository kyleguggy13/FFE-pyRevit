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
from Autodesk.Revit.DB import FilteredElementCollector, Mechanical, Transaction, Family, BuiltInParameter, ElementType, BuiltInCategory, ElementCategoryFilter, ElementId

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

selection = revit.get_selection()

# Set up the output panel
output_window = output.get_output()
output_window.set_title("Duct Estimation")
output_window.print_md("## 🛠 Duct & Insulation Estimation Tool")
output_window.print_md("### ⚠️ This tool estimates the total surface area and volume of sheet metal and insulation in the project.")

output_window.print_md("### 📋 Selected Elements:"
                       "\n- {}".format(len(selection.elements)))
output_window.print_md("### 📋 Selected Element IDs:"
                       "\n- {}".format(', '.join([str(el.Id) for el in selection.elements])))


# Test duct_collector
# duct_collector = FilteredElementCollector(doc).OfClass(ElementType).WhereElementIsElementType().ToElements()
# for duct_type in duct_collector:
#     if duct_type.FamilyName == "Round Duct" or duct_type.FamilyName == "Rectangular Duct":
#         ducttype_Name = DB.Element.Name.__get__(duct_type)
#         ducttype_ID = DB.Element.Id.__get__(duct_type)
#         output_window.print_md("### ✅ Found Duct Types: {}, {}".format(ducttype_Name, ducttype_ID))


# duct_instance_collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
# for duct_instance in duct_instance_collector:
#     if duct_instance.Symbol.FamilyName == "Round Duct" or duct_instance.Symbol.FamilyName == "Rectangular Duct":
#         ductinstance_Name = DB.Element.Name.__get__(duct_instance)
#         ductinstance_ID = DB.Element.Id.__get__(duct_instance)
#         output_window.print_md("### ✅ Found Duct Instances: {}, {}".format(ductinstance_Name, ductinstance_ID))


##################
### CURRENTLY WORKING:
# for duct in selection:
#     if isinstance(duct, FamilyInstance):
#         if duct.Symbol.FamilyName == "Round Duct" or duct.Symbol.FamilyName == "Rectangular Duct":
#             ductinstance_Name = DB.Element.Name.__get__(duct)
#             ductinstance_ID = DB.Element.Id.__get__(duct)
#             output_window.print_md("### ✅ Found Duct Instance: {}, {}".format(ductinstance_Name, ductinstance_ID))
#             break
##################

### Steps:
# 1. Select duct
# 2. Get "DuctNetwork" property
    # Property: DuctNetwork
    # Type: ElementSet
    # Full type: Autodesk.Revit.DB.ElementSet
    # Value: ElementSet
# This gives us the duct network, which contains Duct, DuctInsulation, and FamilyInstance elements.

# DuctNetwork = []
# for duct_instance in duct_instance_collector:
#     if duct_instance.Symbol.FamilyName == "Round Duct" or duct_instance.Symbol.FamilyName == "Rectangular Duct":
#         duct_network = duct_instance.MEPModel.DuctNetwork
#         if duct_network:
#             for element in duct_network:
#                 if element not in DuctNetwork:
#                     DuctNetwork.append(element)

# output_window.print_md("### ✅ Found Duct Network with {} elements.".format(len(DuctNetwork)))


### Revit Schedule Calculations (Reference)
# Density               = 0.294
# Gauge_Sq              = if(not((Width + Height) > 2'  6"), 24, if(not((Width + Height) > 4'  6"), 22, if(not((Width + Height) > 7'), 20, if((Width + Height) > 7', 18, 0))))
# Gauge_Thickness_Sq    = if(Gauge_Sq = 24, 0.028, if(Gauge_Sq = 22, 0.034, if(Gauge_Sq = 20, 0.04, if(Gauge_Sq = 18, 0.052, 0))))
# Weight_Sq             = Density * (2 * (Width + Height) / 0'  1") * Gauge_Thickness_Sq * (Length / 0'  1")
# SqFt_Sq               = (2 * (Width + Height) / 1') * (Length / 1')
# Gauge_Rnd             = if(not(Diameter > 2'  4"), 24, if(not(Diameter > 3'  2"), 22, if(not(Diameter > 4'), 20, if(Diameter > 4', 18, 0))))
# Gauge_Thickness_Rnd   = if(Gauge_Rnd = 24, 0.028, if(Gauge_Rnd = 22, 0.034, if(Gauge_Rnd = 20, 0.04, if(Gauge_Rnd = 18, 0.052, 0))))
# Weight_Rnd            = Density * (pi() * Diameter / 0'  1") * Gauge_Thickness_Rnd * (Length / 0'  1")
# SqFt_Rnd              = (pi() * Diameter / 1') * (Length / 1')

#_____________________________________________________________________ 🏃‍➡️ RUN 

# Collect all duct elements
duct_collector = FilteredElementCollector(doc)\
                    .OfCategory(BuiltInCategory.OST_DuctCurves)\
                    .WhereElementIsNotElementType()

# Print header
output_window.print_md("### Duct Elements Found")
output_window.print_table(
    table_data=[
        ["Id", "Type Name", "System Name", "Length (ft)", "Level"]
    ] + [
        [
            duct.Id.ToString(),
            duct.Name,
            duct.MEPSystem.Name if duct.MEPSystem else "N/A",
            "{:.2f}".format(duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()) if duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH) else "N/A",
            doc.GetElement(duct.LevelId).Name if duct.LevelId != ElementId.InvalidElementId else "N/A"
        ]
        for duct in duct_collector
    ],
    title="Ducts in Model"
)
#==================================================
