# -*- coding: utf-8 -*-
__title__     = "Create Sheet Sets"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 04.21.2025
# ___________________________________________________________________
# Description:

# This will create sheet sets that can be used for exporting to PDF and for publishing to ACC.
# ___________________________________________________________________
# How-to:

# -> Click on the button
# -> Select All or By Discipline
# ___________________________________________________________________
# Last update:
# - [05.22.2025] - 1.0 RELEASE
# ___________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

from operator import mul
from math import exp
import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms


#____________________________________________________________________ VARIABLES
# app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


#____________________________________________________________________ MAIN

# 1️⃣ Get Sheets

# Get all sheets
collector = FilteredElementCollector(doc).OfClass(ViewSheet)
all_sheets = [sheet for sheet in collector if sheet.GetParameters("Appears In Sheet List") and 
              sheet.LookupParameter("Appears In Sheet List").AsInteger() == 1]


export_options = ["ALL SHEETS", "GENERAL", "STRUCTURAL", "ARCHITECTURE", "PLUMBING", "MECHANICAL", "ELECTRICAL"]

# Select ALL SHEETS or BY DISCIPLINE
select_group = forms.SelectFromList.show(export_options, multiselect=False, button_name="Select Group")
export_group = select_group
print("Selected group: ", select_group, "Type: ", type(select_group))
print("Selected group: ", export_group, "Type: ", type(export_group))


# Check if the user selected "ALL SHEETS" or a specific discipline
if export_group == "ALL SHEETS":
    # Select all sheets
    all_sheets = [sheet for sheet in collector if sheet.GetParameters("Appears In Sheet List") and 
                  sheet.LookupParameter("Appears In Sheet List").AsInteger() == 1]
    print("ALL SHEETS: ", len(all_sheets))
else:
    # Select sheets by discipline
    all_sheets = [sheet for sheet in collector if sheet.GetParameters("Appears In Sheet List") and 
                  sheet.LookupParameter("Appears In Sheet List").AsInteger() == 1 and 
                  sheet.LookupParameter("FFE_Sheet_Discipline").AsString() == export_group]
    print("Sheets selected by discipline: ", len(all_sheets))





# Create a new PrintSet
print_manager = doc.PrintManager
print_manager.PrintRange = PrintRange.Select
view_set = ViewSet()


# Add all eligible sheets to the ViewSet
for sheet in all_sheets:
    view_set.Insert(sheet)


# Create the print set
print_manager.ViewSetToUse = view_set




# # Set name for the print set
# with Transaction(doc, "Create Sheet Set") as t:
#     t.Start() # 🔓

#     try:
#         printSetName = export_group
#         viewSheetSetting = doc.PrintManager.ViewSheetSetting
#         viewSheetSetting.CurrentViewSheetSet.Views = view_set
#         viewSheetSetting.SaveAs(printSetName)
#         t.Commit()
#         print("Success", "Sheet set '{p}' created with {a} sheets".format(p=printSetName, a=len(all_sheets)))
#         # TaskDialog.Show("Success", f"Sheet set '{printSetName}' created with {len(all_sheets)} sheets")
#     except Exception as e:
#         t.RollBack()
#         print("Error ", "Failed to create sheet set: {e}".format(e=str(e)))
#         # TaskDialog.Show("Error", f"Failed to create sheet set: {str(e)}")

#🤖 Automate Your Boring Work Here



#_____________________________________________________________________ 🏃‍➡️ RUN 

# # Set name for the print set
transaction = Transaction(doc, "Create Sheet Set")
transaction.Start()
try:
    printSetName = "ALL SHEETS"
    viewSheetSetting = doc.PrintManager.ViewSheetSetting
    viewSheetSetting.CurrentViewSheetSet.Views = view_set
    viewSheetSetting.SaveAs(printSetName)
    transaction.Commit()
    print("Success", "Sheet set '{p}' created with {a} sheets".format(p=printSetName, a=len(all_sheets)))
    # TaskDialog.Show("Success", f"Sheet set '{printSetName}' created with {len(all_sheets)} sheets")
except Exception as e:
    transaction.RollBack()
    print("Error ", "Failed to create sheet set: {e}".format(e=str(e)))
    # TaskDialog.Show("Error", f"Failed to create sheet set: {str(e)}")

#🤖 Automate Your Boring Work Here


#==================================================
