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
# -> Review existing print sets
# -> Select SHEETS or a DISCIPLINE
# -> Review the output in the console
#
# -> NOTE: Only sheets with the "Appears In Sheet List" parameter checked on will be included in the sheet set.
# ___________________________________________________________________
# Last update:
# - [05.22.2025] - 0.1 BETA RELEASE
# - [07.25.2025] - 1.0 RELEASE
# ___________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import framework
from pyrevit import revit, script, DB, UI
from pyrevit import forms


#____________________________________________________________________ VARIABLES
# app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


#____________________________________________________________________ MAIN

# 1️⃣ Get Sheets 🚫
def GetSheets():
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
    return all_sheets



### REFER TO: 
### "C:\Users\kyleg\AppData\Roaming\pyRevit-Master\extensions\pyRevitTools.extension\pyRevit.tab\Drawing Set.panel\views.stack\Views.pulldown\Create Print Set From Selected Views.pushbutton"
### "C:\Users\kyleg\AppData\Roaming\pyRevit-Master\extensions\pyRevitTools.extension\pyRevit.tab\Drawing Set.panel\Sheets.pulldown\List TitleBlocks on Sheets.pushbutton\script.py"


output = script.get_output()

### 1️⃣ Get PrintManager / ViewSheetSetting
print_manager = revit.doc.PrintManager
print_manager.PrintRange = DB.PrintRange.Select
viewsheetsetting = print_manager.ViewSheetSetting


### 2️⃣ Collect existing ViewSheetSets (List)
print_sets_existing = DB.FilteredElementCollector(revit.doc)\
    .WhereElementIsNotElementType().OfClass(DB.ViewSheetSet).ToElements()


output.print_md("### EXISTING PRINT SETS:")
print_sets_names_existing = []
for vs in print_sets_existing:
    vs_name = vs.Name
    vs_viewcount = len(vs.OrderedViewList)

    output.print_md("{vs_name} ({vs_viewcount})".format(vs_name=vs_name, vs_viewcount=vs_viewcount))
    
    print_sets_names_existing.append(vs_name)


output.print_md("---")


### 3️⃣ Collect Sheets with "Appears In Sheet List" parameter
SheetCollector = FilteredElementCollector(doc).OfClass(ViewSheet)
all_sheets = []

for sheet in SheetCollector:
    if sheet.GetParameters("Appears In Sheet List") and sheet.LookupParameter("Appears In Sheet List").AsInteger() == 1:
        all_sheets.append(sheet)
        discipline = sheet.LookupParameter("FFE_Sheet_Discipline").AsString()
        index = sheet.LookupParameter("FFE_Sheet_Discipline Index").AsString()
        order = sheet.LookupParameter("FFE_Sheet_Order").AsString()
        SheetNumber = sheet.SheetNumber
        SheetName = sheet.Name


### 4️⃣ Create a dictionary to store sheets by discipline
Disciplines = {}
Disciplines["SHEETS"] = all_sheets
for sheet in all_sheets:
    discipline = sheet.LookupParameter("FFE_Sheet_Discipline").AsString()
    if discipline not in Disciplines:
        Disciplines[discipline] = []
    Disciplines[discipline].append(sheet)


# output.print_md("### DISCIPLINES: **{}**".format(", ".join(Disciplines.keys())))



# Choose a discipline for the sheet set
# If the user selects "ALL SHEETS", use that as the sheet set name
sheetsetname = None
while not sheetsetname:
    sheetsetname = forms.SelectFromList.show(
        Disciplines.keys(),
        title="Select Disciplines",
        multiselect=False
    )
    if not sheetsetname:
        script.exit()


# Create ViewSet
myviewset = DB.ViewSet()
for el in Disciplines[sheetsetname]:
    myviewset.Insert(el)



# Collect existing sheet sets
viewsheetsets = DB.FilteredElementCollector(revit.doc)\
                    .OfClass(framework.get_type(DB.ViewSheetSet))\
                    .WhereElementIsNotElementType()\
                    .ToElements()

allviewsheetsets = {vss.Name: vss for vss in viewsheetsets}




#_____________________________________________________________________ 🏃‍➡️ RUN 
with revit.Transaction('Created Sheet Set'):
    # Delete existing matching sheet set
    try:
        if sheetsetname in allviewsheetsets.keys():
            viewsheetsetting.CurrentViewSheetSet = allviewsheetsets[sheetsetname]
            viewsheetsetting.Delete()
    

        # Create new sheet set
        viewsheetsetting.CurrentViewSheetSet.Views = myviewset
        viewsheetsetting.SaveAs(sheetsetname)
        output.print_md("### Sheet Set Created: **{}** with **{}** sheets".format(sheetsetname, len(Disciplines[sheetsetname])))
        # print("Sheet Set Created: ", sheetsetname, " with ", len(Disciplines[sheetsetname]), " sheets.")

    except Exception as e:
            print("Error ", "Failed to create sheet set: {e}".format(e=str(e)))



# # Set name for the print set
# transaction = Transaction(doc, "Create Sheet Set")
# transaction.Start()
# try:
#     printSetName = "ALL SHEETS"
#     viewSheetSetting = doc.PrintManager.ViewSheetSetting
#     viewSheetSetting.CurrentViewSheetSet.Views = view_set
#     viewSheetSetting.SaveAs(printSetName)
#     transaction.Commit()
#     print("Success", "Sheet set '{p}' created with {a} sheets".format(p=printSetName, a=len(all_sheets)))
#     # TaskDialog.Show("Success", f"Sheet set '{printSetName}' created with {len(all_sheets)} sheets")
# except Exception as e:
#     transaction.RollBack()
#     print("Error ", "Failed to create sheet set: {e}".format(e=str(e)))
#     # TaskDialog.Show("Error", f"Failed to create sheet set: {str(e)}")

#🤖 Automate Your Boring Work Here


#==================================================
