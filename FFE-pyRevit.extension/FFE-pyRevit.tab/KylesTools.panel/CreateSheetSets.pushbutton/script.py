# -*- coding: utf-8 -*-
__title__     = "Create Sheet Sets"
__version__   = 'Version = v1.1'
__doc__       = """Version = v1.1
Date    = 10.31.2025
___________________________________________________________________
Description:
This will create sheet sets that can be used for exporting to PDF and for publishing to ACC.
___________________________________________________________________
How-to:
-> Click on the button
-> Review existing print sets
-> Select SHEETS or a DISCIPLINE
-> Review the output in the console

-> NOTE: Only sheets with the "Appears In Sheet List" parameter checked on will be included in the sheet set.
___________________________________________________________________
Last update:
- [05.22.2025] - v0.1 BETA RELEASE
- [07.25.2025] - v1.0 RELEASE
- [10.31.2025] - v1.1 Updated logging to include status of action
___________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

from math import log
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

log_status = ""
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
### "\pyRevitTools.extension\pyRevit.tab\Drawing Set.panel\views.stack\Views.pulldown\Create Print Set From Selected Views.pushbutton"
### "\pyRevitTools.extension\pyRevit.tab\Drawing Set.panel\Sheets.pulldown\List TitleBlocks on Sheets.pushbutton\script.py"


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
# If the user selects "SHEETS", use that as the sheet set name
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
        log_status = "Success"

    except Exception as e:
        print("Error ", "Failed to create sheet set: {e}".format(e=str(e)))
        log_status = "Failed"



#______________________________________________________ LOG ACTION
action = "Creat Sheet Sets"
def log_action(action, log_status):
    """Log action to user JSON log file."""
    import os, json, time
    from pyrevit import revit

    doc = revit.doc
    doc_path = doc.PathName or "<Untitled>"

    doc_title = doc.Title
    version_build = doc.Application.VersionBuild
    version_number = doc.Application.VersionNumber
    username = doc.Application.Username
    action = action

    # json log location
    # \FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
    log_dir = os.path.join(os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents", "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")
    log_file = os.path.join(log_dir, username + "_revit_log.json")

    dataEntry = {
        "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
        "username": username,
        "doc_title": doc_title,
        "doc_path": doc_path,
        "revit_version_number": version_number,
        "revit_build": version_build,
        "action": action,
        "status": log_status
    }

    # Function to write JSON data
    def write_json(dataEntry, filename=log_file):
        with open(filename,'r+') as file:
            file_data = json.load(file)                 # First we load existing data into a dict.   
            file_data['action'].append(dataEntry)       # Join new_data with file_data inside emp_details
            file.seek(0)                                # Sets file's current position at offset.
            json.dump(file_data, file, indent = 4)      # convert back to json.


    # Check if log file exists, if not create it
    logcheck = False
    if not os.path.exists(log_file):
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        with open(log_file, 'w') as file:    
            file.write('{"action": []}')                # create json structure
        
        # output_window.print_md("### **Created log file:** `{}`".format(log_file))

    with open(log_file,'r+') as file:
        file_data = json.load(file)
        if 'action' not in file_data:
            file_data['action'] = []
            file.seek(0)
            json.dump(file_data, file, indent = 4)

    try:
        write_json(dataEntry)
        logcheck = True
        # output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
    except Exception as e:
        logcheck = False

    return dataEntry

# log_action(action, log_status)
output.print_md("Logging action: {}".format(log_action(action, log_status)))

