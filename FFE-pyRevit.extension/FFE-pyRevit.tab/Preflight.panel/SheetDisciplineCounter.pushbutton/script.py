# -*- coding: utf-8 -*-
__title__     = "Sheet Discipline \nCounter"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 05.13.2026
______________________________________________________________
Description:
-> Enumerates Sheets for the FFE_Sheet_Order parameter, grouped by FFE_Sheet_Discipline Index.
______________________________________________________________
How-to:
-> Press Button
-> Check results
______________________________________________________________
Last update:
- [05.13.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
import re


#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

log_status = ""
action = "Sheet Discipline Counter"

output_window = output.get_output()
"""Output window for displaying results."""



#____________________________________________________________________ FUNCTIONS

"""
pyRevit | FFE - Sheet Discipline Counter

- Collect sheets from current model (and optionally all links)
- Filter: Appears In Sheet List + not placeholder
- Sort by Sheet Number
- Create sequence 00..N for current model sheets
- Write per-sheet, per-discipline index to FFE_Sheet_Order parameter
"""



# ----------------------- CONFIG (EDIT THESE) -----------------------

DISCIPLINE_INDEX_PARAM  = "FFE_Sheet_Discipline Index"
DISCIPLINE_PARAM        = "FFE_Sheet_Discipline"
ORDER_PARAM             = "FFE_Sheet_Order"



# Filters (match common Dynamo "sheet list" expectations)
EXCLUDE_PLACEHOLDERS = True
REQUIRE_APPEARS_IN_SHEET_LIST = True


# ----------------------- PARAMETER VALIDATION -----------------------
def parameter_exists_on_category(doc, param_name, bic):
    """Check if a parameter is bound to a category."""
    bindings = doc.ParameterBindings
    it = bindings.ForwardIterator()
    it.Reset()


    while it.MoveNext():
        definition = it.Key
        # binding = it.Current

        if definition.Name == param_name:
            return True

    return False


def validate_parameters(doc):
    """Check that the necessary parameters exist in the model."""    
    missing_params = []
    required_sheet_params = [DISCIPLINE_INDEX_PARAM, ORDER_PARAM]

    for pname in required_sheet_params:
        if not parameter_exists_on_category(doc, pname, DB.BuiltInCategory.OST_Sheets):
            missing_params.append("Sheet parameter missing: '{}'".format(pname))

    if missing_params:
        forms.alert(
            "Required parameters are missing:\n\n" + "\n".join(missing_params) +
            "\n\nAdd Parameters from Parameter Service" +
            "\n\nScript cancelled.",
            title="FFE Sheet Counter – Missing Parameters",
            warn_icon=True
        )
        return False

    return True


# ----------------------- FUNCTIONS -----------------------
def collect_sheets(doc):
    """Collect sheets from the model, applying filters."""
    sheets = (DB.FilteredElementCollector(doc)
              .OfClass(DB.ViewSheet)
              .WhereElementIsNotElementType()
              .ToElements())

    out = []
    for s in sheets:
        if EXCLUDE_PLACEHOLDERS and getattr(s, "IsPlaceholder", False):
            continue

        if REQUIRE_APPEARS_IN_SHEET_LIST:
            p = s.get_Parameter(DB.BuiltInParameter.SHEET_SCHEDULED)
            if p and p.StorageType == DB.StorageType.Integer and p.AsInteger() != 1:
                continue

        out.append(s)
    return out


def sort_sheets_by_number(sheets):
    """Sort sheets by their Sheet Number, handling common formatting."""
    def sheet_number_key(sheet):
        num = sheet.SheetNumber
        # Extract numeric part for sorting, handle common formats like "A101", "101A", "1-01"
        match = re.search(r'(\d+)', num)
        return int(match.group(1)) if match else float('inf')

    return sorted(sheets, key=sheet_number_key)




# ----------------------- MAIN -----------------------
def main():
    host_doc = revit.doc

    if not validate_parameters(host_doc):
        return
    
    # ---- Build MASTER list ----
    host_sheets = collect_sheets(host_doc)

    # ---- Group by discipline index ----
    sheets_by_discipline = {}
    for s in host_sheets:
        discipline_index = s.LookupParameter(DISCIPLINE_PARAM).AsString()
        if discipline_index not in sheets_by_discipline:
            sheets_by_discipline[discipline_index] = []
        sheets_by_discipline[discipline_index].append(s)
    
    # ---- For each discipline sort by specified sheet number ----
    


if __name__ == "__main__":
    main()








log_status = "Success"
#______________________________________________________ LOG ACTION
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

    # If it does exist, write to it
    # Check if "action" key exists, if not create it
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