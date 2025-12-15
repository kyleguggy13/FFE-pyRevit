# -*- coding: utf-8 -*-
__title__   = "Copy View Filters"
__version__   = 'Version = v1.2'
__doc__     = """Version = v1.2
Date    = 10.31.2025
________________________________________________________________
Tested Revit Versions: 2026, 2024
________________________________________________________________
Description:
This buttom will copy view filters from one view to multiple views.
________________________________________________________________
How-To:
1. Press the button
2. Select the view with filters
3. Select the destination views
________________________________________________________________
Last update:
- [03.06.2025] - v1.0 RELEASE
- [10.03.2025] - v1.1 Compatible with Revit 2026
- [10.31.2025] - v1.2 Updated logging to include status of action
________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS

from math import e, log
from Autodesk.Revit.DB import *


from pyrevit.script import output

#____________________________________________________________________ VARIABLES
# app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document

log_status = ""

output_window = output.get_output()

#____________________________________________________________________ MAIN

# 1️⃣ Get Views
from pyrevit import forms

all_views_and_vt = FilteredElementCollector(doc)\
                    .OfCategory(BuiltInCategory.OST_Views)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

dict_views = {}
for v in all_views_and_vt:
    view_or_template = 'View' if not v.IsTemplate else 'ViewTemplate'
    key = '[{a}] ({b}) {c}'.format(a=view_or_template, b=v.ViewType, c=v.Name)
    dict_views[key] = v

# View From
views_from      = sorted(dict_views.keys())
sel_view_from   = forms.SelectFromList.show(views_from, multiselect=False, button_name='Select View with Filters')
view_from       = dict_views[sel_view_from]

# View To
views_to            = sorted(dict_views.keys())
sel_list_views_to   = forms.SelectFromList.show(views_to, multiselect=True, button_name='Select Destination Views')
list_views_to       = [dict_views[view_name] for view_name in sel_list_views_to]



# 2️⃣ Extract Filters
filter_ids  = view_from.GetOrderedFilters()
filters     = [doc.GetElement(f_id) for f_id in filter_ids]


# This was copied from PyRevit Docs: Select From List (Single, Multiple)
# https://pyrevitlabs.notion.site/Effective-Input-ea95e95282a24ba9b154ef88f4f8d056
sel_filters = forms.SelectFromList.show(filters,
                                        multiselect=True,
                                        name_attr='Name', # This needs to be something that exists in the objects in the list
                                        button_name='Select View Filters')



#_____________________________________________________________________ 🏃‍➡️ RUN 
# 3️⃣ Paste Filters
with Transaction(doc, 'Copy View Filters') as t:
    t.Start() # 🔓
    # ALL CHANGES HAVE TO BE HERE

    for view_to in list_views_to:

        if view_to.ViewTemplateId == ElementId.InvalidElementId: # Only applies view filters if no template.
            for v_filter in sel_filters:
            
                # 1. Copy Filters
                overrides = view_from.GetFilterOverrides(v_filter.Id)
                view_to.SetFilterOverrides(v_filter.Id, overrides)

                # 2. Copy Settings
                visibility = view_from.GetFilterVisibility(v_filter.Id)
                enable = view_from.GetIsFilterEnabled(v_filter.Id)

                view_to.SetFilterVisibility(v_filter.Id, visibility)
                view_to.SetIsFilterEnabled(v_filter.Id, enable)
        else:
            vt = doc.GetElement(view_to.ViewTemplateId)

            print("⚠️View ({v}) has a ViewTemplate({vt}) assigned. It will be skipped".format(v=view_to.Name, vt=vt.Name))


    t.Commit() # 🔒

log_status = "Success"


#______________________________________________________ LOG ACTION
action = "Copy View Filters"
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

log_action(action, log_status)
# output.print_md("Logging action: {}".format(log_action(action, log_status)))

