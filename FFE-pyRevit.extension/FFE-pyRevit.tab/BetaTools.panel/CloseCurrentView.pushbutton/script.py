# -*- coding: utf-8 -*-
__title__     = "Close Current View"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 18.06.2026
________________________________________________________________
Tested Revit Versions: 
______________________________________________________________
Description:
This tool closes the currently active Revit view tab using the UIView class.
______________________________________________________________
How-to:
 -> Assign a Revit keyboard shortcut to this pyRevit button.
 -> Press the shortcut while the target view is active.
______________________________________________________________
Last update:
 - [18.06.2026] - v0.1 RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (SYSTEM)
import json
import os
import time

#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms, revit


#____________________________________________________________________ VARIABLES
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document

log_status = ""
action = "Close Current View"



#____________________________________________________________________ FUNCTIONS


def get_active_uiview(active_uidoc, active_view_id):
    """Return the open UIView that matches the active document view."""
    open_uiviews = list(active_uidoc.GetOpenUIViews())
    for uiview in open_uiviews:
        if uiview.ViewId == active_view_id:
            return uiview, open_uiviews
    return None, open_uiviews


def close_current_view(active_uidoc, active_doc):
    """Close the active UIView and return a log status message."""
    active_view = active_doc.ActiveView
    active_view_id = active_view.Id
    active_view_name = active_view.Name

    active_uiview, open_uiviews = get_active_uiview(active_uidoc, active_view_id)
    if len(open_uiviews) <= 1:
        return False, "Failed: Revit requires at least one open view."

    if active_uiview is None:
        return False, "Failed: Could not find an open UI view for '{0}'.".format(active_view_name)

    try:
        active_uiview.Close()
    except Exception as exc:
        return False, "Failed to close '{0}': {1}".format(active_view_name, exc)

    return True, "Closed active view: {0}".format(active_view_name)




#______________________________________________________ LOG ACTION
# action = "Project Info Comparison"
def log_action(action, log_status):
    """Log action to user JSON log file."""
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

success, log_status = close_current_view(uidoc, doc)

try:
    log_action(action, log_status)
except Exception:
    pass

if not success:
    forms.alert(log_status, title=__title__, warn_icon=True)
