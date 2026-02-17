# -*- coding: UTF-8 -*-
"""
__title__   = "transferred-project-standards"
__doc__     = Version = v1.0
Date    = 02.17.2026
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs after a transfer project standards is completed.
# It logs the event to a JSON file.
________________________________________________________________
Last update:
- [02.16.2026] - v0.1 BETA
- [02.17.2026] - v1.0 RELEASE
________________________________________________________________
"""
#⬇️ Imports
from pyrevit import revit, EXEC_PARAMS
from Autodesk.Revit.UI import TaskDialog

#--------------------------------------------------
#📦 Variables
# sender = __eventsender__ # UIApplication
# args   = __eventargs__   # Autodesk.Revit.UI.Events.BeforeExecutedEventArgs
doc = revit.doc


#--------------------------------------------------
# # 🎯 MAIN
# if not doc.IsFamilyDocument:
#     #⚠️ Show Warning
#     TaskDialog.Show('Big Brother is Watching!',
#                     'Import CAD is not Allowed! Use Link CAD Instead.')

#     #🔒 Ask user for Password
#     from pyrevit.forms import ask_for_string
#     password   = 'BIM'
#     user_input = ask_for_string(prompt='Only users with a password can Import CAD.',
#                                 title='Import CAD Blocked')
#     # ❌ Stop Execution
#     if user_input != password:
#         args.Cancel = True
# else:
#     TaskDialog.Show('Family CAD Import',
#                     'Import CAD is Allowed in families!')



#____________________________________________________________________ IMPORTS
import os, json, time

# from pyrevit import forms, revit
from pyrevit.script import output
# from pyrevit import HOST_APP, EXEC_PARAMS

output_window = output.get_output()

# Gather doc info
doc = revit.doc
doc_path = doc.PathName or "<Untitled>"

# Gather Revit info
doc_title = doc.Title
version_build = doc.Application.VersionBuild
version_number = doc.Application.VersionNumber

# Gather user info
username = doc.Application.Username



args = EXEC_PARAMS.event_args




output_window.print_md("### print_args: ")

output_window.print_md("- source_doc_path: {}".format(args))

# output_window.print_md("- target_doc: {}".format(str(args.TargetDocument.Title)))
# output_window.print_md("- target_doc_path: {}".format(str(args.TargetDocument.PathName)))




# json log location
# \FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
# C:\Users\kyleg\FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
log_dir = os.path.join(os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents", "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")

log_file = os.path.join(log_dir, username + "_revit_log.json")

dataEntry = {
    "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
    "username": username,
    "doc_title": doc_title,
    "doc_path": doc_path,
    "revit_version_number": version_number,
    "revit_build": version_build,
    "action": "ID_FILE_IMPORT",
}

output_window.print_md("{}".format(dataEntry))


"""
# Function to write JSON data
def write_json(dataEntry, filename=log_file):
    with open(filename,'r+') as file:
        file_data = json.load(file)                 # First we load existing data into a dict.
        file_data['action'].append(dataEntry)       # Join new_data with file_data inside emp_details.
        file.seek(0)                                # Sets file's current position at offset.
        json.dump(file_data, file, indent = 4)      # convert back to json.


# Check if log file exists, if not create it
logcheck = False
if not os.path.exists(log_file):
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    with open(log_file, 'w') as file:
        file.write('{"action": []}')                # create json structure


# If it does exist, write to it
# Check if "action" key exists, if not create it
with open(log_file,'r+') as file:
    file_data = json.load(file)
    if 'action' not in file_data:
        file_data['action'] = []
        file.seek(0)
        json.dump(file_data, file, indent = 4)
        

# If it does exist, write to it
try:
    write_json(dataEntry)
    logcheck = True
    # output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
except Exception as e:
    logcheck = False
    # output_window.print_md("### **Failed to log sync to JSON:** `{}`".format(e))
# """
