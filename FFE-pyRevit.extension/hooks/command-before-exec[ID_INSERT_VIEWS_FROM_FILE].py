# -*- coding: utf-8 -*-
# -------------------------------------------
"""
__title__   = "command-before-exec[ID_INSERT_VIEWS_FROM_FILE]"
__doc__     = Version = v0.1
Date    = 10.24.2025
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs before the "Insert Views from File" command is executed.
# It logs the sevent to a JSON file.
________________________________________________________________
Last update:
- [10.24.2025] - v0.1 BETA
________________________________________________________________
"""

#____________________________________________________________________ IMPORTS
#⬇️ Imports
from pyrevit import revit, EXEC_PARAMS
import os, json, time

from pyrevit import forms, revit
from pyrevit.script import output

from Autodesk.Revit.DB.Events import RevitEventArgs, RevitAPIPostEventArgs
from Autodesk.Revit.UI.Events import ExecutedEventArgs
#____________________________________________________________________ VARIABLES
sender = __eventsender__ # UIApplication
args   = __eventargs__   # Autodesk.Revit.UI.Events.BeforeExecutedEventArgs






# doc = revit.doc

output_window = output.get_output()


# Gather doc info
doc = revit.doc
doc_path = doc.PathName or "<Untitled>"



output_window.print_md(doc.Title)

output_window.print_md("---")

output_window.print_md("### **Logging 'Insert Views from File' command**")

print(sender)
print(args)
print(args.ActiveDocument.Title)
print(args.CommandId)

### TEST THIS ###


"""
doc = revit.doc
doc_path = doc.PathName or "<Untitled>"



doc_title = doc.Title
version_build = doc.Application.VersionBuild
version_number = doc.Application.VersionNumber
username = doc.Application.Username



# json log location
log_dir = os.path.join(os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents", "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")
log_file = os.path.join(log_dir, username + "_revit_log.json")


# log data entry
dataEntry = {
    "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
    "username": username,
    "doc_title": doc_title,
    "doc_path": doc_path,
    "revit_version_number": version_number,
    "revit_build": version_build,
    "action": "Insert Views from File"
}



# Function to write JSON data
def write_json(dataEntry, filename=log_file):
    with open(filename,'r+') as file:
        # First we load existing data into a dict.
        file_data = json.load(file)
        # Join new_data with file_data inside emp_details
        file_data['family-loaded'].append(dataEntry)
        # Sets file's current position at offset.
        file.seek(0)
        # convert back to json.
        json.dump(file_data, file, indent = 4)
"""





"""
# Check if log file exists, if not create it
synclog = False
if not os.path.exists(log_file):
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    with open(log_file, 'w') as file:
        # create json structure
        file.write('{"family-loaded": []}')
    # output_window.print_md("### **Created log file:** `{}`".format(log_file))

# If it does exist, write to it
# Check if "family-loaded" key exists, if not create it
with open(log_file,'r+') as file:
    file_data = json.load(file)
    if 'family-loaded' not in file_data:
        file_data['family-loaded'] = []
        file.seek(0)
        json.dump(file_data, file, indent = 4)

try:
    write_json(dataEntry)
    synclog = True
    # output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
except Exception as e:
    synclog = False
    # output_window.print_md("### **Failed to log sync to JSON:** `{}`".format(e))
# """