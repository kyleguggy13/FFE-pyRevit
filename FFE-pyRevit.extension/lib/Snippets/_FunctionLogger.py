# -*- coding: UTF-8 -*-
# -------------------------------------------
"""
__title__   = "Function Logger""
__doc__     = Version = v1.0
Date    = 10.30.2025
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs after as a function is performed.
# It logs the action event to a JSON file.
________________________________________________________________
Last update:
- [10.30.2025] - v0.1 BETA
________________________________________________________________
"""
#____________________________________________________________________ IMPORTS
import os, json, time

from pyrevit import forms, revit
from pyrevit.script import output

# output_window = output.get_output()

doc = revit.doc
doc_path = doc.PathName or "<Untitled>"

doc_title = doc.Title
version_build = doc.Application.VersionBuild
version_number = doc.Application.VersionNumber
username = doc.Application.Username



action = "action placeholder"  # This should be set to the specific action being logged


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
    "action": action
}



# Function to write JSON data
def write_json(dataEntry, filename=log_file):
    with open(filename,'r+') as file:
        # First we load existing data into a dict.
        file_data = json.load(file)
        # Join new_data with file_data inside emp_details
        file_data['action'].append(dataEntry)
        # Sets file's current position at offset.
        file.seek(0)
        # convert back to json.
        json.dump(file_data, file, indent = 4)



# Check if log file exists, if not create it
synclog = False
if not os.path.exists(log_file):
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    with open(log_file, 'w') as file:
        # create json structure
        file.write('{"action": []}')
    # output_window.print_md("### **Created log file:** `{}`".format(log_file))

with open(log_file,'r+') as file:
    file_data = json.load(file)
    if 'action' not in file_data:
        file_data['action'] = []
        file.seek(0)
        json.dump(file_data, file, indent = 4)
        
# If it does exist, write to it
try:
    write_json(dataEntry)
    synclog = True
    # output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
except Exception as e:
    synclog = False
    # output_window.print_md("### **Failed to log sync to JSON:** `{}`".format(e))
