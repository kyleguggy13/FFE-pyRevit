# -*- coding: UTF-8 -*-
# -------------------------------------------
"""
__title__   = "doc-synced"
__doc__     = Version = v1.4
Date    = 11.05.2025
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs after a document is synced to central.
# It logs the sync event to a JSON file.
________________________________________________________________
Last update:
- [09.19.2025] - v0.1 BETA
- [09.23.2025] - v1.0 RELEASE
- [09.24.2025] - v1.1 added revit version to log
- [10.07.2025] - v1.2 added doc title to log 
- [10.19.2025] - v1.3 removed redundant revit version info
- [11.05.2025] - v1.4 changed log structure to have action as main key
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
# filesize = doc.GetDocumentSize()


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
    "action": "sync"
}

output.get_output().print_md(dataEntry)


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

    # output_window.print_md("### **Created log file:** `{}`".format(log_file))


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
