# -*- coding: UTF-8 -*-
# -------------------------------------------
# This hook runs after a family is loaded into a Revit project.
# Log the family-loaded event to a JSON file.
# -------------------------------------------
from calendar import c
import os, json, time

# import pyrevit modules
from pyrevit import forms, revit
from pyrevit.script import output
from pyrevit import EXEC_PARAMS

# output_window = output.get_output()

# Gather doc info
doc = revit.doc
doc_path = doc.PathName or "<Untitled>"
doc_title = doc.Title

# Gather Revit info
version_build = doc.Application.VersionBuild
version_name = doc.Application.VersionName
version_number = doc.Application.VersionNumber

# Gather user info
username = doc.Application.Username

# Gather family info
family_name = EXEC_PARAMS.event_args.FamilyName
family_path = EXEC_PARAMS.event_args.FamilyPath


# json log location
# \FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
# C:\Users\kyleg\FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
log_dir = os.path.join(os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents", "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")

log_file = os.path.join(log_dir, username + "_revit_log.json")


# Determine family origin based on path
if family_path is None:
    family_origin = "Family Editor"
elif "AppData" in family_path:
    family_origin = "Content Catalog"
elif "172.16.1.7" in family_path:
    if "Drafting" in family_path:
        family_origin = "FFE Server - Revit Library"
    else:
        family_origin = "FFE Server"
else:
    family_origin = "Local"


# Create data entry
dataEntry = {
    "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
    "username": username,
    "doc_title": doc_title,
    "doc_path": doc_path,
    "revit_version": version_name,
    "revit_version_number": version_number,
    "revit_build": version_build,
    "action": "family-loaded",
    "family_name": family_name,
    "family_path": family_path,
    "family_origin": family_origin
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


# """
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
