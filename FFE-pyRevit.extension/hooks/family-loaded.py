# -*- coding: UTF-8 -*-
# -------------------------------------------
"""
__title__   = "family-loaded"
__doc__     = Version = v1.2
Date    = 10.07.2025
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
- This hook runs after a family is loaded into a Revit project.
- Log the family-loaded event to a JSON file.
________________________________________________________________
Last update:
- [10.07.2025]  - v1.0 RELEASE
- [10.08.2025]  - v1.1 added Family Editor origin; corrected origin
- [10.19.2025]  - v1.2 changed server address from IP to "Internal Share"
                - removed redundant revit version info
________________________________________________________________
"""
#____________________________________________________________________ IMPORTS
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


### TEST THIS ###
# Define origin types and keywords
# origin_types = {
#     "Family Editor": [], 
#     "Content Catalog": ['AppData'], 
#     "FFE Server": ["172.16.1.7"], 
#     "FFE Server - Revit Library": ["172.16.1.7", "Drafting"],
#     "Local": []
#     }

# family_origin = None
# # Determine family origin based on path and keywords
# if family_path is None:
#     family_origin = "Family Editor"
# else:
#     for origin, keywords in origin_types.items():
#         if all(keyword in family_path for keyword in keywords):
#             family_origin = origin
#             break
#     if family_origin is None:
#         family_origin = "Local"
### TEST THIS ###


# Determine family origin based on path
if family_path is None:
    family_origin = "Family Editor"
elif "AppData" in family_path:
    family_origin = "Content Catalog"
elif "Internal Share" in family_path:
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
    # output_window.print_md("### **Failed to log sync to JSON:** `{}`".format(e))
