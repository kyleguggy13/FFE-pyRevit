# -*- coding: utf-8 -*-
__title__     = "Text Leader \nPosition"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.11.2025
# _____________________________________________________________________
# Description:
# - This script allows you to change the left/right 
#   attachment position of text leaders in Revit.
# - The script will prompt you to select the 
#   attachment position
# - The available positions are Top, Middle, and Bottom.
# 
# _____________________________________________________________________
# How-to:
#
# -> Select the text elements you want to change
# -> Click the button
# -> Choose the Left Attachment position from the list
# -> Choose the Right Attachment position from the list
# _____________________________________________________________________
# Last update:
# - [06.05.2025] - 1.0 RELEASE
# _____________________________________________________________________
Inspiration: Olivia Bates
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

from logging import Filter
from requests import get
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import Selection


#____________________________________________________________________ IMPORTS (PYREVIT)

from pyrevit import revit, DB, forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


#____________________________________________________________________ MAIN

### NOTES ###
# LeaderAtachement Enumeration
# https://www.revitapidocs.com/2026/82ed0368-6da3-53a3-8c07-4061efd0be56.htm



# GET SELECTION
selected_elements = selection.GetElementIds()

print("Selected Elements: ", len(selected_elements))

# CONTAINER
elements_list = []
param_Left_List = []
param_Right_List = []

# LOOP THROUGH SELECTED ELEMENTS
for element_id in selected_elements:
    element = doc.GetElement(element_id)
    elements_list.append(element)
    param_Left_List.append(element.get_Parameter(BuiltInParameter.LEADER_LEFT_ATTACHMENT))
    param_Right_List.append(element.get_Parameter(BuiltInParameter.LEADER_RIGHT_ATTACHMENT))

    # print("Element ID: ", element.Id, " Element Parameters: ", element.GetParameters('Left Attachment')[0].AsValueString())
    # print("get_Parameter: ", element.get_Parameter(BuiltInParameter.LEADER_LEFT_ATTACHMENT).AsValueString())




# Step 1: Ask user for Top, Middle, or Bottom attachment
options = ["Top", "Middle", "Bottom"]
attachment_choice_left = forms.SelectFromList.show(options, title="Select Left Leader Attachment Position", button_name="Apply")
attachment_choice_right = forms.SelectFromList.show(options, title="Select Right Leader Attachment Position", button_name="Apply")

print("Left Attachment Position: ", attachment_choice_left)
print("Right Attachment Position: ", attachment_choice_right)


### If not selected, exit the script
if not attachment_choice_left and not attachment_choice_right:
    forms.alert("No attachment selected. Exiting.", exitscript=True)



# Step 2: Use raw int for AttachmentType enum
attachment_map = {
    "Top": 0,
    "Middle": 1,
    "Bottom": 2
}

if attachment_choice_left not in attachment_map:
    attachment_value_left = ""
else:
    attachment_value_left = attachment_map[attachment_choice_left]


if attachment_choice_right not in attachment_map:
    attachment_value_right = ""
else:
    attachment_value_right = attachment_map[attachment_choice_right]


#____________________________________________________________________ 🤖 Transaction

# Set Attachment Parameters
transaction = Transaction(doc, "Text Leader")
transaction.Start()
try:
    print("Changing text leader position...")
    for param_left, param_right in zip(param_Left_List, param_Right_List):
        ## Set the Left Attachment parameter to the selected attachment position
        if attachment_value_left == "":
            print("Left Attachment not set, skipping...")
        else:
            param_left.Set(attachment_value_left)
            print("Left Leader position changed to: ", attachment_choice_left)
        
        ## Set the Right Attachment parameter to the selected attachment position
        if attachment_value_right == "":
            print("Right Attachment not set, skipping...")
        else:
            param_right.Set(attachment_value_right)
            print("Right Leader position changed to: ", attachment_choice_right)


        ##
        # print("Left Leader position changed to: ", attachment_choice_left)
        # print("Right Leader position changed to: ", attachment_choice_right)

    transaction.Commit()
    log_status = "Success"
except Exception as e:

    transaction.RollBack()
    log_status = "Failed"
    print("Error ", "Failed to change leader position: ", str(e))



#______________________________________________________ LOG ACTION
action = "Text Leader Position"
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


#==================================================
