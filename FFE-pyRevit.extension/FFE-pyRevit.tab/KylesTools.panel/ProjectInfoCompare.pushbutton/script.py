# -*- coding: utf-8 -*-
__title__     = "Project Info\nComparison"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 11.04.2025
________________________________________________________________
Tested Revit Versions: 
______________________________________________________________
Description:
This tool will compare the Project Information in the host model to the Project Information in the linked Revit models.
______________________________________________________________
How-to:
 -> Click the button
 -> Review the output for each link in the console
 -> Select the open in browser button to open the output in your web browser for easier reading
______________________________________________________________
Last update:
 - [11.04.2025] - v1.0 RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (SYSTEM)
import re
import sys
from collections import OrderedDict
#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance, PhaseFilter


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

log_status = ""

output_window = output.get_output()

action = "Project Info Comparison"

#____________________________________________________________________ FUNCTIONS
def get_project_info_elements(document):
    """Get Project Information elements from the given document."""
    project_info = {}
    pi = document.ProjectInformation
    if pi:
        project_info['FFE_Sheet_Project Number'] = pi.GetParameters("FFE_Sheet_Project Number")[0].AsString() if pi.GetParameters("FFE_Sheet_Project Number") else "N/A"
        project_info['FFE_Sheet_Campus Name'] = pi.GetParameters("FFE_Sheet_Campus Name")[0].AsString() if pi.GetParameters("FFE_Sheet_Campus Name") else "N/A"
        project_info['FFE_Sheet_Project Location'] = pi.GetParameters("FFE_Sheet_Project Location")[0].AsString() if pi.GetParameters("FFE_Sheet_Project Location") else "N/A"
        project_info['FFE_Sheet_Project Phase'] = pi.GetParameters("FFE_Sheet_Project Phase")[0].AsString() if pi.GetParameters("FFE_Sheet_Project Phase") else "N/A"
        project_info['FFE_Sheet_Project Title'] = pi.GetParameters("FFE_Sheet_Project Title")[0].AsString() if pi.GetParameters("FFE_Sheet_Project Title") else "N/A"
        project_info['Issue Date'] = pi.IssueDate
    return project_info


#____________________________________________________________________ MAIN

def main():
    global log_status

    host_project_info = get_project_info_elements(doc)

    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

    if not link_instances:
        forms.alert("No linked Revit models found in the current project.", title=action)
        return

    output_window.print_md("# {action}".format(action=action))
    output_window.print_md("## Host Model Project Information:")
    for key, value in host_project_info.items():
        output_window.print_md("- **{key}:** {value}".format(key=key, value=value))

    output_window.print_md("---")

    for link_instance in link_instances:
        link_doc = link_instance.GetLinkDocument()
        if link_doc:
            link_project_info = get_project_info_elements(link_doc)
            output_window.print_md("## Linked Model: {link_doc}".format(link_doc=link_doc.Title))
            for key in host_project_info.keys():
                host_value = host_project_info.get(key, "N/A")
                link_value = link_project_info.get(key, "N/A")
                if host_value != link_value:
                    output_window.print_md("- **{key}:** Host = {host_value} | (❌) | Link = {link_value}".format(key=key, host_value=host_value, link_value=link_value))
                else:
                    output_window.print_md("- **{key}:** {host_value} (✅)".format(key=key, host_value=host_value))

    log_status = "Success"
    output_window.print_md("---")
    output_window.print_md("### {action} - {log_status}".format(action=action, log_status=log_status))


#____________________________________________________________________ RUN
if __name__ == "__main__":
    main()



#______________________________________________________ LOG ACTION
action = "Project Info Comparison"
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

log_action(action, log_status)
# output_window.print_md("Logging action: {}".format(log_action(action, log_status)))