# -*- coding: utf-8 -*-
__title__     = "Revision\nComparison"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 11.07.2025
________________________________________________________________
Tested Revit Versions: 
______________________________________________________________
Description:
This tool will compare the Revisions in the host model to the Revisions in the linked Revit models.
______________________________________________________________
How-to:
 -> Click the button
 -> Review the output for each link in the console
 -> Select the open in browser button to open the output in your web browser for easier reading
______________________________________________________________
Last update:
 - [11.07.2025] - v1.0 RELEASE
 - [11.10.2025] - v1.0.1 Temporarily updated to handle models set to "Per Sheet".
______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (SYSTEM)
from json import load
import re
import sys
from collections import OrderedDict
#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance, Revision

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
action = "Revision Comparison"

output_window = output.get_output()


#____________________________________________________________________ FUNCTIONS
def sanitize(v):
    """Return a friendly string for table cells."""
    if v is True:  return "True"
    if v is False: return "False"
    if v is None:  return "N/A"
    return str(v)

def check_revision_settings(document):
    """ Check if the document is set to by project or by sheet."""
    rev_settings = FilteredElementCollector(document).OfClass(RevisionSettings).FirstElement()
    if rev_settings:
        return rev_settings.RevisionNumbering
    return None

def get_revisions(document):
    """Get Revisions from the given document."""
    data = OrderedDict()
    rev_collector = FilteredElementCollector(document).OfClass(Revision)
    rev_numbering_sequence = FilteredElementCollector(document).OfClass(RevisionNumberingSequence)

    # Create a dictionary of Revision Numbering Sequences for easy lookup
    NumberingSequence = {}
    for rev_seq in rev_numbering_sequence:    
        NumberingSequence[rev_seq.Id.ToString()] = rev_seq.Name

    # Collect revision data
    for rev in rev_collector:
        revisions = {}
        rev_num_seq = rev.RevisionNumberingSequenceId.ToString()
        
        revisions['Sequence Number'] = rev.SequenceNumber
        revisions['Revision Number'] = rev.RevisionNumber
        revisions['Numbering'] = NumberingSequence[rev_num_seq] if rev_num_seq in NumberingSequence else "None"
        revisions['Revision Date'] = rev.RevisionDate
        revisions['Description'] = rev.Description
        revisions['Issued'] = rev.Issued
        revisions['Issued To'] = rev.IssuedTo
        revisions['Issued By'] = rev.IssuedBy
        revisions['Show'] = rev.Visibility

        data[rev.SequenceNumber] = revisions
    
    return data

def get_revisions_PerSheet(document):
    """Get Revisions from the given document."""
    data = OrderedDict()
    rev_collector = FilteredElementCollector(document).OfClass(Revision)
    rev_numbering_sequence = FilteredElementCollector(document).OfClass(RevisionNumberingSequence)

    # Create a dictionary of Revision Numbering Sequences for easy lookup
    NumberingSequence = {}
    for rev_seq in rev_numbering_sequence:    
        NumberingSequence[rev_seq.Id.ToString()] = rev_seq.Name

    # Collect revision data
    for rev in rev_collector:
        revisions = {}
        rev_num_seq = rev.RevisionNumberingSequenceId.ToString()
        
        revisions['Sequence Number'] = rev.SequenceNumber
        # revisions['Revision Number'] = rev.RevisionNumber
        revisions['Numbering'] = NumberingSequence[rev_num_seq] if rev_num_seq in NumberingSequence else "None"
        revisions['Revision Date'] = rev.RevisionDate
        revisions['Description'] = rev.Description
        revisions['Issued'] = rev.Issued
        revisions['Issued To'] = rev.IssuedTo
        revisions['Issued By'] = rev.IssuedBy
        revisions['Show'] = rev.Visibility

        data[rev.SequenceNumber] = revisions
    
    return data


def compare_revisions(host_revisions, link_revisions):
    """Compare Revisions between host and link documents."""
    comparison = []
    all_keys = set(host_revisions.keys()).union(set(link_revisions.keys()))

    for key in all_keys:
        host_rev = host_revisions.get(key)
        link_rev = link_revisions.get(key)

        if host_rev != link_rev:
            comparison.append((key, host_rev, link_rev))
    
    return comparison


def compare_values(host_value, link_value):
    """Compare two values and return a status icon."""
    if host_value != link_value:
        return "{} ❌".format(link_value)
    else:
        return "{} ✅".format(link_value)

#____________________________________________________________________ MAIN
rev_numbering = check_revision_settings(doc).ToString()
if rev_numbering == "PerSheet":
    TaskDialog.Show(
        __title__, 
        "This tool currently only works with models that have their Revision settings set to 'Per Project'.")
    sys.exit()


# Host Data
host_title = doc.Title
host_revisions = get_revisions(doc)

# Link Data
link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
loaded_links = [(li, li.GetLinkDocument()) for li in link_instances if li.GetLinkDocument()]

# Exit if no links found
if not loaded_links:
    TaskDialog.Show(__title__, "No loaded Revit links found in this model.")
    sys.exit()

# Header
output_window.print_md("# {action}".format(action=action))
output_window.print_md("## **Host Model:** `{host_title}`".format(host_title=host_title))
output_window.print_md("**Revision Numbering Settings:** `{rev_numbering}`".format(rev_numbering=rev_numbering))

### Host table
host_columns = ["Sequence Number", "Revision Number", "Numbering", "Revision Date", "Description", "Issued", "Issued To", "Issued By", "Show"]

host_rows = [
    [v["Sequence Number"], v["Revision Number"], v["Numbering"], v["Revision Date"], v["Description"], v["Issued"], v["Issued To"], v["Issued By"], v["Show"]]
    for n, v in sorted(host_revisions.items(), key=lambda item: item[0])
    ]

# Print Host Revisions Table
output_window.print_table(table_data=host_rows, columns=host_columns, title="Host Revisions ({})".format(len(host_revisions)))


### Per-Link
for li, ldoc in loaded_links:
    link_name = ldoc.Title
    output_window.print_md("---")
    
    # Link Header
    output_window.print_md("## Link: `{}`".format(link_name))
    
    # Get Link Revisions
    link_revisions = get_revisions(ldoc)
    # output_window.print_md("**Link Revisions ({}):**".format(len(link_revisions)))

    # Compute Differences early so we can tag mismatches in the table
    comparison = compare_revisions(host_revisions, link_revisions)
    

    link_rows = [
        [
            str(compare_values(host_revisions[n]["Sequence Number"] if n in host_revisions.keys() else "None",  v["Sequence Number"])), 
            compare_values(host_revisions[n]["Revision Number"] if n in host_revisions.keys() else "None",      v["Revision Number"]), 
            compare_values(host_revisions[n]["Numbering"] if n in host_revisions.keys() else "None",            v["Numbering"]), 
            compare_values(host_revisions[n]["Revision Date"] if n in host_revisions.keys() else "None",        v["Revision Date"]), 
            compare_values(host_revisions[n]["Description"] if n in host_revisions.keys() else "None",          v["Description"]), 
            compare_values(host_revisions[n]["Issued"] if n in host_revisions.keys() else "None",               v["Issued"]), 
            compare_values(host_revisions[n]["Issued To"] if n in host_revisions.keys() else "None",            v["Issued To"]), 
            compare_values(host_revisions[n]["Issued By"] if n in host_revisions.keys() else "None",            v["Issued By"]), 
            compare_values(host_revisions[n]["Show"] if n in host_revisions.keys() else "None",                 v["Show"]), 
            ]
        for n, v in sorted(link_revisions.items(), key=lambda item: item[0])
        ]
    
    
    output_window.print_table(table_data=link_rows, columns=host_columns, title="Link Revisions ({})".format(len(link_revisions)))



#____________________________________________________________________ RUN



log_status = "Success"
#______________________________________________________ LOG ACTION
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