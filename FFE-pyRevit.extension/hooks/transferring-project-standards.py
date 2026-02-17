# -*- coding: UTF-8 -*-
"""
__title__   = "transferring-project-standards"
__doc__     = Version = v1.0
Date    = 11.04.2025
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs after a transferring project standards is completed.
# It logs the event to a JSON file.
________________________________________________________________
Last update:
- [10.23.2025] - v0.1 BETA
________________________________________________________________
"""
#____________________________________________________________________ IMPORTS
import os, json, time

from pyrevit import forms, revit
from pyrevit.script import output
from pyrevit import HOST_APP, EXEC_PARAMS

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
print_args = {
    "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
    "cancellable?": str(args.Cancellable),
    "doc": str(revit.doc),
    "source_doc": str(args.SourceDocument.PathName),
    "target_doc": str(args.TargetDocument.PathName),
    "ext_items": str(args.GetExternalItems()),
}
print("transferring-project-standards: {}".format(print_args))







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
    "action": "transferred-project-standards"
}

output_window.print_md("{}".format(dataEntry))