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
    "cancellable?": str(args.Cancellable),
    "doc": str(revit.doc),
    "source_doc": str(args.SourceDocument),
    "target_doc": str(args.TargetDocument),
    "ext_items": str(args.GetExternalItems()),
}
print("transferring-project-standards: {}".format(print_args))

