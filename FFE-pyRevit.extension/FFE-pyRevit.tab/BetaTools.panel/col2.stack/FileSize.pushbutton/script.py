# -*- coding: utf-8 -*-
__title__   = "Model File Size"
__author__  = "Kyle Guggenheim"
__doc__     = """Show file size for the active Revit document (local & central when available)."""

# pyRevit
from pyrevit.script import output
output_window = output.get_output()  # needed for print_md()

# .NET / Revit
import clr
clr.AddReference("System")
from System.IO import FileInfo, File, Path

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (Document, ModelPathUtils, WorksharingUtils)


#____________________________________________________________________ FUNCTIONS
def bytes_to_human(nbytes):
    # Returns (value, unit) like (123.4, 'MB')
    if nbytes is None:
        return (None, '')
    step = 1024.0
    units = ['B','KB','MB','GB','TB']
    size = float(nbytes)
    for u in units:
        if size < step or u == units[-1]:
            return (round(size, 2), u)
        size /= step

def file_size_bytes(user_visible_path):
    """Return size in bytes if path points to an existing file; else None."""
    try:
        if user_visible_path and File.Exists(user_visible_path):
            return FileInfo(user_visible_path).Length
    except:
        pass
    return None

def is_cloud_doc(doc):
    """Best-effort cloud detection: check central path if workshared; fallback to path name."""
    try:
        if doc.IsWorkshared:
            mpath = WorksharingUtils.GetCentralModelPath(doc)
            return ModelPathUtils.IsCloudPath(mpath)
    except:
        pass
    # Non-workshared cloud docs typically don’t exist; keep False
    return False


#____________________________________________________________________ MAIN
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document  # type: Document

doc_title = doc.Title
is_workshared = doc.IsWorkshared
is_cloud = is_cloud_doc(doc)

# Local (or RFA) path for the *opened* document
local_user_path = doc.PathName  # empty for unsaved docs and most cloud docs
local_size_b = file_size_bytes(local_user_path)

# Central model (workshared only; may be UNC or cloud)
central_user_path = None
central_size_b = None
if is_workshared:
    try:
        central_mpath = WorksharingUtils.GetCentralModelPath(doc)
        central_user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_mpath)
        # Only works if it's an actual filesystem path (not BIM 360/ACC URI)
        central_size_b = file_size_bytes(central_user_path)
    except:
        pass

# Build report
rows = []
def add_row(k, v):
    rows.append("| {k} | {v} |".format(k=k, v=v))

add_row("Document", doc_title)
add_row("Saved", "Yes" if local_user_path else "No")
add_row("Workshared", "Yes" if is_workshared else "No")
add_row("ACC/BIM 360 (cloud)", "Yes" if is_cloud else "No")

# Local size
if local_size_b is not None:
    # print(local_size_b)
    v, u = bytes_to_human(local_size_b)
    add_row("Local file size", "{} {}".format(v, u))
else:
    add_row("Local file size", "N/A (unsaved or non-filesystem path)")

# Central size
if is_workshared:
    if central_size_b is not None:
        # print(central_size_b)
        v, u = bytes_to_human(central_size_b)
        add_row("Central file size", "{} {}".format(v, u))
    else:
        # If cloud, explain limitation
        if is_cloud:
            add_row("Central file size", "N/A for ACC cloud models via API")
        else:
            add_row("Central file size", "N/A (central path not accessible as a file)")

# Paths (shown if meaningful)
if local_user_path:
    add_row("Local path", local_user_path)
if central_user_path and central_user_path != local_user_path:
    add_row("Central path", central_user_path)

# Output table
header = "| Field | Value |\n|---|---|"
output_window.print_md(header + "\n" + "\n".join(rows))

# Helpful notes
notes = []
if not local_user_path:
    notes.append("- The active document hasn’t been saved yet (or is a cloud model without a local file path).")
if is_cloud:
    notes.append("- For ACC/BIM 360 cloud models opened directly from the cloud, Revit’s API does not expose the RVT file on disk, so size cannot be read this way.")
    notes.append("- Workarounds: use Desktop Connector to a synced filesystem path, or query size via Autodesk/ACC APIs outside Revit.")
if notes:
    output_window.print_md("\n**Notes**\n" + "\n".join(notes))
