# -*- coding: utf-8 -*-
__title__     = "Phase Filter\nComparison"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 09.13.2025
________________________________________________________________
Tested Revit Versions: 2026, 2024
______________________________________________________________
Description:
This tool will compare the phase filters in the host model to the phase filters in the linked Revit models.
______________________________________________________________
How-to:
 -> Click the button
 -> Review the output for each link in the console
 -> Select the open in browser button to open the output in your web browser for easier reading
______________________________________________________________
Last update:
 - [09.13.2025] - v0.1 Beta Release
 - [09.14.2025] - v1.0 First Release
______________________________________________________________
Author: Kyle Guggenheim"""

from math import log
import re
import sys
from collections import OrderedDict
#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance, PhaseFilter

# Newer API enums (safe-import with fallback)
try:
    from Autodesk.Revit.DB import ElementOnPhaseStatus, PhaseStatusPresentation
except Exception:
    ElementOnPhaseStatus = None
    PhaseStatusPresentation = None

#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


output_window = output.get_output()


action = "Phase Filter Comparison"

# li = link instance
# ldoc = link document
# lmap = link phase filter map
# hmap = host phase filter map

#____________________________________________________________________ FUNCTIONS
def md_list(items):
    """Render a Python list as one markdown string (never pass lists to print_md)."""
    try:
        lines = [u"- {}".format(u"{}".format(it)) for it in items]
    except Exception:
        lines = [u"- {}".format(str(it)) for it in items]
    return u"\n".join(lines)


def sanitize(v):
    """Return a friendly string for table cells."""
    if v is True:  return "True"
    if v is False: return "False"
    if v is None:  return "N/A"
    return str(v)


def _try_bool_show_prop(pf, label):
    """Try classic bool properties: ShowNew/Existing/Demolished/Temporary."""
    # Direct property
    try:
        val = getattr(pf, "Show" + label)
        if isinstance(val, bool):
            return val
    except Exception:
        pass
    # Common getter variants
    for name in ("get_Show" + label, "GetShow" + label, "Is" + label + "Shown"):
        try:
            m = getattr(pf, name)
            if callable(m):
                v = m()
                if isinstance(v, bool):
                    return v
        except Exception:
            pass
    return None


def _try_get_phase_status_presentation(pf, status_enum):
    """Try PhaseFilter.GetPhaseStatusPresentation(ElementOnPhaseStatus)."""
    if ElementOnPhaseStatus is None:
        return None
    try:
        getter = getattr(pf, "GetPhaseStatusPresentation")
    except Exception:
        return None
    if not callable(getter):
        return None
    try:
        return getter(status_enum)
    except Exception:
        return None


def _read_cell(pf, label, status_enum):
    """Best-effort read for one column (New/Existing/Demolished/Temporary)."""
    # Prefer the modern enum-based API when available
    v = _try_get_phase_status_presentation(pf, status_enum)
    if v is not None:
        return v
    # Fallback to legacy bools
    b = _try_bool_show_prop(pf, label)
    if isinstance(b, bool):
        return b
    return None


def get_phase_filters_map(rvt_doc):
    """Return dict: { filter_name: {New, Existing, Demolished, Temporary} }.
       Values are PhaseStatusPresentation enums or bools (or None if unknown).
    """
    data = OrderedDict()
    col = FilteredElementCollector(rvt_doc).OfClass(PhaseFilter)
    for pf in col:
        name = pf.Name
        # Map labels to enum members (if available)
        new_enum        = getattr(ElementOnPhaseStatus, "New", None) if ElementOnPhaseStatus else None
        existing_enum   = getattr(ElementOnPhaseStatus, "Existing", None) if ElementOnPhaseStatus else None
        demolished_enum = getattr(ElementOnPhaseStatus, "Demolished", None) if ElementOnPhaseStatus else None
        temporary_enum  = getattr(ElementOnPhaseStatus, "Temporary", None) if ElementOnPhaseStatus else None

        entry = {
            "New":         _read_cell(pf, "New",         new_enum),
            "Existing":    _read_cell(pf, "Existing",    existing_enum),
            "Demolished":  _read_cell(pf, "Demolished",  demolished_enum),
            "Temporary":   _read_cell(pf, "Temporary",   temporary_enum),
        }
        data[name] = entry
    return data


def compare_phase_filters(host_map, link_map):
    """Return (missing_in_link, extra_in_link, diffs).
       diffs: [(filter_name, column, host_val, link_val)]
    """
    host_names = set(host_map.keys())
    link_names = set(link_map.keys())

    missing_in_link = sorted(list(host_names - link_names), key=lambda s: s.lower())
    extra_in_link   = sorted(list(link_names - host_names), key=lambda s: s.lower())

    diffs = []
    cols = ("New", "Existing", "Demolished", "Temporary")
    for fname in sorted(host_names & link_names, key=lambda s: s.lower()):
        for colname in cols:
            hv = host_map[fname].get(colname, None)
            lv = link_map[fname].get(colname, None)
            # Only compare when both sides have a real value
            if hv is None or lv is None:
                continue
            if str(hv) != str(lv):
                diffs.append((fname, colname, hv, lv))
    return missing_in_link, extra_in_link, diffs


def log_action(action):
    """Log action to user JSON log file."""
    #____________________________________________________________________ IMPORTS
    import os, json, time

    from pyrevit import revit
    # from Snippets import _FunctionLogger as func_logger

    doc = revit.doc
    doc_path = doc.PathName or "<Untitled>"

    doc_title = doc.Title
    version_build = doc.Application.VersionBuild
    version_number = doc.Application.VersionNumber
    username = doc.Application.Username
    action = action

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

    # func_logger.write_json(dataEntry, filename=log_file)

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
    logcheck = False
    if not os.path.exists(log_file):
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        with open(log_file, 'w') as file:
            # create json structure
            file.write('{"action": []}')
        output_window.print_md("### **Created log file:** `{}`".format(log_file))

    with open(log_file,'r+') as file:
        file_data = json.load(file)
        if 'action' not in file_data:
            file_data['action'] = []
            file.seek(0)
            json.dump(file_data, file, indent = 4)

    try:
        write_json(dataEntry)
        logcheck = True
        output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
    except Exception as e:
        logcheck = False

    return dataEntry



output_window.print_md("Logging action: {}".format(log_action(action)))


#____________________________________________________________________ MAIN
# Host data
host_title = doc.Title
host_map = get_phase_filters_map(doc)

# Link data
link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
loaded_links = [(li, li.GetLinkDocument()) for li in link_instances if li.GetLinkDocument()]

if not loaded_links:
    TaskDialog.Show(__title__, "No loaded Revit links found in this model.")
    sys.exit()


# Header
output_window.print_md("# Phase Filter Comparison")
output_window.print_md("## **Host model:** `{}`".format(host_title))


### Host table
host_columns = ["Name", "New", "Existing", "Demolished", "Temporary"]
host_rows = [
    [n, sanitize(v["New"]), sanitize(v["Existing"]), sanitize(v["Demolished"]), sanitize(v["Temporary"])]
    for n, v in sorted(host_map.items(), key=lambda kv: kv[0].lower())
]
output_window.print_table(table_data=host_rows, columns=host_columns, title="Host Phase Filters ({})".format(len(host_map)))

### Per-link
for li, ldoc in loaded_links:
    link_name = ldoc.Title
    output_window.print_md("---")
    output_window.print_md("## Link: `{}`".format(link_name))

    link_map = get_phase_filters_map(ldoc)
    output_window.print_md("**Link Phase Filters ({}):**".format(len(link_map)))

    # Compute diffs early so we can tag mismatched filters with ❌ in the link table
    missing, extra, diffs = compare_phase_filters(host_map, link_map)
    mismatch_names = {fname for (fname, _, _, _) in diffs}

    # Link table (append ❌ to names that mismatch host settings)
    host_columns = ["Name", "New", "Existing", "Demolished", "Temporary"]
    link_rows = []
    for n, v in sorted(link_map.items(), key=lambda kv: kv[0].lower()):
        display_name = u"{} ❌".format(n) if n in mismatch_names else n
        link_rows.append([
            display_name,
            sanitize(v["New"]), sanitize(v["Existing"]), sanitize(v["Demolished"]), sanitize(v["Temporary"])
        ])
    output_window.print_table(table_data=link_rows, columns=host_columns, title="Link Filters")
    if mismatch_names:
        output_window.print_md("_Legend: ❌ = settings differ from host for this filter._")

    # Lists + mismatched table (reuse the diffs we already computed)
    if missing:
        output_window.print_md("### Missing in Link (present in Host, absent in Link)")
        output_window.print_md(md_list(missing))

    if extra:
        output_window.print_md("### Extra in Link (present in Link, absent in Host)")
        output_window.print_md(md_list(extra))

    if diffs:
        output_window.print_md("### Setting Mismatches")
        diff_columns = ["Filter", "Column", "Host", "Link"]
        diff_rows = [
            [fname, col, sanitize(hv), sanitize(lv)]
            for (fname, col, hv, lv) in diffs
        ]
        output_window.print_table(table_data=diff_rows, columns=diff_columns, title="Mismatched Properties")

    if (not missing) and (not extra) and (not diffs):
        output_window.print_md("_No differences found for this link._")




#_____________________________________________________________________ 🏃‍➡️ RUN 


