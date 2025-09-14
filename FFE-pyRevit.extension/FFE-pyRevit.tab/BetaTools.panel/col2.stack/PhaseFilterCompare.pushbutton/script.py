# -*- coding: utf-8 -*-
__title__     = "Phase Filter\nComparison"
__version__   = 'Version = 0.1'
__doc__       = """Version = 0.1
Date    = 09.13.2025
# ______________________________________________________________
# Description:
# -> 
# -> 
# 
# ______________________________________________________________
# How-to:
#
# -> Click the button
# -> 
# -> 
# -> 
#   
# ______________________________________________________________
# Last update:
# - [09.14.2025] - First release
# ______________________________________________________________
Author: Kyle Guggenheim"""

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


#____________________________________________________________________ TO-DO
# - [ ] Add error handling for missing parameters
# - [ ] Add option to select different arrowhead types
# - [ ] Check for duplicate numbers or texts
# - [ ] Add element links to the output table for easier identification
# - [ ] Add a confirmation dialog before renaming types 



#____________________________________________________________________ HELPERS
def md_list(items):
    """Render a Python list as a single markdown string (never pass lists to print_md)."""
    try:
        lines = [u"- {}".format(u"{}".format(it)) for it in items]
    except Exception:
        lines = [u"- {}".format(str(it)) for it in items]
    return u"\n".join(lines)


def format_bool(v):
    if v is True:  return "True"
    if v is False: return "False"
    return "N/A"


def get_phase_filters_map(rvt_doc):
    """Return dict: { filter_name: {ShowNew, ShowExisting, ShowDemolished, ShowTemporary} }.
       Properties are set to None if unavailable in this Revit version.
    """
    data = OrderedDict()
    col = FilteredElementCollector(rvt_doc).OfClass(PhaseFilter)
    for pf in col:
        name = pf.Name
        entry = {"ShowNew": None, "ShowExisting": None, "ShowDemolished": None, "ShowTemporary": None}
        for prop in ("ShowNew", "ShowExisting", "ShowDemolished", "ShowTemporary"):
            try:
                entry[prop] = getattr(pf, prop)
            except Exception:
                pass
        data[name] = entry
    return data


def compare_phase_filters(host_map, link_map):
    """Return (missing_in_link, extra_in_link, diffs).
       diffs: [(filter_name, prop, host_val, link_val)]
    """
    host_names = set(host_map.keys())
    link_names = set(link_map.keys())

    missing_in_link = sorted(list(host_names - link_names))
    extra_in_link   = sorted(list(link_names - host_names))

    diffs = []
    common = host_names & link_names
    props = ("ShowNew", "ShowExisting", "ShowDemolished", "ShowTemporary")
    for fname in sorted(common):
        for p in props:
            hv = host_map[fname].get(p, None)
            lv = link_map[fname].get(p, None)
            if (hv is not None) and (lv is not None) and (hv != lv):
                diffs.append((fname, p, hv, lv))
    return missing_in_link, extra_in_link, diffs

#____________________________________________________________________ MAIN

host_title = doc.Title
host_map = get_phase_filters_map(doc)

link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
loaded_links = [(li, li.GetLinkDocument()) for li in link_instances if li.GetLinkDocument()]

if not loaded_links:
    TaskDialog.Show(__title__, "No loaded Revit links found in this model.")
    sys.exit()

# Header
output_window = output.get_output()
output_window.print_md("# Phase Filter Comparison")
output_window.print_md("**Host model:** `{}`".format(host_title))
output_window.print_md("**Host Phase Filters ({}):**".format(len(host_map)))

# --- Host table (use columns= for headers) ---
host_columns = ["Name", "ShowNew", "ShowExisting", "ShowDemolished", "ShowTemporary"]
host_rows = [
    [n, format_bool(v["ShowNew"]), format_bool(v["ShowExisting"]),
     format_bool(v["ShowDemolished"]), format_bool(v["ShowTemporary"])]
    for n, v in host_map.items()
]
output_window.print_table(table_data=host_rows, columns=host_columns, title="Host Filters")


for li, ldoc in loaded_links:
    link_name = ldoc.Title
    output_window.print_md("---")
    output_window.print_md("## Link: `{}`".format(link_name))

    link_map = get_phase_filters_map(ldoc)
    output_window.print_md("**Link Phase Filters ({}):**".format(len(link_map)))

    # Differences
    missing, extra, diffs = compare_phase_filters(host_map, link_map)

    if missing:
        output_window.print_md("### Missing in Link (present in Host, absent in Link)")
        output_window.print_md(md_list(missing))

    if extra:
        output_window.print_md("### Extra in Link (present in Link, absent in Host)")
        output_window.print_md(md_list(extra))

    if diffs:
        output_window.print_md("### Setting Mismatches")
        output_window.print_table(
            table_data=[["Filter", "Property", "Host", "Link"]] + [
                [fname, prop, format_bool(hv), format_bool(lv)]
                for (fname, prop, hv, lv) in diffs
            ],
            title="Mismatched Properties"
        )

    if (not missing) and (not extra) and (not diffs):
        output_window.print_md("_No differences found for this link._")



#_____________________________________________________________________ 🏃‍➡️ RUN 


