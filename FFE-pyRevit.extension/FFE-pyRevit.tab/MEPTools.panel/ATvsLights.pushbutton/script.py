# -*- coding: utf-8 -*-
__title__     = "AT vs Lights"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 03.03.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [03.03.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""

# #____________________________________________________________________ IMPORTS (SYSTEM)
# from System import String
# from collections import defaultdict
# import time


# #____________________________________________________________________ IMPORTS (AUTODESK)
# import sys
# import clr
# clr.AddReference("System")
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from Autodesk.Revit.DB import FilteredElementCollector, Mechanical, Family, BuiltInParameter, ElementType, UnitTypeId
# from Autodesk.Revit.DB import BuiltInCategory, ElementCategoryFilter, ElementId, FamilyInstance
# from Autodesk.Revit.DB.ExtensibleStorage import Schema


# #____________________________________________________________________ IMPORTS (PYREVIT)
# from pyrevit import revit, DB, UI, script
# from pyrevit.script import output
# from pyrevit import forms

# #____________________________________________________________________ VARIABLES
# app         = __revit__.Application
# uidoc       = __revit__.ActiveUIDocument
# doc         = __revit__.ActiveUIDocument.Document   #type: Document
# selection   = uidoc.Selection                       #type: Selection

# log_status = ""
# action = "Duct Network Summary"

# output_window = output.get_output()
# """Output window for displaying results."""






from pyrevit import revit
from pyrevit.script import output

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    BoundingBoxIntersectsFilter,
    Outline,
    XYZ
)

doc = revit.doc
view = doc.ActiveView
out = output.get_output()
out.set_title("Air Terminals vs Lighting Fixtures - Active View")

# ------------------------------------------------------
# Settings
# ------------------------------------------------------
# Revit internal units (feet)
# 0.10 ft ≈ 1.2 inches
TOL_FT = -0.50


# ------------------------------------------------------
# Utility Functions
# ------------------------------------------------------

def get_bbox(elem):
    try:
        return elem.get_BoundingBox(view)
    except:
        return None


def bbox_intersects(bbA, bbB):
    if not bbA or not bbB:
        return False

    a0, a1 = bbA.Min, bbA.Max
    b0, b1 = bbB.Min, bbB.Max

    return (a0.X <= b1.X and a1.X >= b0.X and
            a0.Y <= b1.Y and a1.Y >= b0.Y and
            a0.Z <= b1.Z and a1.Z >= b0.Z)


def bbox_to_outline(bb, tol):
    if not bb:
        return None

    mn = bb.Min
    mx = bb.Max

    min_pt = XYZ(mn.X - tol, mn.Y - tol, mn.Z - tol)
    max_pt = XYZ(mx.X + tol, mx.Y + tol, mx.Z + tol)

    return Outline(min_pt, max_pt)


def elem_name(e):
    try:
        sym = getattr(e, "Symbol", None)
        if sym:
            fam = getattr(sym, "Family", None)
            famname = fam.Name if fam else ""
            return "{} : {}".format(famname, sym.Name)
    except:
        pass
    return e.Name if hasattr(e, "Name") else "<Unnamed>"


# ------------------------------------------------------
# Collect Elements (ACTIVE VIEW ONLY)
# ------------------------------------------------------

air_terms = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
)

lights = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_LightingFixtures)
    .WhereElementIsNotElementType()
)

out.print_md("### Active View: **{}**".format(view.Name))
out.print_md("- Air Terminals: **{}**".format(len(air_terms)))
out.print_md("- Lighting Fixtures: **{}**".format(len(lights)))

if not air_terms or not lights:
    out.print_md("\nNo elements found in one or both categories.")
    raise SystemExit


# Pre-cache lighting bounding boxes
light_bboxes = {}
for l in lights:
    bb = get_bbox(l)
    if bb:
        light_bboxes[l.Id.ToString()] = bb


# ------------------------------------------------------
# Overlap Detection
# ------------------------------------------------------

results = []
seen = set()

for at in air_terms:

    at_bb = get_bbox(at)
    if not at_bb:
        continue

    outline = bbox_to_outline(at_bb, TOL_FT)
    if not outline:
        continue

    bb_filter = BoundingBoxIntersectsFilter(outline)

    candidates = (
        FilteredElementCollector(doc, view.Id)
        .OfCategory(BuiltInCategory.OST_LightingFixtures)
        .WherePasses(bb_filter)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for l in candidates:
        l_bb = light_bboxes.get(l.Id.ToString())
        if not l_bb:
            continue

        if bbox_intersects(at_bb, l_bb):
            key = (at.Id.ToString(), l.Id.ToString())
            if key in seen:
                continue
            seen.add(key)

            results.append([
                out.linkify(at.Id),
                elem_name(at),
                out.linkify(l.Id),
                elem_name(l)
            ])


# ------------------------------------------------------
# Output
# ------------------------------------------------------

out.print_md("\n### Overlaps Found: **{}**".format(len(results)))

if results:
    out.print_table(
        table_data=results,
        columns=["Air Terminal Id", "Air Terminal", "Light Fixture Id", "Light Fixture"],
        title="Overlapping (Bounding Box) Pairs"
    )
else:
    out.print_md("No overlaps detected in active view.")




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

# log_action(action, log_status)