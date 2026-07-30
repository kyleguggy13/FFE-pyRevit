# -*- coding: utf-8 -*-
__title__ = "CAD Tracker"
__doc__     = """Version = V1.4
Date    = 2026.07.30
________________________________________________________________
Last Updates:
- [2026.07.30] V1.4 Added workset, hosted level, and offset from level to report
- [2026.07.19] V1.3 Improved error handling inside functions
- [2026.07.19] V1.2 Optimize removed viewtype filtering since cad.OwnerViewId already returns a valid view type
- [2026.07.19] V1.1 Optimize scanning view from FilteredElementCollector to utilizing cad.OwnerViewId
- [2026.07.18] V1.0 Rewrite code, removed FilteredElementCollector inside the loop
- [2026.07.12] V0.2 Stress test code (Code runs too slow and very inefficient)
- [2026.07.11] v0.1 Proof of concept 
________________________________________________________________
Author: Frederick"""


#IMPORTS
from Autodesk.Revit.DB import *
from pyrevit import forms, script
from collections import defaultdict


#FUNCTIONS
def has_valid_category(cad):
    """Returns True if cad has a Category otherwise False"""
    return bool(cad.Category)
    

def get_creator(doc, cad, is_workshared):
    if not is_workshared:
        return "Not workshared"
    try:
        return WorksharingUtils.GetWorksharingTooltipInfo(doc, cad.Id).Creator
    except Exception:
        return "Unknown"


def get_workset_name(doc, cad, is_workshared):
    """Returns the name of the workset that owns the CAD instance."""
    if not is_workshared:
        return "Not workshared"
    try:
        workset = doc.GetWorksetTable().GetWorkset(cad.WorksetId)
        return workset.Name if workset else "Unknown"
    except Exception:
        return "Unknown"


def get_parameter_value(element, built_in_parameter, doc=None):
    """Returns a parameter's display value, resolving ElementIds when needed."""
    try:
        parameter = element.get_Parameter(built_in_parameter)
        if not parameter or not parameter.HasValue:
            return "N/A"

        value = parameter.AsValueString()
        if value:
            return value

        if parameter.StorageType == StorageType.ElementId and doc:
            referenced_element = doc.GetElement(parameter.AsElementId())
            if referenced_element:
                return referenced_element.Name

        value = parameter.AsString()
        return value if value else "N/A"
    except Exception:
        return "Unknown"


def get_cad_properties(cad, doc, is_workshared, view = None):
    """Returns the properties of the cad data"""

    cad_link   = output.linkify(cad.Id)
    name       = cad.Category.Name if has_valid_category(cad) else cad.Parameter[BuiltInParameter.IMPORT_SYMBOL_NAME].AsString() + " [ERROR: dwg not loaded]"
    view_name  = view.Name if view else "Global/3D"

    if view:
        status = "Linked/2D" if cad.IsLinked else "Import/2D"
    else:
        status = "Linked/3D" if cad.IsLinked else "Import/3D"
    
    workset     = get_workset_name(doc=doc, cad=cad, is_workshared=is_workshared)
    hosted_level = get_parameter_value(cad, BuiltInParameter.IMPORT_BASE_LEVEL, doc)
    level_offset = get_parameter_value(cad, BuiltInParameter.IMPORT_BASE_LEVEL_OFFSET)
    owner       = get_creator(doc = doc, cad = cad, is_workshared = is_workshared)

    visibility = "Visible"
    if not view:
        visibility = "3D object"
    else:
        try:
            category_hidden = has_valid_category(cad) and view.GetCategoryHidden(cad.Category.Id)
            if cad.IsHidden(view) or category_hidden:
                visibility = "Hidden"
        except Exception:
            visibility = "Unknown"

    return [
        cad_link,
        name,
        view_name,
        visibility,
        status,
        workset,
        hosted_level,
        level_offset,
        owner,
    ]


#VARIABLES
doc    = __revit__.ActiveUIDocument.Document
output = script.get_output()

#Check if active document is workshared
is_workshared =  doc.IsWorkshared

# Get all cad links and import in the project
all_cads = FilteredElementCollector(doc).OfClass(ImportInstance)

if all_cads.GetElementCount() == 0:
    forms.alert("There are no linked cads in the project", exitscript=True, title=__title__)

ThreeD_cads   = []
list_of_cads = defaultdict(list)

for cad in all_cads:
    if cad.OwnerViewId != ElementId.InvalidElementId:
        list_of_cads[cad.OwnerViewId].append(cad)
    else:
        ThreeD_cads.append(cad)

view_cache = {}
data_report = []

for view_id, cad_in_view in list_of_cads.items():
    if view_id not in view_cache:
        view_cache[view_id] = doc.GetElement(view_id)
    
    view = view_cache[view_id]

    if not view or view.IsTemplate:
        continue

    for cad in cad_in_view:
        data_report.append(get_cad_properties(cad= cad, doc= doc, is_workshared= is_workshared, view= view))

data_report.sort(key=lambda x: (x[1], x[2]))

if ThreeD_cads:
    for cad in ThreeD_cads:
        threeD_cad_properties = get_cad_properties(cad= cad, doc= doc, is_workshared= is_workshared)
        data_report.append(threeD_cad_properties)

output.print_md("##CAD LINKS REPORT:")
output.print_table(
    table_data=data_report,
    columns=[
        "CAD",
        "NAME",
        "VIEW",
        "VISIBILITY",
        "STATUS",
        "WORKSET",
        "HOSTED LEVEL",
        "OFFSET FROM LEVEL",
        "CREATED BY",
    ],
)
