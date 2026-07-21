# -*- coding: utf-8 -*-
__title__     = "HVAC Calculated Values"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 02.09.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [02.11.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
from collections import defaultdict
import time


#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import ElementCategoryFilter, ElementId, FamilyInstance


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

log_status = ""
action = "HVAC Calculated Values"

output_window = output.get_output()
"""Output window for displaying results."""



#____________________________________________________________________ FUNCTIONS

"""
pyRevit | HVAC Calculcated Values

- Collect Spaces from current model
- Filter: FFE_Space_Air Handling Unit_Mark (Or other parameter)
- Sort by Space Number, Name
- Collect Parameters
- Calculate intended values
- Write write calculated values to parameters
"""



import re
from pyrevit import revit, DB, forms


# ----------------------- CONFIG (EDIT THESE) -----------------------

# Parameters
AHU_MARK = "FFE_Space_Air Handling Unit_Mark"

# Options
WRITE_PARAMETERS = True  # False = read/count only

# Filters
EXCLUDE_WITHOUT_VALUE = True
REQUIRE_VALUE_IN_AHU_MARK = True

#_________________________________________________________ FUNCTIONS


def set_param_value(elem, param_name, value):
    p = elem.LookupParameter(param_name)
    if not p or p.IsReadOnly:
        return False, "Missing or read-only param: {}".format(param_name)

    try:
        st = p.StorageType
        if st == DB.StorageType.Integer:
            p.Set(int(value))
        elif st == DB.StorageType.String:
            p.Set("" if value is None else str(value))
        else:
            # best-effort
            p.Set(str(value))
        return True, None
    except Exception as ex:
        return False, str(ex)


def collect_spaces(doc):
    # sheets = (DB.FilteredElementCollector(doc)
    #           .OfClass(DB.ViewSheet)
    #           .WhereElementIsNotElementType()
    #           .ToElements())

    out = []
    # for s in sheets:
    #     if EXCLUDE_PLACEHOLDERS and getattr(s, "IsPlaceholder", False):
    #         continue

    #     if REQUIRE_APPEARS_IN_SHEET_LIST:
    #         p = s.get_Parameter(DB.BuiltInParameter.SHEET_SCHEDULED)
    #         if p and p.StorageType == DB.StorageType.Integer and p.AsInteger() != 1:
    #             continue

    #     out.append(s)
    return out


def get_link_documents(host_doc):
    link_docs = []
    instances = (DB.FilteredElementCollector(host_doc)
                 .OfClass(DB.RevitLinkInstance)
                 .WhereElementIsNotElementType()
                 .ToElements())
    for inst in instances:
        try:
            ldoc = inst.GetLinkDocument()
            if ldoc:
                link_docs.append(ldoc)
        except Exception:
            pass
    return link_docs


def sheet_sort_key(sheet):
    di_raw = sheet.LookupParameter(DISCIPLINE_PARAM).AsString()
    od_raw = sheet.LookupParameter(ORDER_PARAM).AsString()

    # Put missing values at the end (very large)
    di_key = di_raw if di_raw is not None else ""
    od_key = od_raw if od_raw is not None else ""

    return (di_key, od_key, natural_key(sheet.SheetNumber))



##########################
##########################
### NEED TO UPDATE SO THAT COUNT STARTS AT APPROPRIATE LOCATION ###

def main():
    doc = revit.doc

    scope = forms.CommandSwitchWindow.show(
        ["Current model only", "Current + all links (count links too)"],
        message="Sheet Counter Scope",
    )
    if not scope:
        return

    include_links = scope.startswith("Current + all links")

    # Current model sheets sorted by (Discipline Index, Order)
    current_sheets = collect_sheets(doc)
    current_sorted = sorted(current_sheets, key=sheet_sort_key)
    seq = list(range(1, len(current_sorted) + 1))

    # Linked sheets count only (cannot write to link docs)
    linked_total = 0
    link_breakdown = []
    if include_links:
        for ldoc in get_link_documents(doc):
            lsheets = collect_sheets(ldoc)
            linked_total += len(lsheets)
            link_breakdown.append((ldoc.Title, len(lsheets)))

    total_all = len(current_sorted) + (linked_total if include_links else 0)

    write_notes = []
    if WRITE_PARAMETERS:
        with revit.Transaction("FFE Sheet Counter (Discipline+Order)"):
            if PROJECT_INFO_TOTAL_PARAM:
                ok, err = set_param_value(doc.ProjectInformation, PROJECT_INFO_TOTAL_PARAM, total_all)
                if not ok:
                    write_notes.append("ProjectInfo: {}".format(err))

            if SHEET_INDEX_PARAM:
                for s, i in zip(current_sorted, seq):
                    ok2, err2 = set_param_value(s, SHEET_INDEX_PARAM, i)
                    if not ok2 and len(write_notes) < 8:
                        write_notes.append("Sheet {}: {}".format(s.SheetNumber, err2))

    # Preview
    # print(current_sorted)
    preview_n = min(12, len(current_sorted))
    lines = []
    for s, i in zip(current_sorted[:preview_n], seq[:preview_n]):
        di = s.LookupParameter(DISCIPLINE_PARAM).AsString()
        od = s.LookupParameter(ORDER_PARAM).AsString()
        lines.append("{:>3} | DI={} | ORD={} | {}  —  {}".format(
            i, di, od, s.SheetNumber, s.Name
        ))
    preview = "\n".join(lines) or "(no sheets found)"

    link_lines = "\n".join(["- {}: {}".format(t, c) for t, c in link_breakdown]) or "(none)"

    msg = []
    msg.append("Current model sheets: {}".format(len(current_sorted)))
    msg.append("Linked model sheets:  {}".format(linked_total if include_links else 0))
    msg.append("TOTAL: {}".format(total_all))
    msg.append("")
    msg.append("Sorted by: ({}, {})".format(DISCIPLINE_PARAM, ORDER_PARAM))
    msg.append("")
    msg.append("Top {} preview:".format(preview_n))
    msg.append(preview)
    msg.append("")
    msg.append("Link breakdown:")
    msg.append(link_lines)

    if write_notes:
        msg.append("")
        msg.append("Write notes (first few):")
        msg.extend(["- " + n for n in write_notes])

    forms.alert("\n".join(msg), title="FFE Sheet Counter (Discipline+Order)", warn_icon=False)


if __name__ == "__main__":
    main()