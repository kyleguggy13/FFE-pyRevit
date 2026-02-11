# -*- coding: utf-8 -*-
__title__     = "Sheet \nCounter"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 02.11.2026
______________________________________________________________
Description:
-> Enumerates Sheets for "SHEET ## OUT OF ###" on the Titleblock.
______________________________________________________________
How-to:
-> Press Button and select option.
-> Check results
______________________________________________________________
Last update:
- [02.09.2026] - v0.1 BETA RELEASE
- [02.11.2026] - v1.1 RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)
from System import String


#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory


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
action = "Sheets_Counter"

output_window = output.get_output()
"""Output window for displaying results."""



#____________________________________________________________________ FUNCTIONS

"""
pyRevit | FFE - Sheet Counter (Dynamo recreation)

- Collect sheets from current model (and optionally all links)
- Filter: Appears In Sheet List + not placeholder
- Sort by Sheet Number
- Create sequence 1..N for current model sheets
- Write total count to Project Information parameter
- Write per-sheet index to a sheet parameter
"""



import re
from pyrevit import revit, DB, forms


# ----------------------- CONFIG (EDIT THESE) -----------------------

DISCIPLINE_PARAM = "FFE_Sheet_Discipline Index"
ORDER_PARAM      = "FFE_Sheet_Order"

# Optional outputs (set to None to disable)
# PROJECT_INFO_TOTAL_PARAM = "FFE_Sheet_Counter_Total"   # ProjectInformation param
# SHEET_INDEX_PARAM        = "FFE_Sheet_Index"           # Sheet param for overall 1..N
PROJECT_INFO_TOTAL_PARAM = "Total Number of Sheets"     # ProjectInformation param
SHEET_INDEX_PARAM        = "Number of Sheets"           # Sheet param for overall 1..N

WRITE_PARAMETERS = True  # False = read/count only

# Filters (match common Dynamo "sheet list" expectations)
EXCLUDE_PLACEHOLDERS = True
REQUIRE_APPEARS_IN_SHEET_LIST = True


# ----------------------- PARAMETER VALIDATION -----------------------

def parameter_exists_on_category(doc, param_name, bic):
    """Check if a parameter is bound to a category."""
    bindings = doc.ParameterBindings
    it = bindings.ForwardIterator()
    it.Reset()

    print("bic: {}, {}".format(bic, type(bic)))

    while it.MoveNext():
        definition = it.Key
        binding = it.Current
        print("definition.Name: {}, {}".format(definition.Name, type(definition.Name)))

        if definition.Name == param_name:
            return True

    return False


def project_info_has_param(doc, param_name):
    p = doc.ProjectInformation.LookupParameter(param_name)
    return p is not None


def validate_required_parameters(doc):
    missing = []

    # Required sheet parameters
    required_sheet_params = [
        DISCIPLINE_PARAM,
        ORDER_PARAM
    ]

    if WRITE_PARAMETERS and SHEET_INDEX_PARAM:
        required_sheet_params.append(SHEET_INDEX_PARAM)

    for pname in required_sheet_params:
        if not parameter_exists_on_category(doc, pname, DB.BuiltInCategory.OST_Sheets):
            missing.append("Sheet parameter missing: '{}'".format(pname))

    # Required Project Information parameter
    if WRITE_PARAMETERS and PROJECT_INFO_TOTAL_PARAM:
        if not project_info_has_param(doc, PROJECT_INFO_TOTAL_PARAM):
            missing.append("Project Information parameter missing: '{}'".format(PROJECT_INFO_TOTAL_PARAM))

    if missing:
        forms.alert(
            "Required parameters are missing:\n\n" + "\n".join(missing) +
            "\n\nScript cancelled.",
            title="FFE Sheet Counter – Missing Parameters",
            warn_icon=True
        )
        return False

    return True


# ----------------------- FUNCTIONS -----------------------


def natural_key(text):
    parts = re.split(r"(\d+)", text or "")
    key = []
    for part in parts:
        key.append(int(part) if part.isdigit() else part.lower())
    return key


def parse_int_maybe(val):
    """Parse int from parameter value that might be int, string '00', ' 11 ', etc."""
    if val is None:
        return None
    if isinstance(val, int):
        return val
    s = str(val).strip()
    # keep only leading sign + digits
    m = re.match(r"^-?\d+", s)
    return int(m.group(0)) if m else None


def get_param_as_python(elem, param_name):
    p = elem.LookupParameter(param_name)
    if not p:
        return None
    st = p.StorageType
    if st == DB.StorageType.Integer:
        return p.AsInteger()
    if st == DB.StorageType.String:
        return p.AsString()
    if st == DB.StorageType.Double:
        return p.AsDouble()
    return None


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


def collect_sheets(doc):
    sheets = (DB.FilteredElementCollector(doc)
              .OfClass(DB.ViewSheet)
              .WhereElementIsNotElementType()
              .ToElements())

    out = []
    for s in sheets:
        if EXCLUDE_PLACEHOLDERS and getattr(s, "IsPlaceholder", False):
            continue

        if REQUIRE_APPEARS_IN_SHEET_LIST:
            p = s.get_Parameter(DB.BuiltInParameter.SHEET_SCHEDULED)
            if p and p.StorageType == DB.StorageType.Integer and p.AsInteger() != 1:
                continue

        out.append(s)
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



# ----------------------- MAIN -----------------------
def main():
    host_doc = revit.doc

    # Validate parameters BEFORE doing anything else
    if not validate_required_parameters(host_doc):
        return

    scope = forms.CommandSwitchWindow.show(
        ["Current model only", "Current + all links (count links too)"],
        message="Sheet Counter Scope",
    )
    if not scope:
        return

    include_links = scope.startswith("Current + all links")

    # ---- Build MASTER list ----
    host_sheets = collect_sheets(host_doc)

    master = []  # list of tuples: (doc, sheet, is_host)
    for s in host_sheets:
        master.append((host_doc, s, True))


    # Linked sheets count only (cannot write to link docs)
    link_breakdown = []
    if include_links:
        for ldoc in get_link_documents(host_doc):
            lsheets = collect_sheets(ldoc)
            link_breakdown.append((ldoc.Title, len(lsheets)))
            for s in lsheets:
                master.append((ldoc, s, False))

    # Sort master by the sheet parameters (same sort applied to host+links)
    master_sorted = sorted(master, key=lambda t: sheet_sort_key(t[1]))

    # Global sequence over MASTER list
    # We build a lookup for host sheets: sheet.UniqueId -> global_index
    host_index_by_uid = {}
    for idx, (d, s, is_host) in enumerate(master_sorted, start=1):
        if is_host:
            host_index_by_uid[s.UniqueId] = idx

    total_all = len(master_sorted)

    # ---- Write (host only) ----
    write_notes = []
    if WRITE_PARAMETERS:
        with revit.Transaction("FFE Sheet Counter (Host+Links)"):
            if PROJECT_INFO_TOTAL_PARAM:
                ok, err = set_param_value(host_doc.ProjectInformation, PROJECT_INFO_TOTAL_PARAM, total_all)
                if not ok:
                    write_notes.append("ProjectInfo: {}".format(err))

            if SHEET_INDEX_PARAM:
                # write sequence to host sheets using master index
                for s in host_sheets:
                    global_i = host_index_by_uid.get(s.UniqueId)
                    if global_i is None:
                        continue
                    ok2, err2 = set_param_value(s, SHEET_INDEX_PARAM, global_i)
                    if not ok2 and len(write_notes) < 10:
                        write_notes.append("Sheet {}: {}".format(s.SheetNumber, err2))


    # ---- Preview ----
    preview_n = min(15, len(master_sorted))
    lines = []
    for i in range(preview_n):
        d, s, is_host = master_sorted[i]
        di = get_param_as_python(s, DISCIPLINE_PARAM)
        od = get_param_as_python(s, ORDER_PARAM)
        tag = "HOST" if is_host else "LINK"
        lines.append("{:>3} | {:4} | DI={} | ORD={} | {} — {}".format(
            i + 1, tag, di, od, s.SheetNumber, s.Name
        ))
    preview = "\n".join(lines) or "(no sheets found)"

    link_lines = "\n".join(["- {}: {}".format(t, c) for t, c in link_breakdown]) or "(none)"

    msg = []
    msg.append("Mode: {}".format("Host + Links" if include_links else "Host only"))
    msg.append("Host sheets: {}".format(len(host_sheets)))
    msg.append("Master total (host+links): {}".format(total_all))
    msg.append("")
    msg.append("Sorted by: ({}, {})".format(DISCIPLINE_PARAM, ORDER_PARAM))
    msg.append("")
    msg.append("Master preview (top {}):".format(preview_n))
    msg.append(preview)
    msg.append("")
    msg.append("Link breakdown:")
    msg.append(link_lines)

    if write_notes:
        msg.append("")
        msg.append("Write notes (first few):")
        msg.extend(["- " + n for n in write_notes])

    forms.alert("\n".join(msg), title="FFE Sheet Counter (Host+Links)", warn_icon=False)


if __name__ == "__main__":
    main()


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