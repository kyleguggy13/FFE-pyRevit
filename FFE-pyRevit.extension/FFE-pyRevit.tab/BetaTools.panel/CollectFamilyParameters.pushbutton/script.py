# -*- coding: utf-8 -*-
__title__     = "Collect Family \nParameters"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 02.19.2026
______________________________________________________________
Description:
-> Export all loadable families and their Family Parameters to JSON.
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [02.09.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


import clr
import os
import json

from System import Guid

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import FilteredElementCollector, Family
from Autodesk.Revit.DB import InternalDefinition, BuiltInParameter

doc = __revit__.ActiveUIDocument.Document


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def safe_str(x):
    try:
        if x is None:
            return None
        return str(x)
    except:
        return None


def get_param_type_string(fp):
    d = fp.Definition

    try:
        dt = d.GetDataType()
        if dt:
            return safe_str(dt.TypeId) if hasattr(dt, "TypeId") else safe_str(dt)
    except:
        pass

    try:
        return safe_str(d.ParameterType)
    except:
        pass

    return None


def get_formula(fm, fp):
    try:
        return fm.GetFormula(fp)
    except:
        return None


def get_shared_guid(fp):
    try:
        if fp.IsShared:
            return safe_str(fp.GUID)
    except:
        pass
    return None


def try_get_built_in_parameter(defn):
    """
    Returns (is_built_in: bool, bip_name: str|None, bip_int: int|None)
    """
    try:
        idef = defn if isinstance(defn, InternalDefinition) else None
        if idef is None:
            return (False)

        bip = idef.BuiltInParameter  # BuiltInParameter enum
        if bip == BuiltInParameter.INVALID:
            return (False)

        return (True)
    except:
        return (False)


def read_family_parameters(fam_doc):
    fm = fam_doc.FamilyManager
    output = []

    for fp in fm.Parameters:
        try:
            d = fp.Definition

            is_bip = try_get_built_in_parameter(d)

            param_data = {
                "name": safe_str(d.Name),
                "is_instance": bool(fp.IsInstance),
                # "is_reporting": bool(fp.IsReporting),
                "is_shared": bool(fp.IsShared),
                "group": safe_str(d.ParameterGroup),
                "param_value_type": get_param_type_string(fp),
                # "formula": get_formula(fm, fp),
                "guid": get_shared_guid(fp),
                "is_built_in": bool(is_bip),
            }

            output.append(param_data)

        except:
            continue

    output.sort(key=lambda x: (x["group"] or "", x["name"] or ""))
    return output


# ------------------------------------------------------------
# Main Collector
# ------------------------------------------------------------

def collect_family_data(document):

    result = {}

    families = (
        FilteredElementCollector(document)
        .OfClass(Family)
        .ToElements()
    )

    for fam in families:

        if fam.IsInPlace:
            continue

        try:
            fam_name = safe_str(fam.Name)

            fam_doc = document.EditFamily(fam)
            try:
                params = read_family_parameters(fam_doc)
                result[fam_name] = params
            finally:
                fam_doc.Close(False)

        except:
            continue

    return dict(sorted(result.items(), key=lambda kv: kv[0].lower()))


# ------------------------------------------------------------
# Run + Export
# ------------------------------------------------------------

family_data = collect_family_data(doc)

downloads = os.path.join(os.path.expanduser("~"), "Downloads")
file_path = os.path.join(downloads, "Revit_Family_Parameters.json")

with open(file_path, "w") as f:
    json.dump(family_data, f, indent=4)

print("Export complete.")
print("Families exported: {}".format(len(family_data)))
print("Saved to:")
print(file_path)