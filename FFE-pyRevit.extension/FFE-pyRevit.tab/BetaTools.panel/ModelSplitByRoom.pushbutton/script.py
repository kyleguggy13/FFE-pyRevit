# -*- coding: utf-8 -*-
__title__     = "Model Split By Room"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 12.03.2025
___________________________________________________________________
Description:
Split current model into a per-SEPS-Code layout model.

Workflow:
1. Assumes a project/shared parameter "SEPS Code" exists on:
   - Rooms
   - Views (including schedules)
   - Sheets
   - (Optionally) model elements you want to tag per layout
2. Reads all distinct SEPS Codes from Rooms.
3. Lets user pick one SEPS Code.
4. Prompts user for a target .rvt file name.
5. Saves current document as that file (Save As).
6. In the newly saved model, deletes all elements that DO NOT belong
   to that SEPS layout.

Belonging to the layout is defined as:
- element has "SEPS Code" == selected code, OR
- element's bounding box intersects the union bounding box of
  all Rooms whose "SEPS Code" == selected code.
___________________________________________________________________
How-to:
-> Click on the button

___________________________________________________________________
Last update:
- [12.01.2025] - v0.1 BETA RELEASE

___________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, forms, script


#____________________________________________________________________ CONSTANTS
# Name of the shared / project parameter used for grouping.
SEPS_PARAM_NAME = "SEPS Code"

# Model categories to prune based on SEPS layout.
# Extend this list with any categories that are layout-specific.
MODEL_INSTANCE_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Furniture,
    DB.BuiltInCategory.OST_PlumbingFixtures,
    DB.BuiltInCategory.OST_SpecialityEquipment,
    DB.BuiltInCategory.OST_Casework,
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_MechanicalEquipment,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_MedicalEquipment,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_IOSModelGroups,
    DB.BuiltInCategory.OST_Lines,
    DB.BuiltInCategory.OST_VolumeOfInterest
]

# Use room bounding boxes to keep geometry that sits inside the layout area
# even if the elements do not have a SEPS Code parameter.
USE_ROOM_BBOX_FILTER = True
# Buffer around room bounding box, in feet
ROOM_BBOX_BUFFER_FT = 1.0


#____________________________________________________________________ VARIABLES
output = script.get_output()


#____________________________________________________________________ FUNCTIONS

def get_param_str(elem, param_name):
    """Return the string value of a parameter by name, or None."""
    if not elem:
        return None
    param = elem.LookupParameter(param_name)
    if param and param.HasValue:
        try:
            return param.AsString()
        except Exception:
            try:
                return param.AsValueString()
            except Exception:
                return None
    return None


def element_has_seps(elem, seps_code):
    """True if element's SEPS Code param matches the given code."""
    val = get_param_str(elem, SEPS_PARAM_NAME)
    if not val:
        return False
    return val.strip() == seps_code.strip()


# def collect_seps_codes_from_rooms(doc):
#     """Collect distinct non-empty SEPS Codes from Rooms."""
#     rooms = (DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())
#     keys = set()
#     for r in rooms:
#         val = get_param_str(r, SEPS_PARAM_NAME)
#         if val:
#             keys.add(val.strip())
#     return sorted(keys)


def collect_seps_codes_from_sheets(doc):
    """Collect distinct non-empty SEPS Codes from sheets."""
    sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
    codes = set()
    for s in sheets:
        val = get_param_str(s, SEPS_PARAM_NAME)
        if val:
            codes.add(val.strip())
    return sorted(codes)



def get_scopebox_bbox(doc, seps_code, buffer_ft=1.0):
    """
    Get bounding box (XYZ min, XYZ max) of Scope Box tagged with SEPS Code.
    Returns (min_xyz, max_xyz) or (None, None) if not found.
    """
    scopeboxes = (DB.FilteredElementCollector(doc)
                  .OfClass(DB.VolumeOfInterest)
                  .WhereElementIsNotElementType())
    
    # Find scope box with matching SEPS Code in Name parameter
    for sb in scopeboxes:
        name_param = sb.LookupParameter("Name")
        if name_param and name_param.HasValue:
            name_val = name_param.AsString()
            if name_val and name_val.strip() == seps_code.strip():
                bb = sb.get_BoundingBox(None)
                if not bb:
                    continue
                min_pt = bb.Min
                max_pt = bb.Max

                # Expand by buffer in XY; small buffer in Z
                buf = buffer_ft
                min_pt = DB.XYZ(min_pt.X - buf, min_pt.Y - buf, min_pt.Z - buf)
                max_pt = DB.XYZ(max_pt.X + buf, max_pt.Y + buf, max_pt.Z + buf)

                output.print_md(min_pt, max_pt)

                return min_pt, max_pt


def get_layout_bbox_from_rooms(doc, seps_code, buffer_ft=1.0):
    """
    Get a union bounding box (XYZ min, XYZ max) of all Rooms with the SEPS Code.
    Returns (min_xyz, max_xyz) or (None, None) if no rooms found.
    """
    rooms = (DB.FilteredElementCollector(doc)
             .OfCategory(DB.BuiltInCategory.OST_Rooms)
             .WhereElementIsNotElementType())

    bbox_min = None
    bbox_max = None

    for r in rooms:
        if not element_has_seps(r, seps_code):
            continue
        bb = r.get_BoundingBox(None)
        if not bb:
            continue
        min_pt = bb.Min
        max_pt = bb.Max

        if bbox_min is None:
            bbox_min = DB.XYZ(min_pt.X, min_pt.Y, min_pt.Z)
            bbox_max = DB.XYZ(max_pt.X, max_pt.Y, max_pt.Z)
        else:
            bbox_min = DB.XYZ(min(bbox_min.X, min_pt.X),
                              min(bbox_min.Y, min_pt.Y),
                              min(bbox_min.Z, min_pt.Z))
            bbox_max = DB.XYZ(max(bbox_max.X, max_pt.X),
                              max(bbox_max.Y, max_pt.Y),
                              max(bbox_max.Z, max_pt.Z))

    if bbox_min is None or bbox_max is None:
        return None, None

    # Expand by buffer in XY; small buffer in Z
    buf = buffer_ft
    bbox_min = DB.XYZ(bbox_min.X - buf, bbox_min.Y - buf, bbox_min.Z - buf)
    bbox_max = DB.XYZ(bbox_max.X + buf, bbox_max.Y + buf, bbox_max.Z + buf)

    return bbox_min, bbox_max


def bbox_intersects(elem_bbox, layout_min, layout_max):
    """Simple AABB intersection test in model coordinates."""
    if not elem_bbox or not layout_min or not layout_max:
        return False

    e_min = elem_bbox.Min
    e_max = elem_bbox.Max

    # Separating axis test
    if e_max.X < layout_min.X or e_min.X > layout_max.X:
        return False
    if e_max.Y < layout_min.Y or e_min.Y > layout_max.Y:
        return False
    if e_max.Z < layout_min.Z or e_min.Z > layout_max.Z:
        return False
    return True


def element_belongs_to_layout(elem, seps_code, layout_min, layout_max):
    """
    Decide if an element belongs to the SEPS layout.

    Rules:
    - If element has SEPS Code param == seps_code => keep.
    - Else, if USE_ROOM_BBOX_FILTER and element's bbox intersects the
      layout room bbox => keep.
    - Otherwise => does not belong.
    """
    # 1. Parameter match
    if element_has_seps(elem, seps_code):
        return True

    # 2. Spatial match (inside layout room area)
    if USE_ROOM_BBOX_FILTER and layout_min and layout_max:
        bb = elem.get_BoundingBox(None)
        if bbox_intersects(bb, layout_min, layout_max):
            return True

    return False


#____________________________________________________________________ MAIN

doc = revit.doc
logger = script.get_logger() # Logger for output messages

# 1. Collect SEPS Codes from rooms
# seps_codes = collect_seps_codes_from_rooms(doc)
seps_codes = collect_seps_codes_from_sheets(doc)

if not seps_codes:
    forms.alert(
        "No Rooms found with parameter '{0}'.\n"
        "Please ensure Rooms have this parameter populated per SEPS layout."
        .format(SEPS_PARAM_NAME),
        title="Split by SEPS Code",
        exitscript=True
    )

# 2. Prompt user to select one SEPS Code
selected_code = forms.SelectFromList.show(
    seps_codes,
    title="Select SEPS Code to Export",
    multiselect=False,
    button_name="Export SEPS Layout"
)

if not selected_code:
    script.exit()

seps_code = selected_code

# 3. Compute bounding box of Rooms for that SEPS Code (for spatial filtering)
layout_min, layout_max = (None, None)
if USE_ROOM_BBOX_FILTER:
    layout_min, layout_max = get_scopebox_bbox(
        doc, seps_code, ROOM_BBOX_BUFFER_FT
    )
    # layout_min, layout_max = get_layout_bbox_from_rooms(
    #     doc, seps_code, ROOM_BBOX_BUFFER_FT
    # )

# 4. Prompt for Save As path
default_filename = "{}_{}.rvt".format(doc.Title, seps_code)
save_path = forms.save_file(
    file_ext="rvt",
    default_name=default_filename,
    title="Save SEPS Layout Model As"
)

if not save_path:
    script.exit()

save_options = DB.SaveAsOptions()
save_options.OverwriteExistingFile = True

logger.info("Saving new SEPS layout model to: {}".format(save_path))
doc.SaveAs(save_path, save_options)


# ----------------------------------------------------------------------------
# After SaveAs, we are now operating on the NEW file.
# Delete everything that does not belong to the chosen SEPS layout.
# ----------------------------------------------------------------------------


from System.Collections.Generic import List as DotNetList

tgroup = DB.TransactionGroup(doc, "Prune to SEPS Layout: {}".format(seps_code))
tgroup.Start()

to_delete = DotNetList[DB.ElementId]()

# 5. Prune views (including schedules)
views_collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
for v in views_collector:
    if v.IsTemplate:
        # You may or may not want to keep templates; currently we delete those
        # that are not tagged with the selected SEPS Code.
        if not element_belongs_to_layout(v, seps_code, layout_min, layout_max):
            to_delete.Add(v.Id)
        continue

    if not element_belongs_to_layout(v, seps_code, layout_min, layout_max):
        to_delete.Add(v.Id)

# 6. Prune sheets
sheets_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
for s in sheets_collector:
    if not element_belongs_to_layout(s, seps_code, layout_min, layout_max):
        to_delete.Add(s.Id)

# 7. Prune rooms
rooms_collector = (DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())
for r in rooms_collector:
    if not element_belongs_to_layout(r, seps_code, layout_min, layout_max):
        to_delete.Add(r.Id)

# 8. Prune model instance elements in configured categories
for bic in MODEL_INSTANCE_CATEGORIES:
    elems = (DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType())
    for e in elems:
        if not element_belongs_to_layout(e, seps_code, layout_min, layout_max):
            to_delete.Add(e.Id)

# 9. Execute deletions
if to_delete.Count > 0:
    t = DB.Transaction(doc, "Delete Non-SEPS Elements")
    t.Start()
    try:
        logger.info("Deleting {} elements not in SEPS layout '{}'"
                    .format(to_delete.Count, seps_code))
        doc.Delete(to_delete)
        t.Commit()
    except Exception as ex:
        logger.error("Error deleting elements: {}".format(ex))
        t.RollBack()

tgroup.Assimilate()

forms.alert(
    "SEPS layout '{0}' exported.\nNew model saved to:\n{1}".format(
        seps_code, save_path),
    title="Split by SEPS Code"
)
