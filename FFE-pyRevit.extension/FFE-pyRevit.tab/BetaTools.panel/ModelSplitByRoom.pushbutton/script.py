# -*- coding: utf-8 -*-
__title__     = "Design Guide \nModel Splitter"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 12.15.2025
_________________________________________________________________
Description:
Split current model into a per-SEPS-Code layout model.

_________________________________________________________________
How-to:
1. Open Revit model detached and discard worksets.
2. Run this script.
3. Assumes a project/shared parameter "SEPS Code" exists on:
   - Rooms
   - Views (including schedules)
   - Sheets
   - Scope Boxes
4. Reads all distinct SEPS Codes from Sheets.
5. Lets user pick one SEPS Code.
6. Prompts user for a target .rvt file name.
7. Saves current document as that file (Save As).
8. In the newly saved model, deletes all elements that DO NOT belong
   to that SEPS layout.

Belonging to the layout is defined as:
- Sheet has "SEPS Code" == selected code.
- View/Schedule has "SEPS Code" == selected code.
- Element's bounding box intersects the bounding box of the Scope
  Box whose "SEPS Code" == selected code.
_________________________________________________________________
Last update:
- [12.01.2025] - v0.1 BETA RELEASE
- [12.15.2025] - v1.0 RELEASE
_________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)
from logging import Filter

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Transaction, TransactionGroup, FilteredElementCollector


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, forms, script

from pyrevit.script import output


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
    DB.BuiltInCategory.OST_VolumeOfInterest,
    DB.BuiltInCategory.OST_Grids,
    DB.BuiltInCategory.OST_RoomSeparationLines,
]

# Use room bounding boxes to keep geometry that sits inside the layout area
# even if the elements do not have a SEPS Code parameter.
USE_SCOPEBOX_FILTER = True

# Buffer around room bounding box, in feet
SCOPEBOX_BBOX_BUFFER_FT = 1.0

#____________________________________________________________________ VARIABLES
output_window = output.get_output()


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


def collect_seps_codes_from_sheets(doc):
    """Collect distinct non-empty SEPS Codes from sheets."""
    sheets = FilteredElementCollector(doc).OfClass(ViewSheet)
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
    scopeboxes = (FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType())
    
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
                # buf = buffer_ft
                # min_pt = DB.XYZ(min_pt.X - buf, min_pt.Y - buf, min_pt.Z - buf)
                # max_pt = DB.XYZ(max_pt.X + buf, max_pt.Y + buf, max_pt.Z + buf)

                # output.print_md([min_pt, max_pt])

                return min_pt, max_pt



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
    - Else, if USE_SCOPEBOX_FILTER and element's bbox intersects the
      layout room bbox => keep.
    - Otherwise => does not belong.
    """
    # 1. Parameter match
    if element_has_seps(elem, seps_code):
        return True

    # 2. Spatial match (inside layout room area)
    if USE_SCOPEBOX_FILTER and layout_min and layout_max:
        bb = elem.get_BoundingBox(None)
        if bbox_intersects(bb, layout_min, layout_max):
            return True

    return False


# def centralmodel(doc):
#     """ Get the central model path as string """
#     from Autodesk.Revit.DB import ModelPathUtils
#     central_model = doc.GetWorksharingCentralModelPath()
#     path_string = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_model)
#     return path_string


#____________________________________________________________________ MAIN

doc = revit.doc
logger = script.get_logger() # Logger for output messages


# 1. Collect SEPS Codes from Sheets
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


# 3. Compute bounding box of Scope Box for that SEPS Code (for spatial filtering)
layout_min, layout_max = (None, None)
if USE_SCOPEBOX_FILTER:
    layout_min, layout_max = get_scopebox_bbox(doc, seps_code, SCOPEBOX_BBOX_BUFFER_FT)


# 4. Prompt for Save As path
# Checl if model is workshared
if doc.IsWorkshared:
    forms.alert(
        "This model is workshared.\n"
        "Please detach the model and discard worksets before running this script.",
        title="Split by SEPS Code",
        exitscript=True
    )

# Set default filename based on selected SEPS Code
default_filename = "{}.rvt".format(seps_code)
save_path = forms.save_file(
    file_ext="rvt",
    default_name=default_filename,
    title="Save SEPS Layout Model As"
)

# Exit if user cancels the save dialog
if not save_path:
    script.exit()

# Set save options
save_options = DB.SaveAsOptions()
save_options.OverwriteExistingFile = True

# Save the document as a new file
logger.info("Saving new SEPS layout model to: {}".format(save_path))
doc.SaveAs(save_path, save_options)

output_window.print_md("Model saved to: **{}**".format(save_path))

# ----------------------------------------------------------------------------
# After SaveAs, we are now operating on the NEW file.
# Delete everything that does not belong to the chosen SEPS layout.
# ----------------------------------------------------------------------------

doc = revit.doc


from System.Collections.Generic import List as DotNetList

tgroup = TransactionGroup(doc, "Prune to SEPS Layout: {}".format(seps_code))
tgroup.Start()

to_delete       = DotNetList[DB.ElementId]()


# 5. Prune sheets
sheet_collector = FilteredElementCollector(doc).OfClass(ViewSheet)
sheets = []
to_delete_sheets = []
for s in sheet_collector:
    s_param = s.LookupParameter("SEPS Code")
    if s_param and s_param.HasValue:                # Check SEPS Code parameter has value
        s_val = s_param.AsString()
        if s_val.strip() != seps_code.strip():      # If SEPS Code does not match
            to_delete.Add(s.Id)
            to_delete_sheets.append(s.Id)
        else:
            sheets.append(s.Id)


# 6. Prune views (including schedules)
view_collector = FilteredElementCollector(doc).OfClass(View)
filteredViewTypes = [
	"CostReport","Internal","LoadsReport",
    "PresureLossReport","ProjectBrowser","Report",
	"SystemBrowser","SystemsAnalysisReport","Undefined", "DrawingSheet"
	]
to_delete_views = []
for v in view_collector:
    if not v.IsTemplate:                                        # Exclude template views
        if v.ViewType.ToString() not in filteredViewTypes:      # Exclude certain view types
            if v.Name != "Starting View":                       # Exclude "Starting View"
                v_param = v.LookupParameter("SEPS Code")
                if v_param and v_param.HasValue:                # Check SEPS Code parameter has value
                    v_val = v_param.AsString()
                    if v_val.strip() != seps_code.strip():      # If SEPS Code does not match
                        to_delete.Add(v.Id)
                        to_delete_views.append(v.Id)
                else:                                           # If SEPS Code parameter is empty
                    to_delete.Add(v.Id)
                    to_delete_views.append(v.Id)


# viewport_collector = FilteredElementCollector(doc).OfClass(Viewport)
# to_delete_views_alt = []
# views = []
# for vp in viewport_collector:
#     s_id = vp.SheetId
#     if s_id in sheets:
#         views.append(vp.ViewId)
#     else:
#         to_delete_views_alt.append(vp.ViewId)
#         to_delete.Add(vp.ViewId)


# schedule_collector = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
# to_delete_schedules = []
# schedules = []
# for sch in schedule_collector:
#     sch_owner = sch.OwnerViewId
#     if sch_owner in sheets:
#         schedules.append(sch.ScheduleId)
#     else:
#         to_delete_schedules.append(sch.ScheduleId)
#         to_delete.Add(sch.ScheduleId)



# 7. Prune rooms
rooms_collector = (DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())
for r in rooms_collector:
    if not element_belongs_to_layout(r, seps_code, layout_min, layout_max):
        to_delete.Add(r.Id)


# 8. Prune model instance elements in configured categories
to_delete_elements = []
for bic in MODEL_INSTANCE_CATEGORIES:
    elems = (DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType())
    for e in elems:
        if not element_belongs_to_layout(e, seps_code, layout_min, layout_max):
            to_delete.Add(e.Id)
            to_delete_elements.append(e.Id)


#____________________________________________________ DEBUGGING OUTPUT
output_window.print_md("### DEBUG INFO:")
output_window.print_md("- Selected SEPS Code: **{}**".format(seps_code))
output_window.print_md("- Layout BBox Min: **{}**".format(layout_min))
output_window.print_md("- Layout BBox Max: **{}**".format(layout_max))
output_window.print_md("- Elements to Delete: **{}**".format(to_delete.Count))
output_window.print_md("- Elements to Delete (Sheets): **{}**".format(len(to_delete_sheets)))
output_window.print_md("- Elements to Delete (Views): **{}**".format(len(to_delete_views)))
output_window.print_md("- Elements to Delete (Elements): **{}**".format(len(to_delete_elements)))
output_window.print_md("---")


# 9. Execute deletions
if to_delete.Count > 0:
    t = Transaction(doc, "Delete Non-SEPS Elements")
    t.Start()
    try:
        logger.info("Deleting {} elements not in SEPS layout '{}'".format(to_delete.Count, seps_code))
        doc.Delete(to_delete)
        t.Commit()
    except Exception as ex:
        logger.error("Error deleting elements: {}".format(ex))
        t.RollBack()

# 10. Complete transaction group
# tgroup.Assimilate()
tgroup.Commit()




#____________________________________________________ COMPLETED
forms.alert(
    "SEPS layout '{0}' exported.\nNew model saved to:\n{1}".format(
        seps_code, save_path),
    title="Split by SEPS Code"
)
