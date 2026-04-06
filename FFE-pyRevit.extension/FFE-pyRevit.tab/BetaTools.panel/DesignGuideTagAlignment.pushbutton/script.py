# -*- coding: utf-8 -*-
__title__ = "Design Guide Tag Alignment"
__version__ = "Version = v0.3"
__doc__ = """Version = v0.3
Date    = 04.06.2026
______________________________________________________________
Description:
-> Arrange selected tags by left/right tag family in the active view.
-> Places tags 2'-0" from the wall and sets a 1'-0" leader shoulder.
______________________________________________________________
How-to:
-> Select the tags you want to arrange.
-> Click the button.
______________________________________________________________
Last update:
- [04.06.2026] - v0.3 Added family-name-driven left/right placement
______________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS
import clr

clr.AddReference("System")

from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import IndependentTag
from Autodesk.Revit.DB import LeaderEndCondition
from Autodesk.Revit.DB import XYZ

from pyrevit import revit
from pyrevit import forms
from pyrevit import script


# ____________________________________________________________________ VARIABLES
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView
selection = uidoc.Selection
output = script.get_output()

ONE_FOOT = 1.0
WALL_OFFSET = 2.0
TOP_BOTTOM_PADDING = 0.25
MIN_VERTICAL_GAP = 0.15


def format_float(value):
    return "{0:.4f}".format(value)


def format_xyz(point):
    if not point:
        return "(None)"
    return "({0}, {1}, {2})".format(
        format_float(point.X),
        format_float(point.Y),
        format_float(point.Z),
    )


def format_bbox(bbox):
    if not bbox:
        return "bbox=None"
    return "min={0} max={1}".format(format_xyz(bbox.Min), format_xyz(bbox.Max))


def debug_log(message):
    print(message)


# ____________________________________________________________________ HELPERS
def get_selected_tags():
    """Return selected IndependentTag instances."""
    tags = []
    for element_id in selection.GetElementIds():
        element = doc.GetElement(element_id)
        if isinstance(element, IndependentTag):
            tags.append(element)
    return tags


def get_view_bounds(active_view):
    """Return the active view crop bounds in model coordinates."""
    crop_box = active_view.CropBox
    if not crop_box:
        return None

    transform = crop_box.Transform
    min_pt = transform.OfPoint(crop_box.Min)
    max_pt = transform.OfPoint(crop_box.Max)

    min_x = min(min_pt.X, max_pt.X)
    max_x = max(min_pt.X, max_pt.X)
    min_y = min(min_pt.Y, max_pt.Y)
    max_y = max(min_pt.Y, max_pt.Y)
    debug_log(
        "VIEW_BOUNDS | crop_min={0} crop_max={1} | resolved min_x={2} max_x={3} min_y={4} max_y={5}".format(
            format_xyz(min_pt),
            format_xyz(max_pt),
            format_float(min_x),
            format_float(max_x),
            format_float(min_y),
            format_float(max_y),
        )
    )
    return min_x, max_x, min_y, max_y


def get_tag_bbox(tag):
    """Return the tag bounding box in the active view."""
    bbox = tag.get_BoundingBox(view)
    if bbox:
        return bbox
    return None


def get_bbox_height(bbox):
    return bbox.Max.Y - bbox.Min.Y


def get_tag_family_label(tag):
    """Return a searchable label that prefers the tag family name."""
    tag_type = doc.GetElement(tag.GetTypeId())
    if not tag_type:
        return ""

    try:
        family_name = tag_type.FamilyName or ""
    except Exception:
        family_name = ""

    if not family_name:
        try:
            family_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if family_param:
                family_name = family_param.AsString() or ""
        except Exception:
            family_name = ""

    try:
        type_name = tag_type.Name or ""
    except Exception:
        type_name = ""

    label = "{} {}".format(family_name, type_name).strip()
    if label:
        return label

    try:
        family = tag_type.Family
        if family and family.Name:
            return family.Name
    except Exception:
        pass

    try:
        symbol_name_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if symbol_name_param:
            return symbol_name_param.AsString() or ""
    except Exception:
        pass

    return ""


def get_tag_layout_side(tag):
    """Map tag family naming to the side of the view and leader shoulder direction."""
    family_label = get_tag_family_label(tag)
    family_label_lower = family_label.lower()

    if "right" in family_label_lower:
        return "Left", "Right", family_label
    if "left" in family_label_lower:
        return "Right", "Left", family_label
    return None, None, family_label


def get_single_tag_reference(tag):
    """Return the first tagged reference for a tag."""
    try:
        references = list(tag.GetTaggedReferences())
        if references:
            return references[0]
    except Exception:
        pass

    try:
        return tag.GetTaggedReference()
    except Exception:
        pass

    return None


def set_tag_elbow(tag, elbow_point):
    """Set the elbow point for single-reference leaders when possible."""
    reference = get_single_tag_reference(tag)
    if reference is None:
        debug_log("TAG {0} | no reference found for elbow.".format(tag.Id.IntegerValue))
        return

    try:
        tag.LeaderEndCondition = LeaderEndCondition.Free
    except Exception:
        pass

    try:
        tag.SetLeaderElbow(reference, elbow_point)
        debug_log("TAG {0} | SetLeaderElbow -> {1}".format(tag.Id.IntegerValue, format_xyz(elbow_point)))
    except Exception:
        try:
            tag.LeaderElbow = elbow_point
            debug_log("TAG {0} | LeaderElbow property -> {1}".format(tag.Id.IntegerValue, format_xyz(elbow_point)))
        except Exception:
            debug_log("TAG {0} | failed to set elbow.".format(tag.Id.IntegerValue))
            pass


def align_tag_to_side(tag, view_side, target_edge_x, target_y):
    """Move a tag to the requested side and return its final bounding box."""
    current_head = tag.TagHeadPosition
    debug_log(
        "TAG {0} | align start | side={1} | head={2} | target_edge_x={3} | target_y={4}".format(
            tag.Id.IntegerValue,
            view_side,
            format_xyz(current_head),
            format_float(target_edge_x),
            format_float(target_y),
        )
    )
    tag.TagHeadPosition = XYZ(current_head.X, target_y, current_head.Z)
    doc.Regenerate()

    bbox = get_tag_bbox(tag)
    if not bbox:
        debug_log("TAG {0} | no bbox after Y move.".format(tag.Id.IntegerValue))
        return None

    debug_log("TAG {0} | bbox after Y move | {1}".format(tag.Id.IntegerValue, format_bbox(bbox)))

    if view_side == "Left":
        shift_x = target_edge_x - bbox.Min.X
    else:
        shift_x = target_edge_x - bbox.Max.X

    updated_head = tag.TagHeadPosition
    tag.TagHeadPosition = XYZ(updated_head.X + shift_x, target_y, updated_head.Z)
    doc.Regenerate()
    final_bbox = get_tag_bbox(tag)
    debug_log(
        "TAG {0} | align end | shift_x={1} | final_head={2} | final_bbox={3}".format(
            tag.Id.IntegerValue,
            format_float(shift_x),
            format_xyz(tag.TagHeadPosition),
            format_bbox(final_bbox),
        )
    )
    return final_bbox


def build_vertical_stack(tag_items, min_y, max_y):
    """Return center Y positions that respect actual tag heights."""
    if not tag_items:
        return []

    y_positions = []
    current_top = max_y - TOP_BOTTOM_PADDING

    for index, item in enumerate(tag_items):
        bbox = item["bbox"]
        height = get_bbox_height(bbox)
        center_y = current_top - (height / 2.0)

        if center_y - (height / 2.0) < min_y + TOP_BOTTOM_PADDING:
            center_y = min_y + TOP_BOTTOM_PADDING + (height / 2.0)

        y_positions.append(center_y)
        debug_log(
            "STACK | side={0} | index={1} | tag_id={2} | source_y={3} | height={4} | center_y={5} | next_top={6}".format(
                item["view_side"],
                index,
                item["tag"].Id.IntegerValue,
                format_float(item["y"]),
                format_float(height),
                format_float(center_y),
                format_float(center_y - (height / 2.0) - MIN_VERTICAL_GAP),
            )
        )
        current_top = center_y - (height / 2.0) - MIN_VERTICAL_GAP

    return y_positions


def process_side(tag_items, view_side, view_bounds):
    """Arrange a group of tags on one side of the view."""
    if not tag_items:
        return

    min_x, max_x, min_y, max_y = view_bounds
    tag_items.sort(key=lambda item: item["y"], reverse=True)
    y_positions = build_vertical_stack(tag_items, min_y, max_y)
    target_edge_x = min_x + WALL_OFFSET if view_side == "Left" else max_x - WALL_OFFSET
    debug_log(
        "PROCESS_SIDE | side={0} | tag_count={1} | target_edge_x={2} | min_y={3} | max_y={4}".format(
            view_side,
            len(tag_items),
            format_float(target_edge_x),
            format_float(min_y),
            format_float(max_y),
        )
    )

    for y_position, item in zip(y_positions, tag_items):
        tag = item["tag"]
        debug_log(
            "TAG {0} | process | family_label='{1}' | leader_side={2} | initial_head={3} | initial_bbox={4}".format(
                tag.Id.IntegerValue,
                item["family_label"],
                item["leader_side"],
                format_xyz(tag.TagHeadPosition),
                format_bbox(item["bbox"]),
            )
        )
        final_bbox = align_tag_to_side(tag, view_side, target_edge_x, y_position)
        if not final_bbox:
            continue

        if tag.HasLeader:
            if item["leader_side"] == "Right":
                elbow_x = final_bbox.Max.X + ONE_FOOT
            else:
                elbow_x = final_bbox.Min.X - ONE_FOOT

            elbow_point = XYZ(elbow_x, y_position, tag.TagHeadPosition.Z)
            debug_log(
                "TAG {0} | elbow calc | leader_side={1} | elbow_x={2} | elbow_point={3}".format(
                    tag.Id.IntegerValue,
                    item["leader_side"],
                    format_float(elbow_x),
                    format_xyz(elbow_point),
                )
            )
            set_tag_elbow(tag, elbow_point)


def main():
    output.set_title(__title__)
    debug_log("=== Design Guide Tag Alignment Debug ===")
    debug_log("ACTIVE_VIEW | id={0} | name='{1}'".format(view.Id.IntegerValue, view.Name))
    tags = get_selected_tags()
    if not tags:
        forms.alert("Select one or more tags before running this tool.", exitscript=True)

    debug_log("SELECTION | tag_count={0}".format(len(tags)))

    view_bounds = get_view_bounds(view)
    if not view_bounds:
        forms.alert("The active view does not expose a usable crop box.", exitscript=True)

    left_side_tags = []
    right_side_tags = []
    skipped_tags = []

    for tag in tags:
        bbox = get_tag_bbox(tag)
        if not bbox:
            continue

        view_side, leader_side, type_name = get_tag_layout_side(tag)
        debug_log(
            "TAG {0} | detected | family_label='{1}' | mapped_view_side={2} | mapped_leader_side={3} | head={4} | bbox={5}".format(
                tag.Id.IntegerValue,
                type_name,
                view_side,
                leader_side,
                format_xyz(tag.TagHeadPosition),
                format_bbox(bbox),
            )
        )
        if not view_side:
            skipped_tags.append(type_name or str(tag.Id.IntegerValue))
            continue

        item = {
            "tag": tag,
            "bbox": bbox,
            "y": tag.TagHeadPosition.Y,
            "leader_side": leader_side,
            "view_side": view_side,
            "family_label": type_name,
        }

        if view_side == "Left":
            left_side_tags.append(item)
        else:
            right_side_tags.append(item)

    if not left_side_tags and not right_side_tags:
        forms.alert("No selected tags matched a type name containing 'Left' or 'Right'.", exitscript=True)

    with revit.Transaction("Design Guide Tag Alignment"):
        process_side(left_side_tags, "Left", view_bounds)
        process_side(right_side_tags, "Right", view_bounds)

    debug_log("SUMMARY | left_tags={0} | right_tags={1} | skipped={2}".format(
        len(left_side_tags),
        len(right_side_tags),
        len(skipped_tags),
    ))

    if skipped_tags:
        forms.alert(
            "Some selected tags were skipped because their type names did not include 'Left' or 'Right'.",
            title="Design Guide Tag Alignment",
            warn_icon=False,
        )


if __name__ == "__main__":
    main()
