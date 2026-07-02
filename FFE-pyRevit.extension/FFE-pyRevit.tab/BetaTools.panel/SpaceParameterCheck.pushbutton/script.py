# -*- coding: utf-8 -*-
__title__     = "Space Parameter Check"
__version__   = 'Version = v1.0'
__doc__       = """Version = v1.0
Date    = 07.02.2026
___________________________________________________________________
Description:
Checks each Space number/name against the associated Room number/name.
Writes YES or NO to:
- Space vs Room Number Check
- Space vs Room Name Check
___________________________________________________________________
How-to:
-> Click on the button
-> Review the summary
___________________________________________________________________
Last update:
- [07.02.2026] - v1.0 RELEASE
___________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, script, DB, UI
from pyrevit import forms

#____________________________________________________________________ VARIABLES
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type: Document

NUMBER_CHECK_PARAM = "Space vs Room Number Check"
NAME_CHECK_PARAM   = "Space vs Room Name Check"

YES_VALUE = "YES"
NO_VALUE  = "NO"


#____________________________________________________________________ HELPERS
def clean_text(value):
    """Return trimmed text for reliable comparison."""
    if value is None:
        return ""
    try:
        return value.strip()
    except Exception:
        return str(value).strip()


def param_as_text(parameter):
    if not parameter:
        return ""

    try:
        storage_type = parameter.StorageType
        if storage_type == DB.StorageType.String:
            return clean_text(parameter.AsString())
        if storage_type == DB.StorageType.Integer:
            return clean_text(parameter.AsInteger())
        if storage_type == DB.StorageType.Double:
            return clean_text(parameter.AsValueString())
        if storage_type == DB.StorageType.ElementId:
            return clean_text(parameter.AsElementId().ToString())
    except Exception:
        pass

    try:
        return clean_text(parameter.AsValueString())
    except Exception:
        return ""


def get_element_name(element):
    try:
        name = element.Name
        if clean_text(name):
            return clean_text(name)
    except Exception:
        pass

    try:
        return param_as_text(element.LookupParameter("Name"))
    except Exception:
        return ""


def get_element_number(element):
    try:
        number = element.Number
        if clean_text(number):
            return clean_text(number)
    except Exception:
        pass

    try:
        return param_as_text(element.LookupParameter("Number"))
    except Exception:
        return ""


def values_match(left, right):
    return clean_text(left) == clean_text(right)


def get_location_point(element):
    """Return a representative XYZ point for Spaces and Rooms."""
    try:
        location = element.Location
        if location:
            point = getattr(location, "Point", None)
            if point:
                return point

            curve = getattr(location, "Curve", None)
            if curve:
                try:
                    return curve.Evaluate(0.5, True)
                except Exception:
                    pass
    except Exception:
        pass

    try:
        bounding_box = element.get_BoundingBox(None)
        if bounding_box and bounding_box.Min and bounding_box.Max:
            return DB.XYZ(
                (bounding_box.Min.X + bounding_box.Max.X) / 2.0,
                (bounding_box.Min.Y + bounding_box.Max.Y) / 2.0,
                (bounding_box.Min.Z + bounding_box.Max.Z) / 2.0
            )
    except Exception:
        pass

    return None


def get_space_phase(space, owner_doc):
    try:
        phase_id = space.CreatedPhaseId
        if phase_id and phase_id != DB.ElementId.InvalidElementId:
            phase = owner_doc.GetElement(phase_id)
            if phase:
                return phase
    except Exception:
        pass

    try:
        phase_param = owner_doc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
        if phase_param:
            phase = owner_doc.GetElement(phase_param.AsElementId())
            if phase:
                return phase
    except Exception:
        pass

    return None


def collect_spaces(owner_doc):
    try:
        return list(
            DB.FilteredElementCollector(owner_doc)
            .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def collect_rooms(owner_doc):
    try:
        return list(
            DB.FilteredElementCollector(owner_doc)
            .OfCategory(DB.BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def collect_link_contexts(host_doc):
    contexts = []
    try:
        link_instances = (
            DB.FilteredElementCollector(host_doc)
            .OfClass(DB.RevitLinkInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        link_instances = []

    for link_instance in link_instances:
        try:
            link_doc = link_instance.GetLinkDocument()
            if not link_doc:
                continue
        except Exception:
            continue

        try:
            transform = link_instance.GetTotalTransform()
        except Exception:
            try:
                transform = link_instance.GetTransform()
            except Exception:
                transform = None

        if not transform:
            continue

        try:
            inverse_transform = transform.Inverse
        except Exception:
            continue

        contexts.append({
            "doc": link_doc,
            "title": link_doc.Title,
            "inverse_transform": inverse_transform,
            "rooms": collect_rooms(link_doc)
        })

    return contexts


def find_room_at_point(owner_doc, point, rooms, phase=None):
    if point is None:
        return None

    if phase:
        try:
            room = owner_doc.GetRoomAtPoint(point, phase)
            if room:
                return room
        except Exception:
            pass

    try:
        room = owner_doc.GetRoomAtPoint(point)
        if room:
            return room
    except Exception:
        pass

    for room in rooms:
        try:
            if room and room.IsPointInRoom(point):
                return room
        except Exception:
            continue

    return None


def get_direct_space_room(space):
    try:
        room = space.Room
        if room:
            return room
    except Exception:
        pass

    return None


def find_associated_room(space, host_doc, host_rooms, link_contexts):
    direct_room = get_direct_space_room(space)
    if direct_room:
        return direct_room, "Direct"

    point = get_location_point(space)
    if point is None:
        return None, "No Space location"

    phase = get_space_phase(space, host_doc)
    host_room = find_room_at_point(host_doc, point, host_rooms, phase)
    if host_room:
        return host_room, host_doc.Title

    for context in link_contexts:
        try:
            link_point = context["inverse_transform"].OfPoint(point)
        except Exception:
            continue

        link_room = find_room_at_point(context["doc"], link_point, context["rooms"])
        if link_room:
            return link_room, context["title"]

    return None, "No associated Room found"


def set_check_parameter(element, param_name, passed):
    parameter = element.LookupParameter(param_name)
    if not parameter:
        return False, "Missing parameter: {}".format(param_name)
    if parameter.IsReadOnly:
        return False, "Read-only parameter: {}".format(param_name)

    text_value = YES_VALUE if passed else NO_VALUE

    try:
        if parameter.StorageType == DB.StorageType.Integer:
            parameter.Set(1 if passed else 0)
        elif parameter.StorageType == DB.StorageType.String:
            parameter.Set(text_value)
        else:
            parameter.Set(text_value)
        return True, None
    except Exception as ex:
        return False, "{}: {}".format(param_name, ex)


def validate_required_parameters(spaces):
    if not spaces:
        forms.alert("No Spaces were found in the current model.", title=__title__, warn_icon=True)
        return False

    missing = []
    sample_space = spaces[0]
    for param_name in [NUMBER_CHECK_PARAM, NAME_CHECK_PARAM]:
        if not sample_space.LookupParameter(param_name):
            missing.append(param_name)

    if missing:
        forms.alert(
            "Required Space project parameters are missing:\n\n{}"
            "\n\nAdd/bind these parameters to MEP Spaces and run the tool again.".format(
                "\n".join(["- {}".format(name) for name in missing])
            ),
            title=__title__,
            warn_icon=True
        )
        return False

    return True


def build_detail_line(space, room, room_source, number_match, name_match):
    space_number = get_element_number(space)
    space_name = get_element_name(space)

    if room:
        room_number = get_element_number(room)
        room_name = get_element_name(room)
    else:
        room_number = "(none)"
        room_name = "(none)"

    return "{} - {} | Room: {} - {} | Source: {} | Number: {} | Name: {}".format(
        space_number or "(blank)",
        space_name or "(blank)",
        room_number or "(blank)",
        room_name or "(blank)",
        room_source,
        YES_VALUE if number_match else NO_VALUE,
        YES_VALUE if name_match else NO_VALUE
    )


#____________________________________________________________________ MAIN
def main():
    host_doc = revit.doc
    output = script.get_output()

    spaces = collect_spaces(host_doc)
    if not validate_required_parameters(spaces):
        return

    host_rooms = collect_rooms(host_doc)
    link_contexts = collect_link_contexts(host_doc)

    number_yes = 0
    number_no = 0
    name_yes = 0
    name_no = 0
    no_room_count = 0
    write_failures = []
    review_lines = []

    with revit.Transaction("FFE Space vs Room Check"):
        for space in spaces:
            room, room_source = find_associated_room(space, host_doc, host_rooms, link_contexts)

            if room:
                number_match = values_match(get_element_number(space), get_element_number(room))
                name_match = values_match(get_element_name(space), get_element_name(room))
            else:
                number_match = False
                name_match = False
                no_room_count += 1

            ok_num, err_num = set_check_parameter(space, NUMBER_CHECK_PARAM, number_match)
            ok_name, err_name = set_check_parameter(space, NAME_CHECK_PARAM, name_match)

            if not ok_num and len(write_failures) < 12:
                write_failures.append("Space {}: {}".format(space.Id.ToString(), err_num))
            if not ok_name and len(write_failures) < 12:
                write_failures.append("Space {}: {}".format(space.Id.ToString(), err_name))

            if number_match:
                number_yes += 1
            else:
                number_no += 1

            if name_match:
                name_yes += 1
            else:
                name_no += 1

            if (not room) or (not number_match) or (not name_match):
                review_lines.append(build_detail_line(space, room, room_source, number_match, name_match))

    if review_lines:
        output.print_md("### Space vs Room Check Review")
        for line in review_lines:
            output.print_md("- {}".format(line))

    msg = []
    msg.append("Spaces checked: {}".format(len(spaces)))
    msg.append("Host Rooms available: {}".format(len(host_rooms)))
    msg.append("Loaded linked models checked: {}".format(len(link_contexts)))
    msg.append("")
    msg.append("{}: {} YES / {} NO".format(NUMBER_CHECK_PARAM, number_yes, number_no))
    msg.append("{}: {} YES / {} NO".format(NAME_CHECK_PARAM, name_yes, name_no))
    msg.append("")
    msg.append("Spaces with no associated Room: {}".format(no_room_count))

    if review_lines:
        msg.append("Review lines written to pyRevit output: {}".format(len(review_lines)))

    if write_failures:
        msg.append("")
        msg.append("Write issues:")
        msg.extend(["- {}".format(item) for item in write_failures])

    forms.alert("\n".join(msg), title=__title__, warn_icon=False)


if __name__ == "__main__":
    main()
