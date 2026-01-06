# -*- coding: utf-8 -*-
"""
Export a BIM spatial/relationship graph to JSON for a PixiJS viewer.

Nodes:
- Sheets
- Views (placed on sheets via Viewport)
- Rooms
- Key Equipment (category whitelist)
- Systems (connected to equipment via connectors)

Edges:
- Sheet -> View (placed on)
- View -> Rooms (phase-1: rooms on same GenLevel)
- Room -> Equipment (contained)
- Equipment -> Systems (connected)

Tested pattern: Revit 2024/2025/2026 + pyRevit (IronPython).
"""

import os
import json
import time

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ViewSheet,
    Viewport,
    SpatialElement,
    FamilyInstance,
    XYZ,
    ElementId
)

doc     = __revit__.ActiveUIDocument.Document
uiapp   = __revit__
app     = uiapp.Application

logger = script.get_logger()


# ----------------------------
# Helpers
# ----------------------------
def elementid_int(elementid):
    """
    Convert ElementId to int safely.
    """
    try:
        # return int(elementid.IntegerValue)
        return elementid.ToString()
    except Exception:
        return None

def safe_unique_id(element):
    try:
        return element.UniqueId
    except Exception:
        return None

def node_key(prefix, elementid_int):
    return "{}:{}".format(prefix, elementid_int)

def now_utc_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

def get_location_point(element):
    """
    Return a representative XYZ point for an element.
    Preference order:
      - LocationPoint.Point
      - midpoint of LocationCurve
      - bounding box center (model)
    """
    try:
        location = element.Location
        if location:
            # LocationPoint
            LocaitonPoint = getattr(location, "Point", None)
            if LocaitonPoint:
                return LocaitonPoint
            # LocationCurve
            LocationCurve = getattr(location, "Curve", None)
            if LocationCurve:
                try:
                    return LocationCurve.Evaluate(0.5, True)
                except Exception:
                    pass
    except Exception:
        pass

    # Bounding box fallback
    try:
        boundingbox = element.get_BoundingBox(None)
        if boundingbox and boundingbox.Min and boundingbox.Max:
            return XYZ(
                (boundingbox.Min.X + boundingbox.Max.X) / 2.0,
                (boundingbox.Min.Y + boundingbox.Max.Y) / 2.0,
                (boundingbox.Min.Z + boundingbox.Max.Z) / 2.0
            )
    except Exception:
        pass

    return None

def get_level_id_for_view(view):
    """
    For plan-like views, GenLevel is usually set.
    """
    try:
        gl = getattr(view, "GenLevel", None)
        if gl:
            return gl.Id
    except Exception:
        pass
    return ElementId.InvalidElementId

def get_level_id_for_instance(familyinstance):
    """
    Prefer LevelId when available.
    """
    try:
        level_id = getattr(familyinstance, "LevelId", None)
        if level_id and level_id != ElementId.InvalidElementId:
            return level_id
    except Exception:
        pass
    # fallback: try parameter
    try:
        p = familyinstance.get_Parameter(BuiltInCategory.OST_Levels)  # not correct, but keep safe
    except Exception:
        pass
    return ElementId.InvalidElementId


# ----------------------------
# “Key equipment” category whitelist
# ----------------------------
KEY_EQUIP_CATEGORIES = [
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_DataDevices,
]


# ----------------------------
# Collect nodes + edges
# ----------------------------
nodes = []
edges = []

node_index = {}  # key -> node dict (dedupe)


def add_node(key, nodetype, label, element=None, properties=None, pos=None):
    if key in node_index:
        return node_index[key]

    # Create node
    node = {
        "key": key,
        "type": nodetype,
        "label": label,
        "properties": properties or {},
        "pos": pos or {"x": 0, "y": 0}
    }
    if element is not None:
        node["revit"] = {
            "elementId": element.Id.ToString(),
            "uniqueId": safe_unique_id(element)
        }

    node_index[key] = node
    nodes.append(node)
    return node


def add_edge(element_type, from_key, to_key, properties=None):
    edges.append({
        "type": element_type,
        "from": from_key,
        "to": to_key,
        "properties": properties or {}
    })


# ----------------------------
# Layout: swimlanes by type
# ----------------------------
LANE_X = {
    "sheet": -600,
    "view":  -200,
    "room":   200,
    "equip":  600,
    "system": 1000
}
lane_y = {k: 0 for k in LANE_X.keys()}
LANE_DY = 60

def next_pos(nodetype):
    """
    - nodetype: node type string
    - Returns: {"x": float, "y": float}
    """
    x = LANE_X.get(nodetype, 0)
    y = lane_y.get(nodetype, 0)
    lane_y[nodetype] = y + LANE_DY
    return {"x": x, "y": y}


# ----------------------------
# 1) Sheets + placed Views (Sheet -> View)
# ----------------------------
sheets = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet) if not s.IsPlaceholder]

# Pre-collect all Viewports once
all_viewports = list(FilteredElementCollector(doc).OfClass(Viewport))

views_by_id = {}  # viewId int -> View

for sheet in sheets:
    sheet_key = node_key("sheet", elementid_int(sheet.Id))
    sheet_label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
    add_node(
        sheet_key, "sheet", sheet_label, element=sheet,
        properties={"sheetNumber": sheet.SheetNumber, "sheetName": sheet.Name},
        pos=next_pos("sheet")
    )

    # Viewports on this sheet
    for vp in all_viewports:
        try:
            if vp.SheetId != sheet.Id:
                continue
            viewid = vp.ViewId
            view = doc.GetElement(viewid)
            if view is None:
                continue
            viewid_int = elementid_int(viewid)
            vkey = node_key("view", viewid_int)

            vname = getattr(view, "Name", "View {}".format(viewid_int))
            vtype = str(getattr(view, "ViewType", ""))

            add_node(
                vkey, "view", vname, element=view,
                properties={"viewType": vtype},
                pos=next_pos("view")
            )

            views_by_id[viewid_int] = view
            add_edge("sheet_to_view", sheet_key, vkey, properties={"via": "Viewport"})
        except Exception:
            continue


# ----------------------------
# 2) Rooms
# ----------------------------
rooms = []
rooms_by_level = {}  # levelId int -> [room]
room_col = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()

for r in room_col:
    try:
        # Rooms are SpatialElements
        room = r
        if room is None:
            continue
        rid = elementid_int(room.Id)
        rkey = node_key("room", rid)

        # Label: Number - Name
        number = ""
        name = ""
        try:
            number = room.Number
        except Exception:
            pass
        try:
            name = room.Name
        except Exception:
            pass
        rlabel = "{} - {}".format(number, name).strip(" -")

        # Level
        lvlid = ElementId.InvalidElementId
        try:
            lvlid = room.LevelId
        except Exception:
            pass
        lvl_int = elementid_int(lvlid) if lvlid and lvlid != ElementId.InvalidElementId else None

        add_node(
            rkey, "room", rlabel, element=room,
            properties={"number": number, "name": name, "levelId": lvl_int},
            pos=next_pos("room")
        )

        rooms.append(room)
        if lvl_int is not None:
            rooms_by_level.setdefault(lvl_int, []).append(room)
    except Exception:
        continue


# ----------------------------
# 3) Key Equipment + Room containment (Room -> Equipment)
# ----------------------------
equip_instances = []

for built_in_category in KEY_EQUIP_CATEGORIES:
    try:
        # Collect FamilyInstances in this category
        familyinstance_collector = FilteredElementCollector(doc).OfCategory(built_in_category).WhereElementIsNotElementType()
        for element in familyinstance_collector:
            familyinstance = element  # typically FamilyInstance
            if not isinstance(familyinstance, FamilyInstance):
                continue

            pt = get_location_point(familyinstance)
            if pt is None:
                continue

            eid = elementid_int(familyinstance.Id)
            ekey = node_key("equip", eid)

            fam = ""
            typ = ""
            try:
                sym = familyinstance.Symbol
                if sym:
                    typ = sym.Name
                    fam = sym.Family.Name if sym.Family else ""
            except Exception:
                pass

            jsn = ""
            try:
                p = familyinstance.LookupParameter("JSN")
                if p and p.HasValue:
                    jsn = p.AsString()
                    ekey = node_key("equip", jsn)
            except Exception:
                pass

            label = fam or "Equipment"
            if typ:
                label = "{}: {}".format(label, typ)
            if jsn:
                label = "{} [{}]".format(label, jsn)

            add_node(
                ekey, "equip", label, element=familyinstance,
                properties={"family": fam, "type": typ, "JSN": jsn, "category": str(familyinstance.Category.Name if familyinstance.Category else "")},
                pos=next_pos("equip")
            )

            equip_instances.append((familyinstance, pt))
    except Exception:
        continue

# Room containment edges (Room -> Equip)
# Optimization: only test rooms on same level when possible.
for familyinstance, pt in equip_instances:
    try:
        # Determine candidate rooms by level if possible
        lvl_int = None
        try:
            lid = familyinstance.LevelId
            if lid and lid != ElementId.InvalidElementId:
                lvl_int = elementid_int(lid)
        except Exception:
            pass

        candidate_rooms = rooms_by_level.get(lvl_int, rooms) if lvl_int is not None else rooms

        for room in candidate_rooms:
            try:
                if room and room.IsPointInRoom(pt):
                    rkey = node_key("room", elementid_int(room.Id))
                    jsn = ""
                    try:
                        p = familyinstance.LookupParameter("JSN")
                        if p and p.HasValue:
                            jsn = p.AsString()
                            ekey = node_key("equip", jsn)
                    except Exception:
                        ekey = node_key("equip", elementid_int(familyinstance.Id))
                    add_edge("room_to_equip", rkey, ekey, properties={})
            except Exception:
                continue
    except Exception:
        continue


# ----------------------------
# 4) Systems via connectors (Equip -> System)
# ----------------------------
def iter_connectors(familyinstance):
    """
    Return connectors for a FamilyInstance if available.
    """
    try:
        mep = familyinstance.MEPModel
        if mep is None:
            return []
        cm = getattr(mep, "ConnectorManager", None)
        if cm is None:
            return []
        conns = cm.Connectors
        if conns is None:
            return []
        return list(conns)
    except Exception:
        return []

def system_key_and_label(sys_obj):
    """
    Sys objects can be MEPSystem, ElectricalSystem, or sometimes null.
    Prefer Element-based IDs when available.
    """
    try:
        sid = getattr(sys_obj, "Id", None)
        if sid:
            sid_int = elementid_int(sid)
            if sid_int is not None:
                skey = node_key("system", sid_int)
                sname = getattr(sys_obj, "Name", None) or "System {}".format(sid_int)
                stype = ""
                try:
                    stype = sys_obj.GetType().Name
                except Exception:
                    pass
                return skey, sname, {"systemType": stype}
    except Exception:
        pass

    # Fallback synthetic key (rare)
    skey = "system:{}".format(hash(str(sys_obj)))
    return skey, "System", {"systemType": "Unknown"}

# Track created systems to avoid duplicates
for familyinstance, _pt in equip_instances:
    try:
        element_key = node_key("equip", elementid_int(familyinstance.Id))
        conns = iter_connectors(familyinstance)
        for c in conns:
            system_obj = None

            # Mechanical/Piping
            try:
                system_obj = getattr(c, "MEPSystem", None)
            except Exception:
                system_obj = None

            # Electrical (some builds expose ElectricalSystem)
            if system_obj is None:
                try:
                    system_obj = getattr(c, "ElectricalSystem", None)
                except Exception:
                    system_obj = None

            if system_obj is None:
                continue

            system_key, system_name, system_properties = system_key_and_label(system_obj)

            # Create system node
            add_node(
                system_key, "system", system_name, element=(system_obj if hasattr(system_obj, "UniqueId") else None),
                properties=system_properties,
                pos=next_pos("system")
            )

            add_edge("equip_to_system", element_key, system_key, properties={})
    except Exception:
        continue

"""
# ----------------------------
# 5) View -> Rooms (phase-1: rooms on same GenLevel)
# ----------------------------
for viewid_int, view in views_by_id.items():
    try:
        vkey = node_key("view", viewid_int)
        lvlid = get_level_id_for_view(view)
        if not lvlid or lvlid == ElementId.InvalidElementId:
            continue

        lvl_int = elementid_int(lvlid)
        if lvl_int is None:
            continue

        for room in rooms_by_level.get(lvl_int, []):
            rkey = node_key("room", elementid_int(room.Id))
            add_edge("view_to_room", vkey, rkey, properties={"method": "GenLevel"})
    except Exception:
        continue
"""

# ----------------------------
# Export JSON
# ----------------------------
default_name = "bim_graph.json"
out_path = forms.save_file(file_ext="json", default_name=default_name)

if not out_path:
    script.exit()

payload = {
    "meta": {
        "sourceModelTitle": doc.Title,
        "exportedAtUtc": now_utc_iso(),
        "revitVersion": getattr(app, "VersionNumber", None)
    },
    "nodes": nodes,
    "edges": edges
}

with open(out_path, "w") as f:
    json.dump(payload, f, indent=2)

forms.alert("Exported graph:\n{}".format(out_path), title="BIM Graph Export", warn_icon=False)
