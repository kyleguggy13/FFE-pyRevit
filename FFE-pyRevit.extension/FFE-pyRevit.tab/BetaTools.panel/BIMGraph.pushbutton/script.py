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
def eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        return None

def safe_unique_id(el):
    try:
        return el.UniqueId
    except Exception:
        return None

def node_key(prefix, elid_int):
    return "{}:{}".format(prefix, elid_int)

def now_utc_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

def get_location_point(el):
    """
    Return a representative XYZ point for an element.
    Preference order:
      - LocationPoint.Point
      - midpoint of LocationCurve
      - bounding box center (model)
    """
    try:
        loc = el.Location
        if loc:
            # LocationPoint
            p = getattr(loc, "Point", None)
            if p:
                return p
            # LocationCurve
            crv = getattr(loc, "Curve", None)
            if crv:
                try:
                    return crv.Evaluate(0.5, True)
                except Exception:
                    pass
    except Exception:
        pass

    # Bounding box fallback
    try:
        bb = el.get_BoundingBox(None)
        if bb and bb.Min and bb.Max:
            return XYZ(
                (bb.Min.X + bb.Max.X) / 2.0,
                (bb.Min.Y + bb.Max.Y) / 2.0,
                (bb.Min.Z + bb.Max.Z) / 2.0
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

def get_level_id_for_instance(fi):
    """
    Prefer LevelId when available.
    """
    try:
        lid = getattr(fi, "LevelId", None)
        if lid and lid != ElementId.InvalidElementId:
            return lid
    except Exception:
        pass
    # fallback: try parameter
    try:
        p = fi.get_Parameter(BuiltInCategory.OST_Levels)  # not correct, but keep safe
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


def add_node(key, ntype, label, el=None, props=None, pos=None):
    if key in node_index:
        return node_index[key]

    nd = {
        "key": key,
        "type": ntype,
        "label": label,
        "props": props or {},
        "pos": pos or {"x": 0, "y": 0}
    }
    if el is not None:
        nd["revit"] = {
            "elementId": eid_int(el.Id),
            "uniqueId": safe_unique_id(el)
        }

    node_index[key] = nd
    nodes.append(nd)
    return nd


def add_edge(etype, from_key, to_key, props=None):
    edges.append({
        "type": etype,
        "from": from_key,
        "to": to_key,
        "props": props or {}
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

def next_pos(ntype):
    x = LANE_X.get(ntype, 0)
    y = lane_y.get(ntype, 0)
    lane_y[ntype] = y + LANE_DY
    return {"x": x, "y": y}


# ----------------------------
# 1) Sheets + placed Views (Sheet -> View)
# ----------------------------
sheets = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet) if not s.IsPlaceholder]

# Pre-collect all Viewports once
all_viewports = list(FilteredElementCollector(doc).OfClass(Viewport))

views_by_id = {}  # viewId int -> View

for sheet in sheets:
    skey = node_key("sheet", eid_int(sheet.Id))
    sheet_label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
    add_node(
        skey, "sheet", sheet_label, el=sheet,
        props={"sheetNumber": sheet.SheetNumber, "sheetName": sheet.Name},
        pos=next_pos("sheet")
    )

    # Viewports on this sheet
    for vp in all_viewports:
        try:
            if vp.SheetId != sheet.Id:
                continue
            vid = vp.ViewId
            view = doc.GetElement(vid)
            if view is None:
                continue
            vid_int = eid_int(vid)
            vkey = node_key("view", vid_int)

            vname = getattr(view, "Name", "View {}".format(vid_int))
            vtype = str(getattr(view, "ViewType", ""))

            add_node(
                vkey, "view", vname, el=view,
                props={"viewType": vtype},
                pos=next_pos("view")
            )

            views_by_id[vid_int] = view
            add_edge("sheet_to_view", skey, vkey, props={"via": "Viewport"})
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
        rid = eid_int(room.Id)
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
        lvl_int = eid_int(lvlid) if lvlid and lvlid != ElementId.InvalidElementId else None

        add_node(
            rkey, "room", rlabel, el=room,
            props={"number": number, "name": name, "levelId": lvl_int},
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

for bic in KEY_EQUIP_CATEGORIES:
    try:
        col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        for el in col:
            fi = el  # typically FamilyInstance
            if not isinstance(fi, FamilyInstance):
                continue

            pt = get_location_point(fi)
            if pt is None:
                continue

            eid = eid_int(fi.Id)
            ekey = node_key("equip", eid)

            fam = ""
            typ = ""
            try:
                sym = fi.Symbol
                if sym:
                    typ = sym.Name
                    fam = sym.Family.Name if sym.Family else ""
            except Exception:
                pass

            jsn = ""
            try:
                p = fi.LookupParameter("JSN")
                if p and p.HasValue:
                    jsn = p.AsString()
            except Exception:
                pass

            label = fam or "Equipment"
            if typ:
                label = "{}: {}".format(label, typ)
            if jsn:
                label = "{} [{}]".format(label, jsn)

            add_node(
                ekey, "equip", label, el=fi,
                props={"family": fam, "type": typ, "JSN": jsn, "category": str(fi.Category.Name if fi.Category else "")},
                pos=next_pos("equip")
            )

            equip_instances.append((fi, pt))
    except Exception:
        continue

# Room containment edges (Room -> Equip)
# Optimization: only test rooms on same level when possible.
for fi, pt in equip_instances:
    try:
        # Determine candidate rooms by level if possible
        lvl_int = None
        try:
            lid = fi.LevelId
            if lid and lid != ElementId.InvalidElementId:
                lvl_int = eid_int(lid)
        except Exception:
            pass

        candidate_rooms = rooms_by_level.get(lvl_int, rooms) if lvl_int is not None else rooms

        for room in candidate_rooms:
            try:
                if room and room.IsPointInRoom(pt):
                    rkey = node_key("room", eid_int(room.Id))
                    ekey = node_key("equip", eid_int(fi.Id))
                    add_edge("room_to_equip", rkey, ekey, props={})
            except Exception:
                continue
    except Exception:
        continue


# ----------------------------
# 4) Systems via connectors (Equip -> System)
# ----------------------------
def iter_connectors(fi):
    """
    Return connectors for a FamilyInstance if available.
    """
    try:
        mep = fi.MEPModel
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
            sid_int = eid_int(sid)
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
for fi, _pt in equip_instances:
    try:
        ekey = node_key("equip", eid_int(fi.Id))
        conns = iter_connectors(fi)
        for c in conns:
            sys_obj = None

            # Mechanical/Piping
            try:
                sys_obj = getattr(c, "MEPSystem", None)
            except Exception:
                sys_obj = None

            # Electrical (some builds expose ElectricalSystem)
            if sys_obj is None:
                try:
                    sys_obj = getattr(c, "ElectricalSystem", None)
                except Exception:
                    sys_obj = None

            if sys_obj is None:
                continue

            skey, sname, sprops = system_key_and_label(sys_obj)

            # Create system node
            add_node(
                skey, "system", sname, el=(sys_obj if hasattr(sys_obj, "UniqueId") else None),
                props=sprops,
                pos=next_pos("system")
            )

            add_edge("equip_to_system", ekey, skey, props={})
    except Exception:
        continue


# ----------------------------
# 5) View -> Rooms (phase-1: rooms on same GenLevel)
# ----------------------------
for vid_int, view in views_by_id.items():
    try:
        vkey = node_key("view", vid_int)
        lvlid = get_level_id_for_view(view)
        if not lvlid or lvlid == ElementId.InvalidElementId:
            continue

        lvl_int = eid_int(lvlid)
        if lvl_int is None:
            continue

        for room in rooms_by_level.get(lvl_int, []):
            rkey = node_key("room", eid_int(room.Id))
            add_edge("view_to_room", vkey, rkey, props={"method": "GenLevel"})
    except Exception:
        continue


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
