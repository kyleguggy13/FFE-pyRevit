# -*- coding: utf-8 -*-
__title__     = "Tag Receptacles"
__version__   = "Version = v0.1"
__doc__       = """Version = v0.1
Date    = 11.13.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [11.13.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Joshua Zolla/Kyle Guggenheim"""

# ____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
from collections import defaultdict, deque
import cgi
import time


# ____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.ExtensibleStorage import Schema
from Autodesk.Revit.DB.Mechanical import DuctSystemType


# ____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, UI, script, forms
from pyrevit.script import output


# ____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document
selection   = uidoc.Selection

output_window = output.get_output()




# ____________________________________________________________________ DYNAMO SCRIPT
import clr
import math
import traceback

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("RevitServices")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    ElementId, FilteredElementCollector, BuiltInCategory,
    BuiltInParameter, FamilySymbol, IndependentTag,
    Reference, XYZ, LocationPoint, TagOrientation,
    LeaderEndCondition, ViewType, ViewSheet, StorageType,
    SubTransaction, SpatialElementBoundaryOptions,
    RevitLinkInstance
)
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


# ==========================================================
# FFE AUTOMATIC RECEPTACLE TAGGING - VERSION 8.3
# Revit 2026 / Dynamo CPython3
#
# IN[0] LA tag family name
# IN[1] Write True/False
# IN[2] Landing length - paper inches
# IN[3] Diagonal offset - paper inches
#
# ALL FFE receptacles receive a tag:
#
# Standard, 18 AFF:
#     Panel/Circuit
# Standard, other height:
#     Panel/Circuit + AFF
# Nonstandard, 18 AFF:
#     Type Name
# Nonstandard, other height:
#     Type Name + AFF
#
# Features:
# - Automatic LA/RA selection
# - Group-based collision optimization
# - Avoid crossing leaders and electrical equipment
# - Room-aware secondary preference
# - Preview/Write mode
# - Log all changes
# - Select newly created tags after execution
#
# Existing tags are not modified.
# ==========================================================


def main():

    doc = DocumentManager.Instance.CurrentDBDocument
    active = doc.ActiveView

    LA_FAMILY = str(IN[0]).strip()
    RA_FAMILY = LA_FAMILY.replace("(LA)", "(RA)")

    WRITE = IN[1] == True
    LANDING = float(IN[2])
    OFFSET = float(IN[3])

    STANDARD_HEIGHT = 18.0
    HEIGHT_TOLERANCE = 0.1

    LA_CORRECTION = 0.0747
    RA_CORRECTION = 0.0747
    DOUBLE_CORRECTION = 0.0747

    # Approximate annotation dimensions (paper inches)
    CHAR_WIDTH = 0.085
    TEXT_HEIGHT = 0.17
    TEXT_PADDING = 0.04
    CLEARANCE = 0.035

    DEVICE_RADIUS = 0.15

    NEIGHBOR_DISTANCE = 1.0
    NEIGHBOR_VERTICAL = 0.40

    ROOM_PROBES = [0.0, 0.5, 1.0, 2.0]
    ROOM_MARGIN = 0.5

    PASSES = 16
    PAIR_PASSES = 3

    # Collision priorities
    COST_CROSSING = 1200.0
    COST_TEXT_OVERLAP = 1000.0
    COST_TEXT_DEVICE = 900.0
    COST_LEADER_DEVICE = 850.0
    COST_LEADER_TEXT = 800.0
    COST_INWARD = 120.0
    COST_ROOM = 2.0
    COST_RA_DEFAULT = 0.5

    if "(LA)" not in LA_FAMILY:
        return ["ERROR: Supply the LA tag family name."]

    if LANDING <= 0 or OFFSET <= 0:
        return ["ERROR: Leader dimensions must be positive."]

    logs = []
    errors = []
    jobs = []

    created_ids = List[ElementId]()

    scanned = 0
    skipped = 0
    created = 0
    la_count = 0
    ra_count = 0
    review_count = 0

    type_counts = {
        "PANEL/CIRCUIT": 0,
        "PANEL/CIRCUIT + AFF": 0,
        "TYPE NAME": 0,
        "TYPE NAME + AFF": 0
    }

    # ======================================================
    # BASIC HELPERS
    # ======================================================

    def eid(element):
        return element.Id.Value

    def type_name(element):
        p = element.Symbol.get_Parameter(
            BuiltInParameter.SYMBOL_NAME_PARAM
        )
        return (p.AsString() or "") if p else ""

    def height_inches(element):
        p = element.LookupParameter("Elevation from Level")

        if p and p.StorageType == StorageType.Double:
            return p.AsDouble() * 12.0

        return None

    def position(element, view):
        loc = element.Location

        if isinstance(loc, LocationPoint):
            return loc.Point

        box = element.get_BoundingBox(view)

        if box:
            return XYZ(
                (box.Min.X + box.Max.X) / 2.0,
                (box.Min.Y + box.Max.Y) / 2.0,
                (box.Min.Z + box.Max.Z) / 2.0
            )

        return None

    def model_distance(inches, view):
        return inches * view.Scale / 12.0

    def xy(point, view):
        factor = 12.0 / float(view.Scale)

        return (
            point.DotProduct(view.RightDirection) * factor,
            point.DotProduct(view.UpDirection) * factor
        )

    def rect_overlap(a, b):
        return (
            a[0] < b[2] and b[0] < a[2] and
            a[1] < b[3] and b[1] < a[3]
        )

    def point_in_rect(p, r):
        return (
            r[0] <= p[0] <= r[2] and
            r[1] <= p[1] <= r[3]
        )

    def expand_rect(r, gap):
        return (
            r[0] - gap,
            r[1] - gap,
            r[2] + gap,
            r[3] + gap
        )

    # ======================================================
    # LINE GEOMETRY
    # ======================================================

    def cross(a, b, c):
        return (
            (b[0] - a[0]) * (c[1] - a[1])
            - (b[1] - a[1]) * (c[0] - a[0])
        )

    def point_segment_distance(p, a, b):
        dx = b[0] - a[0]
        dy = b[1] - a[1]

        length2 = dx * dx + dy * dy

        if length2 < 1e-12:
            return math.hypot(
                p[0] - a[0],
                p[1] - a[1]
            )

        t = (
            (p[0] - a[0]) * dx +
            (p[1] - a[1]) * dy
        ) / length2

        t = max(0.0, min(1.0, t))

        return math.hypot(
            p[0] - (a[0] + dx * t),
            p[1] - (a[1] + dy * t)
        )

    def segments_intersect(a, b, c, d):
        eps = 1e-9

        p = cross(a, b, c)
        q = cross(a, b, d)
        r = cross(c, d, a)
        s = cross(c, d, b)

        return p * q < -eps and r * s < -eps

    def line_hits_rect(a, b, rect):
        if point_in_rect(a, rect):
            return True

        if point_in_rect(b, rect):
            return True

        x1, y1, x2, y2 = rect

        corners = [
            (x1, y1),
            (x2, y1),
            (x2, y2),
            (x1, y2)
        ]

        for i in range(4):
            if segments_intersect(
                a, b,
                corners[i],
                corners[(i + 1) % 4]
            ):
                return True

        return False

    def leader_hits_rect(option, rect):
        for a, b in option["segments"]:
            if line_hits_rect(a, b, rect):
                return True
        return False

    def leader_crosses(a, b):
        for seg_a in a["segments"]:
            for seg_b in b["segments"]:
                if segments_intersect(
                    seg_a[0], seg_a[1],
                    seg_b[0], seg_b[1]
                ):
                    return True
        return False

    # ======================================================
    # TEXT AND ELEMENT BOUNDING BOXES
    # ======================================================

    def box_rect(element, view):
        try:
            box = element.get_BoundingBox(view)

            if box is None:
                return None

            corners = [
                xy(XYZ(x, y, z), view)
                for x in [box.Min.X, box.Max.X]
                for y in [box.Min.Y, box.Max.Y]
                for z in [box.Min.Z, box.Max.Z]
            ]

            xs = [p[0] for p in corners]
            ys = [p[1] for p in corners]

            return (
                min(xs), min(ys),
                max(xs), max(ys)
            )
        except:
            return None

    def text_lines(typ, height, tag_type):
        total = int(round(height))

        aff = "{}' - {}\" AFF".format(
            total // 12,
            total % 12
        )

        # Panel/Circuit values can vary in length.
        # These are conservative placeholders used
        # only for collision prediction.
        # Revit's tag family determines actual text.

        if tag_type == "PANEL/CIRCUIT":
            return ["PANEL-CIRCUIT-00"]

        if tag_type == "PANEL/CIRCUIT + AFF":
            return [
                "PANEL-CIRCUIT-00",
                aff
            ]

        if tag_type == "TYPE NAME":
            return [typ + " :"]

        return [
            typ + " :",
            aff
        ]

    def proposed_rect(head, direction, lines, view):
        x, y = xy(head, view)

        width = (
            max(len(s) for s in lines) * CHAR_WIDTH
            + 2 * TEXT_PADDING
        )

        height = (
            len(lines) * TEXT_HEIGHT
            + 2 * TEXT_PADDING
        )

        if direction == "LA":
            rect = (
                x,
                y - height / 2,
                x + width,
                y + height / 2
            )
        else:
            rect = (
                x - width,
                y - height / 2,
                x,
                y + height / 2
            )

        return expand_rect(rect, CLEARANCE)

    # ======================================================
    # TARGET VIEWS
    # ======================================================

    allowed = [
        ViewType.FloorPlan,
        ViewType.EngineeringPlan
    ]

    views = []

    if isinstance(active, ViewSheet):
        for vp_id in active.GetAllViewports():
            vp = doc.GetElement(vp_id)
            v = doc.GetElement(vp.ViewId)

            if v and v.ViewType in allowed:
                views.append(v)

    elif active.ViewType in allowed:
        views = [active]

    else:
        return [
            "ERROR: Open a floor plan or sheet."
        ]

    if not views:
        return ["ERROR: No eligible floor plans."]

    # ======================================================
    # FIND ALL FOUR TAG TYPES IN LA AND RA
    # ======================================================

    names = {
        "LA": LA_FAMILY,
        "RA": RA_FAMILY
    }

    tag_types = {
        "LA": {},
        "RA": {}
    }

    for symbol in (
        FilteredElementCollector(doc)
        .OfClass(FamilySymbol)
        .ToElements()
    ):
        try:
            family = symbol.Family.Name

            for direction in ["LA", "RA"]:
                if family != names[direction]:
                    continue

                p = symbol.get_Parameter(
                    BuiltInParameter.SYMBOL_NAME_PARAM
                )

                if p and p.AsString():
                    key = p.AsString().strip().upper()
                    tag_types[direction][key] = symbol

        except:
            pass

    # NEW: Both Panel/Circuit types are now required.
    required = [
        "PANEL/CIRCUIT",
        "PANEL/CIRCUIT + AFF",
        "TYPE NAME",
        "TYPE NAME + AFF"
    ]

    missing = []

    for direction in ["LA", "RA"]:
        for key in required:
            if key not in tag_types[direction]:
                missing.append(
                    direction + " / " + key
                )

    if missing:
        return [
            "ERROR: Missing tag types.",
            ", ".join(missing)
        ]

    # ======================================================
    # ROOM DATA - SECONDARY PREFERENCE
    # ======================================================

    room_sources = []

    for room in (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    ):
        try:
            if room.Area > 0:
                room_sources.append(
                    (room, None)
                )
        except:
            pass

    for link in (
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
        .ToElements()
    ):
        try:
            link_doc = link.GetLinkDocument()

            if link_doc is None:
                continue

            transform = link.GetTotalTransform()

            for room in (
                FilteredElementCollector(link_doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToElements()
            ):
                if room.Area > 0:
                    room_sources.append(
                        (room, transform)
                    )

        except Exception as ex:
            errors.append(
                "LINK: " + str(ex)
            )

    room_cache = []
    boundary_options = SpatialElementBoundaryOptions()

    for room, transform in room_sources:
        try:
            loops = room.GetBoundarySegments(
                boundary_options
            )

            if not loops:
                continue

            segments = []

            for loop in loops:
                for segment in loop:
                    points = list(
                        segment.GetCurve().Tessellate()
                    )

                    if transform is not None:
                        points = [
                            transform.OfPoint(p)
                            for p in points
                        ]

                    for i in range(len(points) - 1):
                        segments.append((
                            points[i],
                            points[i + 1]
                        ))

            if segments:
                room_cache.append(
                    (room, transform, segments)
                )
        except:
            pass

    def inside_room(room, transform, p):
        try:
            local = (
                transform.Inverse.OfPoint(p)
                if transform is not None else p
            )

            loc = room.Location

            z = (
                loc.Point.Z + 3.0
                if loc else local.Z
            )

            return room.IsPointInRoom(
                XYZ(local.X, local.Y, z)
            )
        except:
            return False

    def find_room(point, view):
        offsets = [(0, 0)]

        for d in ROOM_PROBES[1:]:
            offsets.extend([
                (d, 0),
                (-d, 0),
                (0, d),
                (0, -d)
            ])

        for dx, dy in offsets:
            p = point.Add(
                view.RightDirection.Multiply(dx)
            ).Add(
                view.UpDirection.Multiply(dy)
            )

            for room, transform, segments in room_cache:
                if inside_room(room, transform, p):
                    return segments, p

        return None, None

    def room_distance(origin, sign, segments, view):
        vr = view.RightDirection
        vu = view.UpDirection

        nearest = None

        for a, b in segments:
            av = a.Subtract(origin)
            bv = b.Subtract(origin)

            ax = av.DotProduct(vr)
            ay = av.DotProduct(vu)

            sx = bv.DotProduct(vr) - ax
            sy = bv.DotProduct(vu) - ay

            denom = sign * sy

            if abs(denom) < 1e-9:
                continue

            t = (ax * sy - ay * sx) / denom
            u = -ay / sy

            if t > 0.001 and 0 <= u <= 1:
                if nearest is None or t < nearest:
                    nearest = t

        return nearest

    def room_preference(point, view):
        segments, probe = find_room(
            point, view
        )

        if segments is None:
            return "LA", "No room"

        left = room_distance(
            probe, -1, segments, view
        )

        right = room_distance(
            probe, 1, segments, view
        )

        if left is None or right is None:
            return "LA", "Incomplete room"

        direction = (
            "RA"
            if left > right + ROOM_MARGIN
            else "LA"
        )

        return direction, "Room clearance"

    # ======================================================
    # PROPOSED TAG GEOMETRY
    # ======================================================

    def make_option(item, direction, view):
        sign = 1.0 if direction == "LA" else -1.0

        diagonal = model_distance(
            OFFSET, view
        )

        landing = model_distance(
            LANDING, view
        )

        if item["tag_type"] in [
            "TYPE NAME + AFF",
            "PANEL/CIRCUIT + AFF"
        ]:
            correction = DOUBLE_CORRECTION

        elif direction == "LA":
            correction = LA_CORRECTION

        else:
            correction = RA_CORRECTION

        point = item["point"]

        vr = view.RightDirection
        vu = view.UpDirection

        elbow = point.Add(
            vr.Multiply(diagonal * sign)
        ).Add(
            vu.Multiply(diagonal)
        )

        head = elbow.Add(
            vr.Multiply(landing * sign)
        ).Add(
            vu.Multiply(
                model_distance(correction, view)
            )
        )

        origin_xy = xy(point, view)
        elbow_xy = xy(elbow, view)

        landing_xy = (
            elbow_xy[0] + LANDING * sign,
            elbow_xy[1]
        )

        return {
            "direction": direction,
            "symbol": tag_types[direction][
                item["tag_type"]
            ],
            "elbow": elbow,
            "head": head,
            "origin_xy": origin_xy,
            "rect": proposed_rect(
                head,
                direction,
                item["lines"],
                view
            ),
            "segments": [
                (origin_xy, elbow_xy),
                (elbow_xy, landing_xy)
            ]
        }

    # ======================================================
    # COLLISION SCORING
    # ======================================================

    def static_score(item, option, tags, equipment):
        score = 0.0
        issues = []

        for rect, tag_id in tags:

            if rect_overlap(option["rect"], rect):
                score += COST_TEXT_OVERLAP
                issues.append(
                    "Existing tag overlap " + str(tag_id)
                )

            if leader_hits_rect(option, rect):
                score += COST_LEADER_TEXT
                issues.append(
                    "Leader near tag " + str(tag_id)
                )

        for obstacle in equipment:
            if obstacle["id"] == item["id"]:
                continue

            rect = obstacle["rect"]

            if rect_overlap(option["rect"], rect):
                score += COST_TEXT_DEVICE
                issues.append(
                    "Text over electrical element " +
                    str(obstacle["id"])
                )

            if leader_hits_rect(option, rect):
                score += COST_LEADER_DEVICE
                issues.append(
                    "Leader over electrical element " +
                    str(obstacle["id"])
                )

        if option["direction"] != item["preferred"]:
            score += COST_ROOM

        if (
            option["direction"] == "RA"
            and item["room_info"] == "No room"
        ):
            score += COST_RA_DEFAULT

        return score, issues

    def pair_score(a, oa, b, ob):
        score = 0.0
        issues = []

        if leader_crosses(oa, ob):
            score += COST_CROSSING
            issues.append("LEADER CROSSING")

        if rect_overlap(oa["rect"], ob["rect"]):
            score += COST_TEXT_OVERLAP
            issues.append("TAG TEXT OVERLAP")

        if leader_hits_rect(oa, ob["rect"]):
            score += COST_LEADER_TEXT
            issues.append("A leader crosses B text")

        if leader_hits_rect(ob, oa["rect"]):
            score += COST_LEADER_TEXT
            issues.append("B leader crosses A text")

        pa = oa["origin_xy"]
        pb = ob["origin_xy"]

        dx = pb[0] - pa[0]
        dy = abs(pb[1] - pa[1])

        if (
            math.hypot(dx, dy) < NEIGHBOR_DISTANCE
            and dy < NEIGHBOR_VERTICAL
            and abs(dx) > 0.01
        ):
            if dx > 0:
                left = oa
                right = ob
            else:
                left = ob
                right = oa

            if (
                left["direction"] == "LA"
                and right["direction"] == "RA"
            ):
                score += COST_INWARD
                issues.append(
                    "Inward-facing leaders"
                )

        return score, issues

    # ======================================================
    # GROUP OPTIMIZER
    # ======================================================

    def optimize(items, tags, equipment, view):
        n = len(items)

        if n == 0:
            return [], 0.0

        for item in items:
            item["options"] = {
                d: make_option(item, d, view)
                for d in ["LA", "RA"]
            }

            item["static"] = {
                d: static_score(
                    item,
                    item["options"][d],
                    tags,
                    equipment
                )
                for d in ["LA", "RA"]
            }

        pairs = {}

        for i in range(n):
            for j in range(i + 1, n):
                for di in ["LA", "RA"]:
                    for dj in ["LA", "RA"]:
                        pairs[(i, j, di, dj)] = pair_score(
                            items[i],
                            items[i]["options"][di],
                            items[j],
                            items[j]["options"][dj]
                        )

        def pair_value(i, di, j, dj):
            if i < j:
                return pairs[(i, j, di, dj)][0]

            return pairs[(j, i, dj, di)][0]

        def total_score(choices):
            score = sum(
                items[i]["static"][choices[i]][0]
                for i in range(n)
            )

            for i in range(n):
                for j in range(i + 1, n):
                    score += pair_value(
                        i, choices[i],
                        j, choices[j]
                    )

            return score

        def local_score(i, direction, choices):
            score = items[i]["static"][direction][0]

            for j in range(n):
                if i != j:
                    score += pair_value(
                        i, direction,
                        j, choices[j]
                    )

            return score

        starts = [
            ["LA"] * n,
            ["RA"] * n,
            [item["preferred"] for item in items],
            [
                "RA" if i % 2 == 0 else "LA"
                for i in range(n)
            ],
            [
                "LA" if i % 2 == 0 else "RA"
                for i in range(n)
            ]
        ]

        best = None
        best_value = float("inf")

        for start in starts:
            choices = list(start)

            for iteration in range(PASSES):
                changed = False
                order = list(range(n))

                if iteration % 2:
                    order.reverse()

                for i in order:
                    la = local_score(
                        i, "LA", choices
                    )

                    ra = local_score(
                        i, "RA", choices
                    )

                    desired = (
                        "RA"
                        if ra < la - 1e-7
                        else "LA"
                    )

                    if desired != choices[i]:
                        choices[i] = desired
                        changed = True

                if not changed:
                    break

            # Attempt to improve two directions together.
            for pair_pass in range(PAIR_PASSES):
                improved = False

                for i in range(n):
                    for j in range(i + 1, n):

                        current = total_score(choices)
                        best_pair = None
                        pair_best = current

                        old_i = choices[i]
                        old_j = choices[j]

                        for di in ["LA", "RA"]:
                            for dj in ["LA", "RA"]:

                                choices[i] = di
                                choices[j] = dj

                                trial = total_score(choices)

                                if trial < pair_best - 1e-7:
                                    pair_best = trial
                                    best_pair = (di, dj)

                        if best_pair is not None:
                            choices[i], choices[j] = best_pair
                            improved = True
                        else:
                            choices[i] = old_i
                            choices[j] = old_j

                if not improved:
                    break

            value = total_score(choices)

            if value < best_value - 1e-7:
                best_value = value
                best = list(choices)

        results = []

        for i, item in enumerate(items):
            direction = best[i]
            issues = list(
                item["static"][direction][1]
            )

            for j in range(n):
                if i == j:
                    continue

                if i < j:
                    pair = pairs[(
                        i, j, direction, best[j]
                    )]
                else:
                    pair = pairs[(
                        j, i, best[j], direction
                    )]

                for issue in pair[1]:
                    issues.append(
                        "{} with {}".format(
                            issue,
                            items[j]["id"]
                        )
                    )

            results.append({
                "item": item,
                "direction": direction,
                "option": item["options"][direction],
                "issues": sorted(set(issues))
            })

        return results, best_value

    # ======================================================
    # PROCESS EACH FLOOR PLAN
    # ======================================================

    for view in views:

        logs.append(
            "VIEW {} | 1:{}".format(
                view.Name,
                view.Scale
            )
        )

        receptacles = list(
            FilteredElementCollector(doc, view.Id)
            .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        electrical_equipment = list(
            FilteredElementCollector(doc, view.Id)
            .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        scanned += len(receptacles)

        tagged_ids = set()
        existing_tags = []

        for tag in (
            FilteredElementCollector(doc, view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        ):
            try:
                for tag_id in tag.GetTaggedLocalElementIds():
                    tagged_ids.add(tag_id.Value)
            except:
                pass

            rect = box_rect(tag, view)

            if rect:
                existing_tags.append((
                    expand_rect(rect, CLEARANCE),
                    eid(tag)
                ))

        # Protect all electrical fixtures and equipment.
        obstacles = []

        for element in receptacles + electrical_equipment:
            rect = box_rect(element, view)
            p = position(element, view)

            if rect:
                if p:
                    cx, cy = xy(p, view)
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]

                    if width > 1.5 or height > 1.5:
                        rect = (
                            cx - DEVICE_RADIUS,
                            cy - DEVICE_RADIUS,
                            cx + DEVICE_RADIUS,
                            cy + DEVICE_RADIUS
                        )

                rect = expand_rect(
                    rect,
                    DEVICE_RADIUS * 0.25
                )

                obstacles.append({
                    "id": eid(element),
                    "rect": rect
                })

            elif p:
                cx, cy = xy(p, view)

                obstacles.append({
                    "id": eid(element),
                    "rect": (
                        cx - DEVICE_RADIUS,
                        cy - DEVICE_RADIUS,
                        cx + DEVICE_RADIUS,
                        cy + DEVICE_RADIUS
                    )
                })

        items = []

        for element in receptacles:
            element_id = eid(element)

            if element_id in tagged_ids:
                skipped += 1

                logs.append(
                    "SKIP {} | Already tagged".format(
                        element_id
                    )
                )
                continue

            family = element.Symbol.Family.Name
            typ = type_name(element)

            if not family.upper().startswith(
                "FFE_EF_RECEPTACLES"
            ):
                skipped += 1
                continue

            special = "SPECIAL" in family.upper()
            standard = "STANDARD" in typ.upper()

            nonstandard_type = special or not standard

            height = height_inches(element)

            if height is None:
                errors.append(
                    "MISSING HEIGHT {}".format(
                        element_id
                    )
                )
                continue

            nonstandard_height = (
                abs(height - STANDARD_HEIGHT)
                > HEIGHT_TOLERANCE
            )

            # ==================================================
            # NEW VERSION 8.3 TAG CLASSIFICATION
            # ==================================================
            #
            # Every FFE receptacle with valid height
            # now receives one of four tag types.
            #
            # Standard:
            #   18 AFF  -> PANEL/CIRCUIT
            #   Other   -> PANEL/CIRCUIT + AFF
            #
            # Nonstandard:
            #   18 AFF  -> TYPE NAME
            #   Other   -> TYPE NAME + AFF
            # ==================================================

            if not nonstandard_type:

                if nonstandard_height:
                    tag_type = "PANEL/CIRCUIT + AFF"
                else:
                    tag_type = "PANEL/CIRCUIT"

            else:

                if nonstandard_height:
                    tag_type = "TYPE NAME + AFF"
                else:
                    tag_type = "TYPE NAME"

            p = position(element, view)

            if p is None:
                errors.append(
                    "MISSING LOCATION {}".format(
                        element_id
                    )
                )
                continue

            preferred, room_info = room_preference(
                p, view
            )

            type_counts[tag_type] += 1

            items.append({
                "id": element_id,
                "element": element,
                "point": p,
                "tag_type": tag_type,
                "type_name": typ,
                "lines": text_lines(
                    typ,
                    height,
                    tag_type
                ),
                "preferred": preferred,
                "room_info": room_info
            })

        items.sort(
            key=lambda item: (
                xy(item["point"], view)[0],
                xy(item["point"], view)[1],
                item["id"]
            )
        )

        results, score = optimize(
            items,
            existing_tags,
            obstacles,
            view
        )

        logs.append(
            "OPTIMIZATION SCORE: {:.2f}".format(
                score
            )
        )

        for result in results:
            item = result["item"]
            direction = result["direction"]
            option = result["option"]
            issues = result["issues"]

            if direction == "LA":
                la_count += 1
            else:
                ra_count += 1

            serious = [
                s for s in issues
                if (
                    "CROSSING" in s or
                    "OVERLAP" in s or
                    "over electrical" in s or
                    "near electrical" in s or
                    "crosses" in s or
                    "near tag" in s
                )
            ]

            if serious:
                review_count += 1

            jobs.append((
                view,
                item["element"],
                option["symbol"],
                item["point"],
                option["elbow"],
                option["head"],
                direction,
                item["tag_type"],
                serious
            ))

            logs.append(
                "CANDIDATE {} | {} | {} | {}".format(
                    item["id"],
                    item["type_name"],
                    direction,
                    item["tag_type"]
                )
            )

            if serious:
                logs.append(
                    "  REVIEW: " +
                    "; ".join(serious[:5])
                )

    # ======================================================
    # WRITE NEW TAGS
    # ======================================================

    if WRITE and jobs:

        TransactionManager.Instance.EnsureInTransaction(doc)

        try:
            for job in jobs:
                (
                    view,
                    element,
                    symbol,
                    point,
                    elbow,
                    head,
                    direction,
                    tag_type,
                    issues
                ) = job

                sub = SubTransaction(doc)

                try:
                    sub.Start()

                    if not symbol.IsActive:
                        symbol.Activate()
                        doc.Regenerate()

                    ref = Reference(element)

                    tag = IndependentTag.Create(
                        doc,
                        symbol.Id,
                        view.Id,
                        ref,
                        True,
                        TagOrientation.Horizontal,
                        head
                    )

                    doc.Regenerate()

                    tag.TagHeadPosition = head
                    doc.Regenerate()

                    if tag.CanLeaderEndConditionBeAssigned(
                        LeaderEndCondition.Free
                    ):
                        tag.LeaderEndCondition = (
                            LeaderEndCondition.Free
                        )

                        tag.SetLeaderEnd(
                            ref,
                            point
                        )

                    tag.SetLeaderElbow(
                        ref,
                        elbow
                    )

                    doc.Regenerate()
                    sub.Commit()

                    created_ids.Add(tag.Id)
                    created += 1

                    logs.append(
                        "CREATED | Element {} | Tag {} | "
                        "{} | {}".format(
                            eid(element),
                            eid(tag),
                            direction,
                            tag_type
                        )
                    )

                except Exception as ex:
                    try:
                        sub.RollBack()
                    except:
                        pass

                    errors.append(
                        "ERROR {} | {}".format(
                            eid(element),
                            str(ex)
                        )
                    )

        finally:
            TransactionManager.Instance.TransactionTaskDone()

    # ======================================================
    # AUTOMATICALLY SELECT NEWLY CREATED TAGS
    # ======================================================

    selected_count = 0

    if WRITE and created_ids.Count > 0:

        try:
            uiapp = (
                DocumentManager.Instance.CurrentUIApplication
            )

            uidoc = uiapp.ActiveUIDocument

            uidoc.Selection.SetElementIds(
                created_ids
            )

            selected_count = created_ids.Count

            logs.append(
                "SELECTED {} newly created tags".format(
                    selected_count
                )
            )

        except Exception as ex:
            errors.append(
                "SELECTION WARNING: {}".format(ex)
            )

    # ======================================================
    # EXECUTION REPORT
    # ======================================================

    report = [
        "FFE RECEPTACLE TAGGING - VERSION 8.3",
        "======================================",
        "Mode: {}".format(
            "WRITE" if WRITE else "PREVIEW"
        ),
        "Views processed: {}".format(len(views)),
        "Room boundaries: {}".format(len(room_cache)),
        "Receptacles scanned: {}".format(scanned),
        "Tag candidates: {}".format(len(jobs)),
        "--------------------------------------",
        "Panel/Circuit: {}".format(
            type_counts["PANEL/CIRCUIT"]
        ),
        "Panel/Circuit + AFF: {}".format(
            type_counts["PANEL/CIRCUIT + AFF"]
        ),
        "Type Name: {}".format(
            type_counts["TYPE NAME"]
        ),
        "Type Name + AFF: {}".format(
            type_counts["TYPE NAME + AFF"]
        ),
        "--------------------------------------",
        "LA candidates: {}".format(la_count),
        "RA candidates: {}".format(ra_count),
        "Manual review candidates: {}".format(
            review_count
        ),
        "Tags created: {}".format(created),
        "Tags selected: {}".format(selected_count),
        "Skipped: {}".format(skipped),
        "Errors/warnings: {}".format(len(errors)),
        "======================================",
        "DETAIL LOG"
    ]

    report.extend(logs)
    report.append("======================================")
    report.append("ERRORS / WARNINGS")
    report.extend(errors)

    return report


try:
    OUT = main()

except Exception:
    OUT = [
        "FATAL SCRIPT ERROR",
        traceback.format_exc()
    ]
