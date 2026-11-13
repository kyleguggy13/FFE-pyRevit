# -*- coding: utf-8 -*-
"""Convert the original v8.3 receptacle workflow into a standalone command.

Key behaviors: report-only Preview, best-effort Write, existing-tag protection,
automatic text alignment, flexible leaders, and aligned stacks for nearby devices.
Revit API notes: targets Revit 2025/2026; uses ElementId.Value and native
transactions. Runs on pyRevit's default IronPython engine with WPF forms.
Design decisions: refine the original layout with length variations and group
stacks. Stacks can vary leader slope and height. Write has one Undo step.
"""

__title__     = "Tag Receptacles"
__version__   = "Version = v0.3"
__doc__       = """Version = v0.3
Date    = 11.13.2026
______________________________________________________________
Description:
-> Automatically tag visible FFE receptacles using the Dynamo v8.3 rules.
-> Existing tags are not modified.
______________________________________________________________
How-to:
-> Open a floor/engineering plan or a sheet containing eligible plans.
-> Choose Preview or Write and paper-inch dimensions.
-> FFE_Tag_EF_Receptacle (LA)/(RA) is selected automatically per tag.
-> Enter preferred lengths; conflicting tags can use 75%-200% of those values.
-> Nearby receptacles can use aligned stacks with extra vertical spacing.
-> Review the results and any placement warnings in the pyRevit output.
______________________________________________________________
Last update:
- [11.13.2026] - v0.3 Aligned tag stacks for nearby receptacles.
- [11.13.2026] - v0.2 Flexible offset and landing lengths to reduce collisions.
- [11.13.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Joshua Zolla/Kyle Guggenheim"""

# ____________________________________________________________________ SETUP
import clr
import math
import os
import traceback

clr.AddReference("System")
clr.AddReference("RevitAPI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    ElementId, FilteredElementCollector, BuiltInCategory,
    BuiltInParameter, FamilySymbol, IndependentTag,
    Reference, XYZ, LocationPoint, TagOrientation,
    LeaderEndCondition, ViewType, ViewSheet, StorageType,
    SubTransaction, Transaction, TransactionGroup, TransactionStatus,
    SpatialElementBoundaryOptions,
    RevitLinkInstance
)
from pyrevit import forms, script


LA_TAG_FAMILY = "FFE_Tag_EF_Receptacle (LA)"
RA_TAG_FAMILY = "FFE_Tag_EF_Receptacle (RA)"
DEFAULT_LANDING = 0.074
DEFAULT_OFFSET = 0.375
OFFSET_FACTORS = (1.0, 0.75, 1.25, 1.5, 2.0)
LANDING_FACTORS = (1.0, 0.75, 1.5, 2.0)
STACK_MAX_GROUP_SIZE = 12


# ____________________________________________________________________ SETTINGS UI
def ValidateSettings(write, landing, offset):
    """Return validated inputs; dimensions are positive, finite paper inches."""
    dimensions = []
    for label, value in [("Landing length", landing), ("Diagonal offset", offset)]:
        try:
            number = float(value)
        except (ValueError, TypeError, OverflowError):
            raise ValueError("{} must be a positive number in paper inches.".format(label))
        if math.isnan(number) or math.isinf(number) or number <= 0:
            raise ValueError("{} must be a positive, finite number in paper inches.".format(label))
        dimensions.append(number)

    return {"write": bool(write),
            "landing": dimensions[0], "offset": dimensions[1]}


class SettingsDialog(forms.WPFWindow):
    """Collect mode and dimensions; text alignment and tag families are automatic."""

    def __init__(self):
        forms.WPFWindow.__init__(
            self, os.path.join(os.path.dirname(__file__), "Settings.xaml"))
        self.Settings = None
        self.landingText.Text = str(DEFAULT_LANDING)
        self.offsetText.Text = str(DEFAULT_OFFSET)
        self.writeMode.IsChecked = True

    def RunClick(self, sender, args):
        """Keep the dialog open when inputs need correction."""
        try:
            settings = ValidateSettings(
                self.writeMode.IsChecked,
                self.landingText.Text, self.offsetText.Text)
        except ValueError as ex:
            self.errorText.Text = str(ex)
            return
        self.Settings = settings
        self.Close()

    def CancelClick(self, sender, args):
        """Leave Settings unset so Cancel, Escape, and window close are no-ops."""
        self.Close()


# ____________________________________________________________________ TRANSACTIONS
def RollBackIfStarted(transaction):
    """Roll back an active scope; a rollback failure must stop the workflow."""
    if transaction is not None and transaction.GetStatus() == TransactionStatus.Started:
        if transaction.RollBack() != TransactionStatus.RolledBack:
            raise RuntimeError("Revit could not roll back the tagging transaction.")


def CreateTags(doc, jobs, logs, errors):
    """Create the original job geometry, publishing IDs only after group commit.

Each tag has a subtransaction so unsupported elements can be reported locally.
The outer group also protects against a failed final commit or assimilation.
"""
    createdIds = List[ElementId]()
    committedLogs = []
    group = TransactionGroup(doc, "Tag Receptacles")
    transaction = None
    try:
        if group.Start() != TransactionStatus.Started:
            raise RuntimeError("Could not start the Tag Receptacles transaction group.")
        transaction = Transaction(doc, "Create receptacle tags")
        if transaction.Start() != TransactionStatus.Started:
            raise RuntimeError("Could not start the receptacle tag transaction.")

        # Resolve any native Revit failure dialog before attempting to finish
        # the group; a modeless pending commit cannot be assimilated safely.
        options = transaction.GetFailureHandlingOptions()
        options.SetForcedModalHandling(True)
        transaction.SetFailureHandlingOptions(options)

        for job in jobs:
            (view, element, symbol, point, elbow, head,
             direction, tagType, issues) = job
            sub = SubTransaction(doc)
            try:
                if sub.Start() != TransactionStatus.Started:
                    raise RuntimeError("Could not start the tag subtransaction.")
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()

                reference = Reference(element)
                tag = IndependentTag.Create(
                    doc, symbol.Id, view.Id, reference, True,
                    TagOrientation.Horizontal, head)
                doc.Regenerate()
                tag.TagHeadPosition = head
                doc.Regenerate()
                if tag.CanLeaderEndConditionBeAssigned(LeaderEndCondition.Free):
                    tag.LeaderEndCondition = LeaderEndCondition.Free
                    tag.SetLeaderEnd(reference, point)
                tag.SetLeaderElbow(reference, elbow)
                doc.Regenerate()
                if sub.Commit() != TransactionStatus.Committed:
                    raise RuntimeError("Revit did not commit the tag subtransaction.")
            except Exception as ex:
                errors.append("ERROR {} | {}".format(element.Id.Value, ex))
                RollBackIfStarted(sub)
                continue
            finally:
                sub.Dispose()

            createdIds.Add(tag.Id)
            committedLogs.append(
                "CREATED | Element {} | Tag {} | {} | {}".format(
                    element.Id.Value, tag.Id.Value, direction, tagType))

        if transaction.Commit() != TransactionStatus.Committed:
            raise RuntimeError("Revit did not commit the receptacle tag transaction.")
        if group.Assimilate() != TransactionStatus.Committed:
            raise RuntimeError("Revit did not commit the Tag Receptacles transaction group.")
    except Exception:
        # Roll back the group even if closing the inner transaction fails.
        try:
            RollBackIfStarted(transaction)
        finally:
            RollBackIfStarted(group)
        raise
    finally:
        try:
            if transaction is not None:
                transaction.Dispose()
        finally:
            group.Dispose()

    logs.extend(committedLogs)
    return createdIds


# ==========================================================
# FFE AUTOMATIC RECEPTACLE TAGGING - VERSION 8.3
# Based on Dynamo v8.3 with bounded leader-length refinement; Revit 2025/2026.
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


def RunTagging(doc, uidoc, settings):
    """Classify receptacles and refine the v8.3 layout with flexible leader lengths."""

    active = doc.ActiveView

    settings = ValidateSettings(settings["write"],
                                settings["landing"], settings["offset"])
    LA_FAMILY = LA_TAG_FAMILY
    RA_FAMILY = RA_TAG_FAMILY

    WRITE = settings["write"]
    LANDING = settings["landing"]
    OFFSET = settings["offset"]

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
    FLEX_PASSES = 4
    COST_LENGTH_CHANGE = 4.0
    STACK_ROW_GAP = 0.06  # Paper inches between estimated text rectangles.
    COST_UNSTACKED_GROUP = 8.0  # Aesthetic preference, below any collision cost.
    COST_STACK_FAN_PREFERENCE = 0.5

    # Collision priorities
    COST_CROSSING = 1200.0
    COST_TEXT_OVERLAP = 1000.0
    COST_TEXT_DEVICE = 900.0
    COST_LEADER_DEVICE = 850.0
    COST_LEADER_TEXT = 800.0
    COST_INWARD = 120.0
    COST_ROOM = 2.0
    COST_RA_DEFAULT = 0.5

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
    adjusted_count = 0
    stacked_count = 0

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
            "Required families: {} and {}.".format(LA_FAMILY, RA_FAMILY),
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

    def make_option(item, direction, view, landing_inches=None, offset_inches=None,
                    head_xy=None):
        """Build one alignment/length option in view coordinates and paper inches."""
        if landing_inches is None:
            landing_inches = LANDING
        if offset_inches is None:
            offset_inches = OFFSET
        # LA/RA describe text alignment. A tag to the right of the device uses
        # LA so its leader meets the left text edge; a tag to the left uses RA.
        sign = 1.0 if direction == "LA" else -1.0

        diagonal = model_distance(
            offset_inches, view
        )

        landing = model_distance(
            landing_inches, view
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

        if head_xy is None:
            elbow = point.Add(vr.Multiply(diagonal * sign)).Add(vu.Multiply(diagonal))
        else:
            # A shared text column requires independent horizontal and vertical
            # offsets. Keep the original horizontal landing and family correction.
            origin = xy(point, view)
            elbow = point.Add(vr.Multiply(model_distance(
                head_xy[0] - origin[0] - landing_inches * sign, view
            ))).Add(vu.Multiply(model_distance(
                head_xy[1] - origin[1] - correction, view
            )))
            offset_inches = abs(head_xy[1] - origin[1] - correction)

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
            elbow_xy[0] + landing_inches * sign,
            elbow_xy[1]
        )
        horizontalOffset = abs(elbow_xy[0] - origin_xy[0])
        offsetDeparture = abs(offset_inches / OFFSET - 1.0)
        if head_xy is not None:
            offsetDeparture = (offsetDeparture + abs(horizontalOffset / OFFSET - 1.0)) / 2.0

        rect = proposed_rect(head, direction, item["lines"], view)
        # Include text and both leader segments. Padding also preserves the
        # original nearby-device penalty when leaders themselves do not touch.
        bounds = expand_rect((
            min(rect[0], origin_xy[0], elbow_xy[0], landing_xy[0]),
            min(rect[1], origin_xy[1], elbow_xy[1], landing_xy[1]),
            max(rect[2], origin_xy[0], elbow_xy[0], landing_xy[0]),
            max(rect[3], origin_xy[1], elbow_xy[1], landing_xy[1])
        ), NEIGHBOR_DISTANCE / 2.0 + 1e-9)

        return {
            "direction": direction,
            "landing": landing_inches,
            "offset": offset_inches,
            "horizontal_offset": horizontalOffset,
            "stacked": head_xy is not None,
            "length_cost": COST_LENGTH_CHANGE * (
                abs(landing_inches / LANDING - 1.0) +
                offsetDeparture),
            "bounds": bounds,
            "symbol": tag_types[direction][
                item["tag_type"]
            ],
            "elbow": elbow,
            "head": head,
            "origin_xy": origin_xy,
            "rect": rect,
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

        # Distant options cannot collide or incur the nearby-device penalty.
        if not rect_overlap(oa["bounds"], ob["bounds"]):
            return score, issues

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

        def PairLocalScore(i, di, j, dj, choices):
            """Score only terms changed by a pair move instead of the whole view."""
            score = (items[i]["static"][di][0] + items[j]["static"][dj][0]
                     + pair_value(i, di, j, dj))
            for k in range(n):
                if k != i and k != j:
                    score += pair_value(i, di, k, choices[k])
                    score += pair_value(j, dj, k, choices[k])
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

                        current = PairLocalScore(i, choices[i], j, choices[j], choices)
                        best_pair = None
                        pair_best = current

                        old_i = choices[i]
                        old_j = choices[j]

                        for di in ["LA", "RA"]:
                            for dj in ["LA", "RA"]:

                                trial = PairLocalScore(i, di, j, dj, choices)

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

        # Retain the preferred-length solution unless a candidate reduces its
        # collision cost. Avoid precomputing every combination of flexible
        # lengths for every pair: only compare against the current placements.
        selected = [items[i]["options"][best[i]] for i in range(n)]
        for i, item in enumerate(items):
            for direction in ("LA", "RA"):
                item["options"][direction]["static"] = item["static"][direction]

        def PlacementScore(i, option):
            """Compute the affected score and whether this tag has any conflict."""
            score = option["static"][0] + option["length_cost"]
            hasConflict = bool(option["static"][1])
            for j in range(n):
                if j == i:
                    continue
                value, issues = pair_score(items[i], option, items[j], selected[j])
                score += value
                hasConflict = hasConflict or bool(issues)
            return score, hasConflict

        def BuildFlexibleOptions(item):
            """Cache bounded alternatives only for tags involved in a conflict."""
            if "flex_options" not in item:
                options = []
                for offsetFactor in OFFSET_FACTORS:
                    for landingFactor in LANDING_FACTORS:
                        for direction in ("LA", "RA"):
                            if offsetFactor == 1.0 and landingFactor == 1.0:
                                option = item["options"][direction]
                            else:
                                option = make_option(
                                    item, direction, view,
                                    LANDING * landingFactor, OFFSET * offsetFactor)
                                option["static"] = static_score(item, option, tags, equipment)
                            options.append(option)
                item["flex_options"] = options
            return item["flex_options"]

        for iteration in range(FLEX_PASSES):
            changed = False
            order = list(range(n))
            if iteration % 2:
                order.reverse()
            for i in order:
                option = selected[i]
                currentScore, hasConflict = PlacementScore(i, option)
                if not hasConflict and option["length_cost"] == 0:
                    continue
                chosen = option
                for candidate in BuildFlexibleOptions(items[i]):
                    score, unused = PlacementScore(i, candidate)
                    if score < currentScore - 1e-7:
                        chosen, currentScore = candidate, score
                if chosen is not option:
                    selected[i] = chosen
                    changed = True
            if not changed:
                break

        def BuildNearbyGroups():
            """Find bounded groups in paper space without chaining along a whole wall."""
            positions = [xy(item["point"], view) for item in items]
            remaining = set(range(n))
            groups = []
            for seed in range(n):
                if seed not in remaining:
                    continue
                group = [seed]
                remaining.remove(seed)
                cursor = 0
                while cursor < len(group) and len(group) < STACK_MAX_GROUP_SIZE:
                    anchor = positions[group[cursor]]
                    for j in sorted(remaining):
                        point = positions[j]
                        if math.hypot(point[0] - anchor[0], point[1] - anchor[1]) > NEIGHBOR_DISTANCE:
                            continue
                        xs = [positions[k][0] for k in group] + [point[0]]
                        ys = [positions[k][1] for k in group] + [point[1]]
                        if max(xs) - min(xs) > 1.5 or max(ys) - min(ys) > 1.5:
                            continue
                        group.append(j)
                        remaining.remove(j)
                        if len(group) >= STACK_MAX_GROUP_SIZE:
                            break
                    cursor += 1
                if len(group) > 1:
                    groups.append(tuple(sorted(group)))
            return groups, positions

        groups, positions = BuildNearbyGroups()
        stackCache = {}

        def StackStyleCost(group, options):
            """Prefer compact columns over splitting a pair into two singleton stacks."""
            if not all(options[i]["stacked"] for i in group):
                return COST_UNSTACKED_GROUP * len(group)
            columns = {}
            for i in group:
                headX, headY = xy(options[i]["head"], view)
                key = (options[i]["direction"], round(headX, 8))
                columns.setdefault(key, []).append((headY, options[i]["rect"]))
            score = 6.0 * (len(columns) - 1)
            # Favor the upper-right / lower-left fan-out from the reference
            # arrangement when room clearance and collision scores are comparable.
            for i in group:
                vertical = (xy(options[i]["elbow"], view)[1] - positions[i][1])
                if ((options[i]["direction"] == "LA" and vertical < 0)
                        or (options[i]["direction"] == "RA" and vertical > 0)):
                    score += COST_STACK_FAN_PREFERENCE
            for rows in columns.values():
                rows.sort(key=lambda row: row[0])
                for k in range(1, len(rows)):
                    lowerY, lowerRect = rows[k - 1]
                    upperY, upperRect = rows[k]
                    expectedGap = ((lowerRect[3] - lowerRect[1] + upperRect[3] - upperRect[1])
                                   / 2.0 + STACK_ROW_GAP)
                    score += COST_UNSTACKED_GROUP * max(0.0, (upperY - lowerY) / expectedGap - 1.0)
            return score

        def GroupScore(group, options):
            """Score all changed static and pair terms once, including outside tags."""
            members = set(group)
            score = sum(options[i]["static"][0] + options[i]["length_cost"] for i in group)
            for i in group:
                for j in range(n):
                    if j == i or (j in members and j < i):
                        continue
                    other = options[j] if j in members else selected[j]
                    score += pair_score(items[i], options[i], items[j], other)[0]
            # Prefer an ordered column over scattered but equally clear tags.
            # Even a small collision costs far more than this style preference.
            return score + StackStyleCost(group, options)

        def BuildStack(group, members, direction, verticalSign, distanceFactor):
            """Align text heads, space rows by label height, and fan leaders in order."""
            key = (group, tuple(members), direction, verticalSign, distanceFactor)
            if key in stackCache:
                return stackCache[key]
            sign = 1.0 if direction == "LA" else -1.0
            xs = [positions[i][0] for i in group]
            columnX = ((max(xs) if sign > 0 else min(xs))
                       + sign * (OFFSET * distanceFactor + LANDING))
            # Rank by a 45-degree routing projection, rather than element ID,
            # so adjacent devices fan toward their rows without swapping order.
            ordered = sorted(members, key=lambda i: (
                -(positions[i][1] - verticalSign * sign * positions[i][0]),
                items[i]["id"]))
            rowOrder = list(reversed(ordered)) if verticalSign > 0 else ordered
            first = rowOrder[0]
            base = items[first]["options"][direction]
            correction = xy(base["head"], view)[1] - xy(base["elbow"], view)[1]
            edgeY = (max(positions[i][1] for i in group) if verticalSign > 0
                     else min(positions[i][1] for i in group))
            centerY = edgeY + verticalSign * OFFSET * distanceFactor + correction
            previousHeight = None
            options = {}
            for i in rowOrder:
                rect = items[i]["options"][direction]["rect"]
                height = rect[3] - rect[1]
                if previousHeight is not None:
                    centerY += verticalSign * ((previousHeight + height) / 2.0 + STACK_ROW_GAP)
                option = make_option(items[i], direction, view, head_xy=(columnX, centerY))
                option["static"] = static_score(items[i], option, tags, equipment)
                options[i] = option
                previousHeight = height
            stackCache[key] = options
            return options

        def StackLayouts(group):
            """Try a single stack and split stacks above/below or left/right."""
            xOrder = sorted(group, key=lambda i: (positions[i][0], -positions[i][1], items[i]["id"]))
            yOrder = sorted(group, key=lambda i: (-positions[i][1], positions[i][0], items[i]["id"]))
            risingOrder = sorted(group, key=lambda i: (-(positions[i][1] + positions[i][0]), items[i]["id"]))
            fallingOrder = sorted(group, key=lambda i: (-(positions[i][1] - positions[i][0]), items[i]["id"]))

            def SplitCoordinate(i, axis):
                if axis < 2:
                    return positions[i][axis]
                return positions[i][1] + (1 if axis == 2 else -1) * positions[i][0]

            for distanceFactor in (1.0, 1.5, 2.0):
                for direction in ("LA", "RA"):
                    for verticalSign in (1, -1):
                        yield BuildStack(group, group, direction, verticalSign, distanceFactor)
                for order, axis in ((xOrder, 0), (yOrder, 1), (risingOrder, 2), (fallingOrder, 3)):
                    if len(group) <= 6:
                        cuts = range(1, len(group))
                    else:
                        # Bound the template count for larger groups: balanced
                        # splits and the largest physical gap cover common rows.
                        largestGap = max(range(1, len(group)), key=lambda k:
                            abs(SplitCoordinate(order[k], axis) - SplitCoordinate(order[k - 1], axis)))
                        cuts = sorted(set((len(group) // 2, largestGap)))
                    for cut in cuts:
                        first, second = order[:cut], order[cut:]
                        for a in (1, -1):
                            for b in (1, -1):
                                if axis == 0:
                                    left = BuildStack(group, first, "RA", a, distanceFactor)
                                    right = BuildStack(group, second, "LA", b, distanceFactor)
                                else:
                                    left = BuildStack(group, first, "LA" if a > 0 else "RA", 1, distanceFactor)
                                    right = BuildStack(group, second, "LA" if b > 0 else "RA", -1, distanceFactor)
                                options = dict(left)
                                options.update(right)
                                yield options

        # Move a whole nearby group together: individually moving one tag can
        # look worse until its neighbors also occupy their assigned stack rows.
        for stackPass in range(2):
            changed = False
            for group in groups:
                chosen = {i: selected[i] for i in group}
                currentScore = GroupScore(group, chosen)
                for candidate in StackLayouts(group):
                    score = GroupScore(group, candidate)
                    if score < currentScore - 1e-7:
                        chosen, currentScore = candidate, score
                if any(selected[i] is not chosen[i] for i in group):
                    for i in group:
                        selected[i] = chosen[i]
                    changed = True
            if not changed:
                break

        results = []
        best_value = sum(option["static"][0] + option["length_cost"]
                         for option in selected)
        for i in range(n):
            for j in range(i + 1, n):
                best_value += pair_score(items[i], selected[i], items[j], selected[j])[0]
        for group in groups:
            best_value += StackStyleCost(group, {i: selected[i] for i in group})

        for i, item in enumerate(items):
            option = selected[i]
            direction = option["direction"]
            issues = list(option["static"][1])

            for j in range(n):
                if i == j:
                    continue

                if i < j:
                    pair = pair_score(item, option, items[j], selected[j])
                else:
                    pair = pair_score(items[j], selected[j], item, option)

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
                "option": option,
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

            if (abs(option["landing"] - LANDING) > 1e-9
                    or abs(option["offset"] - OFFSET) > 1e-9
                    or abs(option["horizontal_offset"] - OFFSET) > 1e-9):
                adjusted_count += 1
            if option["stacked"]:
                stacked_count += 1

            if direction == "LA":
                la_count += 1
            else:
                ra_count += 1

            serious = [
                s for s in issues
                if (
                    "CROSSING" in s or
                    "overlap" in s.lower() or
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
                "CANDIDATE {} | {} | {} | {} | Landing {:.4f} in | Offset {:.4f} in{}".format(
                    item["id"],
                    item["type_name"],
                    direction,
                    item["tag_type"],
                    option["landing"],
                    option["offset"],
                    " | Stacked" if option["stacked"] else ""
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

        created_ids = CreateTags(doc, jobs, logs, errors)
        created = created_ids.Count

    # ======================================================
    # AUTOMATICALLY SELECT NEWLY CREATED TAGS
    # ======================================================

    selected_count = 0

    if WRITE and created_ids.Count > 0:

        try:
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
        "FFE RECEPTACLE TAGGING - v0.3",
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
        "Adjusted leader lengths: {}".format(adjusted_count),
        "Stacked tag candidates: {}".format(stacked_count),
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


# ____________________________________________________________________ ENTRY / REPORTING
def RunCommand():
    """Get settings and print results, including startup and fatal diagnostics."""
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            print("ERROR: Open a Revit project before running Tag Receptacles.")
            return
        doc = uidoc.Document
        if doc.IsFamilyDocument:
            print("ERROR: Open a project floor plan or sheet, not a family document.")
            return

        dialog = SettingsDialog()
        dialog.ShowDialog()
        settings = dialog.Settings
        if settings is None:
            return
        if settings["write"] and (doc.IsReadOnly or doc.IsModifiable):
            print("ERROR: The project is read-only or already has an open transaction. "
                  "Finish the current operation or choose Preview.")
            return

        script.get_output()
        for line in RunTagging(doc, uidoc, settings):
            print(line)
    except Exception:
        print("FATAL SCRIPT ERROR")
        print(traceback.format_exc())


if __name__ == "__main__":
    RunCommand()
