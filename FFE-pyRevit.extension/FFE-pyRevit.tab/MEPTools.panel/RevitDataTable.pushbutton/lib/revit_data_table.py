# -*- coding: utf-8 -*-
"""Purpose: collect read-only Space snapshots and validate table navigation.
Behaviors: preserve project formatting, numeric values, phases, and row identity.
Revit API: Revit 2025/2026, ElementId.Value; callers supply a valid API context.
Design: inject DB/ID collection boundaries so behavior can be tested without Revit.
No transactions, model writes, UI objects, or dispatcher calls belong here.
"""
import datetime
import math

try:
    StringTypes = (basestring,)
except NameError:
    StringTypes = (str,)


def IdText(elementId):
    """Keep Revit's 64-bit IDs out of JavaScript's floating point representation."""
    return str(elementId.Value)


def TryParameterText(element, parameterId, fallback):
    """Read the actual Space name rather than a combined element display name."""
    parameter = element.get_Parameter(parameterId)
    if parameter is not None:
        value = parameter.AsString()
        if value is not None:
            return value
    return fallback()


def TrySpacePhaseId(space, phaseNames, db):
    """Resolve the spatial phase, which can differ from Element.CreatedPhaseId."""
    for parameterId in (db.BuiltInParameter.ROOM_PHASE_ID, db.BuiltInParameter.ROOM_PHASE):
        try:
            parameter = space.get_Parameter(parameterId)
            if parameter is not None:
                phaseId = IdText(parameter.AsElementId())
                if phaseId in phaseNames:
                    return phaseId
        except Exception:
            # Either alias may be absent or unreadable; try the other before the
            # caller flags this row as unresolved. Never substitute the view phase.
            continue
    return None


def TryMeasurement(readValue, units, spec, db, label, warnings):
    """Distinguish unavailable values from zero; formatting failures remain visible."""
    try:
        value = float(readValue())
        if math.isnan(value) or math.isinf(value):
            raise ValueError("not a finite measurement")
    except Exception as error:
        warnings.append("{0} unavailable: {1}".format(label, error))
        return None, u"\u2014"
    try:
        return value, db.UnitFormatUtils.Format(units, spec, value, False)
    except Exception as error:
        warnings.append("{0} formatting unavailable: {1}".format(label, error))
        return value, u"\u2014"


def BuildSnapshot(document, documentKey, db):
    """Collect all host Spaces once; filtering operates on this immutable snapshot."""
    phases = [{"id": IdText(phase.Id), "name": phase.Name}
              for phase in document.Phases]
    phaseNames = dict((phase["id"], phase["name"]) for phase in phases)
    defaultPhaseId = phases[-1]["id"] if phases else "all"
    try:
        parameter = document.ActiveView.get_Parameter(db.BuiltInParameter.VIEW_PHASE)
        viewPhaseId = IdText(parameter.AsElementId()) if parameter else None
        if viewPhaseId in phaseNames:
            defaultPhaseId = viewPhaseId
    except Exception:
        # Schedules and other views can have no phase parameter.
        pass

    units = document.GetUnits()
    levels = {}
    rows = []
    issues = []
    skippedCount = 0
    spaces = (db.FilteredElementCollector(document)
              .OfCategory(db.BuiltInCategory.OST_MEPSpaces)
              .WhereElementIsNotElementType())
    for space in spaces:
        try:
            elementId = IdText(space.Id)
            warnings = []
            levelId = IdText(space.LevelId)
            if levelId not in levels:
                level = document.GetElement(space.LevelId)
                levels[levelId] = {
                    "id": levelId,
                    "name": level.Name if level else u"\u2014",
                    "elevation": float(level.Elevation) if level else None,
                }
            level = levels[levelId]
            if level["elevation"] is None:
                warnings.append("Level unavailable.")
            phaseId = TrySpacePhaseId(space, phaseNames, db)
            areaValue, areaText = TryMeasurement(
                lambda: space.Area, units, db.SpecTypeId.Area, db, "Area", warnings)
            heightValue, heightText = TryMeasurement(
                lambda: space.UnboundedHeight, units, db.SpecTypeId.Length, db,
                "Height", warnings)
            if space.Location is None:
                warnings.append("Unplaced Space; Show in Revit is unavailable.")
            elif areaValue == 0:
                warnings.append("Zero area; the Space may be unenclosed or redundant.")
            if phaseId not in phaseNames:
                warnings.append("Space phase is unavailable; view it under All phases.")
            rows.append({
                "id": elementId,
                "uniqueId": space.UniqueId,
                "levelId": levelId,
                "level": level["name"],
                "levelElevation": level["elevation"],
                "phaseId": phaseId,
                "phase": phaseNames.get(phaseId, u"\u2014"),
                "number": TryParameterText(space, db.BuiltInParameter.ROOM_NUMBER,
                                           lambda: space.Number),
                "name": TryParameterText(space, db.BuiltInParameter.ROOM_NAME,
                                         lambda: space.Name),
                "areaValue": areaValue,
                "areaText": areaText,
                "heightValue": heightValue,
                "heightText": heightText,
                "warnings": warnings,
            })
        except Exception as error:
            skippedCount += 1
            # One corrupt Space must not hide the remaining model data.
            try:
                identity = IdText(space.Id)
            except Exception:
                identity = "unknown"
            issues.append("Space {0}: {1}".format(identity, error))

    return {
        "documentKey": documentKey,
        "documentTitle": document.Title,
        "refreshedAt": datetime.datetime.now().isoformat(),
        "defaultPhaseId": defaultPhaseId,
        "phases": phases,
        "levels": list(levels.values()),
        "rows": rows,
        "skippedCount": skippedCount,
        "issues": issues,
    }


def ValidateRequest(message, documentKey, snapshot):
    """Only accept known actions and identities from this document's last snapshot."""
    if not isinstance(message, dict):
        raise ValueError("Expected a table request object.")
    action = message.get("action")
    if action == "ready":
        return {"action": "ready"}
    if action not in ("refresh", "select", "show"):
        raise ValueError("Unknown table action.")
    if message.get("documentKey") != documentKey:
        raise ValueError("This request belongs to another document. Reopen the table.")
    if action == "refresh":
        return {"action": action}
    elementIds = message.get("elementIds")
    if not isinstance(elementIds, list) or not elementIds:
        raise ValueError("Select at least one Space in the table.")
    knownIds = set(row["id"] for row in snapshot.get("rows", []))
    result = []
    seen = set()
    for elementId in elementIds:
        if not isinstance(elementId, StringTypes) or elementId not in knownIds:
            raise ValueError("A selected Space is not in this table. Refresh and retry.")
        if elementId not in seen:
            result.append(elementId)
            seen.add(elementId)
    return {"action": action, "elementIds": result}


def NavigateSpaces(uiDocument, snapshot, request, db, createIds):
    """Resolve stable identities again so deleted/replaced rows cannot select others."""
    rows = dict((row["id"], row) for row in snapshot["rows"])
    elementIds = []
    skipped = []
    for elementId in request["elementIds"]:
        row = rows[elementId]
        try:
            element = uiDocument.Document.GetElement(row["uniqueId"])
            if (element is None or IdText(element.Id) != elementId or
                    element.Category is None or
                    element.Category.Id.Value != int(db.BuiltInCategory.OST_MEPSpaces)):
                raise ValueError("Space was deleted or replaced; refresh the table")
            if request["action"] == "show":
                if element.Location is None or element.get_BoundingBox(None) is None:
                    raise ValueError("Space has no viewable geometry")
            elementIds.append(element.Id)
        except Exception as error:
            skipped.append("Space {0}: {1}".format(elementId, error))
    verb = "Selected" if request["action"] == "select" else "Shown"
    if elementIds:
        ids = createIds(elementIds)
        if request["action"] == "select":
            uiDocument.Selection.SetElementIds(ids)
        else:
            try:
                uiDocument.ShowElements(ids)
            except Exception as error:
                return {"kind": "warning", "message":
                        "Revit could not show the Spaces in an available view: {0}. "
                        "Open a suitable plan or 3D view and retry.".format(error)}
    message = "{0} {1} Space(s). Skipped {2}.".format(verb, len(elementIds), len(skipped))
    if skipped:
        message += " " + " | ".join(skipped[:5])
    return {"kind": "warning" if skipped else "success", "message": message}
