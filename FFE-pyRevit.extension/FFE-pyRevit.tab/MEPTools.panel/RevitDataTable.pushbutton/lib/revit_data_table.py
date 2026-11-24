# -*- coding: utf-8 -*-
"""Purpose: collect read-only Space snapshots and validate table navigation.
Behaviors: collect parameter definitions and selected values, preserving units and identity.
Revit API: Revit 2025/2026, ElementId.Value; callers supply a valid API context.
Design: inject DB/ID collection boundaries so behavior can be tested without Revit.
No transactions, model writes, UI objects, or dispatcher calls belong here.
"""
import datetime
import math
import uuid

try:
    StringTypes = (basestring,)
except NameError:
    StringTypes = (str,)

PublishActions = ("publishCatalog", "publishReview", "publishCommit")


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


def GetParameterType(parameter, document, db):
    """Classify origin separately from the definition's value/specification type."""
    if parameter.Id.Value < -1:
        return "Built-In"
    try:
        if parameter.IsShared:
            return "Shared Parameter"
    except Exception:
        pass
    if document is not None and parameter.Id.Value > 0:
        try:
            element = document.GetElement(parameter.Id)
            # Shared and global definitions derive from ParameterElement, so
            # check those specific origins before the general project definition.
            for typeName, label in (("SharedParameterElement", "Shared Parameter"),
                                    ("GlobalParameter", "Global Parameter"),
                                    ("ParameterElement", "Project Parameter")):
                elementType = getattr(db, typeName, None)
                if elementType is not None and isinstance(element, elementType):
                    return label
        except Exception:
            pass
    return "Other"


def GetParameterIdentity(parameter, document):
    """Persist built-in IDs, shared GUIDs, or definition UniqueIds across sessions."""
    if parameter.Id.Value < -1:
        return "builtin:" + IdText(parameter.Id)
    try:
        if parameter.IsShared:
            return "shared:" + str(parameter.GUID).lower()
    except Exception:
        pass
    if document is not None and parameter.Id.Value > 0:
        try:
            element = document.GetElement(parameter.Id)
            if element is not None:
                try:
                    return "shared:" + str(element.GuidValue).lower()
                except AttributeError:
                    return "project:" + element.UniqueId
        except Exception:
            pass
    return None


def BuildParameterDescriptor(parameter, db, document=None):
    """Identify parameters by ID so equal names never merge unrelated definitions."""
    parameterId = IdText(parameter.Id)
    storageType = str(parameter.StorageType)
    dataType = {"String": "Text", "Integer": "Integer", "Double": "Number",
                "ElementId": "Element reference", "None": "Unknown"}.get(storageType, storageType)
    typeId = ""
    try:
        spec = parameter.Definition.GetDataType()
        typeId = spec.TypeId
        if typeId:
            dataType = db.LabelUtils.GetLabelForSpec(spec) or dataType
    except Exception:
        # Some built-ins (including levels) have an empty or unlabeled spec.
        # Their storage type is still useful and must not hide the parameter.
        pass
    isBoolean = False
    try:
        isBoolean = bool(typeId) and typeId == db.SpecTypeId.Boolean.YesNo.TypeId
    except Exception:
        pass
    baseField = ""
    for builtInName, field in (("ROOM_LEVEL_ID", "level"), ("ROOM_NUMBER", "number"),
                               ("ROOM_NAME", "name"), ("ROOM_AREA", "areaValue"),
                               ("ROOM_HEIGHT", "heightValue"), ("ROOM_PHASE_ID", "phase"),
                               ("ROOM_PHASE", "phase")):
        try:
            if parameterId == str(int(getattr(db.BuiltInParameter, builtInName))):
                baseField = field
                break
        except (AttributeError, TypeError, ValueError):
            continue
    unitTypeId = ""
    if storageType == "Double":
        try:
            unitTypeId = parameter.GetUnitTypeId().TypeId
        except Exception:
            pass
    return {"id": parameterId, "key": "p" + parameterId,
            "name": parameter.Definition.Name, "dataType": dataType,
            "storageType": storageType, "dataTypeId": typeId, "isBoolean": isBoolean,
            "unitTypeId": unitTypeId,
            "baseField": baseField, "parameterType": GetParameterType(parameter, document, db),
            "identity": GetParameterIdentity(parameter, document)}


def ReadParameterValue(parameter, descriptor, document, elementNames):
    """Return a sortable raw value and Revit-formatted text for a selected column."""
    if not parameter.HasValue:
        return None, u"\u2014"
    storageType = descriptor["storageType"]
    if storageType == "String":
        value = parameter.AsString()
        return value, value if value is not None else u"\u2014"
    if storageType in ("Double", "Integer"):
        value = float(parameter.AsDouble()) if storageType == "Double" else int(parameter.AsInteger())
        if math.isnan(value) or math.isinf(value):
            raise ValueError("not a finite value")
        if descriptor["isBoolean"]:
            return value, "Yes" if value else "No"
        text = parameter.AsValueString()
        if text is None and storageType == "Double":
            raise ValueError("formatted value unavailable")
        return value, text if text is not None else str(value)
    if storageType == "ElementId":
        elementId = parameter.AsElementId()
        value = IdText(elementId)
        text = parameter.AsValueString()
        if text:
            return value, text
        if value not in elementNames:
            element = document.GetElement(elementId) if elementId.Value > 0 else None
            elementNames[value] = element.Name if element else (u"\u2014" if value == "-1" else value)
        return value, elementNames[value]
    return None, u"\u2014"


def ReadCapability(parameter, descriptor, document, references, value):
    """Record observed read/write capabilities and only portable element references."""
    result = {"writable": False}
    try:
        result["hasValue"] = bool(parameter.HasValue)
        result["writable"] = not parameter.IsReadOnly
        if descriptor["storageType"] == "ElementId" and value is not None and int(value) > 0:
            if value not in references:
                element = document.GetElement(parameter.AsElementId())
                references[value] = element.UniqueId if element else None
            if references[value]:
                result["referenceUniqueId"] = references[value]
    except Exception:
        # Preserve the observation; unresolved references stay unavailable.
        pass
    if result["writable"] and descriptor["identity"] and descriptor["identity"].startswith("shared:"):
        try:
            definition = document.GetElement(parameter.Id)
            result["canUnset"] = bool(definition and definition.ShouldHideWhenNoValue())
        except Exception:
            # Only shared parameters with this setting support ClearValue.
            pass
    return result


def CollectSpaceParameters(space, document, db, catalog, selectedIds, elementNames, warnings,
                           capabilities=None, references=None):
    """Union all instance definitions, reading values only for requested columns."""
    values, texts = {}, {}
    try:
        parameters = space.Parameters
    except Exception as error:
        warnings.append("Space parameters unavailable: {0}".format(error))
        return values, texts
    for parameter in parameters:
        try:
            parameterId = IdText(parameter.Id)
            if parameterId not in catalog:
                catalog[parameterId] = BuildParameterDescriptor(parameter, db, document)
            descriptor = catalog[parameterId]
            if parameterId in selectedIds:
                try:
                    value, text = ReadParameterValue(parameter, descriptor, document, elementNames)
                except Exception as error:
                    value, text = None, u"\u2014"
                    warnings.append("{0} unavailable: {1}".format(descriptor["name"], error))
                values[descriptor["key"]] = value
                texts[descriptor["key"]] = text
            if capabilities is not None:
                identity = descriptor["identity"]
                if descriptor["baseField"]:
                    identity = "base:" + descriptor["baseField"]
                if identity and (descriptor["baseField"] or parameterId in selectedIds):
                    capability = ReadCapability(parameter, descriptor, document, references,
                                                values.get(descriptor["key"]))
                    if descriptor["baseField"] in ("level", "phase", "areaValue", "heightValue"):
                        # Missing measurements or references are unavailable regardless of alias state.
                        capability.pop("hasValue", None)
                    capabilities[identity] = capability
        except Exception as error:
            warnings.append("Parameter definition unavailable: {0}".format(error))
    return values, texts


def BuildSnapshot(document, documentKey, db, selectedParameterIds=None, capabilities=None):
    """Collect all host Spaces once; filtering operates on this immutable snapshot."""
    phases = [{"id": IdText(phase.Id), "name": phase.Name, "uniqueId": getattr(phase, "UniqueId", None)}
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
    selectedIds = set(selectedParameterIds or [])
    catalog = {}
    elementNames = {}
    references = {}
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
                    "uniqueId": getattr(level, "UniqueId", None) if level else None,
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
            uniqueId = space.UniqueId
            spaceCapabilities = None if capabilities is None else capabilities.setdefault(uniqueId, {})
            parameterValues, parameterTexts = CollectSpaceParameters(
                space, document, db, catalog, selectedIds, elementNames, warnings, spaceCapabilities, references)
            rows.append({
                "id": elementId,
                "uniqueId": uniqueId,
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
                "parameterValues": parameterValues,
                "parameterTexts": parameterTexts,
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
        "parameters": sorted(catalog.values(), key=lambda item: (item["name"].lower(), item["id"])),
        "selectedParameterIds": [parameterId for parameterId in (selectedParameterIds or [])
                                 if parameterId in catalog],
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
    tableActions = ("loadTable", "saveTable", "renameTable", "deleteTable")
    if action not in ("refresh", "select", "show", "closeWindow") + tableActions + PublishActions:
        raise ValueError("Unknown table action.")
    if message.get("documentKey") != documentKey:
        raise ValueError("This request belongs to another document. Reopen the table.")
    if action == "closeWindow":
        return {"action": action}
    if action in PublishActions:
        requestId = message.get("requestId")
        if not isinstance(requestId, StringTypes) or not requestId or len(requestId) > 100:
            raise ValueError("Invalid publish request.")
        request = {"action": action, "requestId": requestId}
        if action == "publishReview":
            projectId = message.get("projectId")
            if not isinstance(projectId, StringTypes) or len(projectId) != 36:
                raise ValueError("Choose a cloud project before collecting changes.")
            request["projectId"] = str(uuid.UUID(projectId))
        elif action == "publishCommit":
            reviewId = message.get("reviewId")
            if not isinstance(reviewId, StringTypes) or len(reviewId) != 36:
                raise ValueError("Collect Spaces before syncing.")
            request["reviewId"] = str(uuid.UUID(reviewId))
        return request
    if action in tableActions:
        tableId = message.get("tableId", "")
        revision = message.get("revision", "")
        requestId = message.get("requestId")
        if (not isinstance(tableId, StringTypes) or len(tableId) > 36 or
                not isinstance(revision, StringTypes) or len(revision) > 36 or
                not isinstance(requestId, StringTypes) or not requestId or len(requestId) > 100):
            raise ValueError("Invalid saved-table request.")
        if action in ("renameTable", "deleteTable") and not tableId:
            raise ValueError("Choose a saved table first.")
        name = message.get("name", "")
        if action in ("saveTable", "renameTable"):
            if not isinstance(name, StringTypes) or not name.strip() or len(name) > 100:
                raise ValueError("Enter a table name of 1 to 100 characters.")
        if action == "saveTable" and not isinstance(message.get("layout"), dict):
            raise ValueError("Invalid table layout.")
        return {"action": action, "tableId": tableId, "revision": revision, "requestId": requestId,
                "name": name.strip(), "layout": message.get("layout")}
    if action == "refresh":
        parameterIds = message.get("parameterIds", snapshot.get("selectedParameterIds", []))
        if not isinstance(parameterIds, list):
            raise ValueError("Expected a list of Space parameter IDs.")
        knownParameters = set(parameter["id"] for parameter in snapshot.get("parameters", []))
        selected, seen = [], set()
        for parameterId in parameterIds:
            if not isinstance(parameterId, StringTypes) or parameterId not in knownParameters:
                raise ValueError("A parameter is no longer in this table. Refresh and reopen Parameters.")
            if parameterId not in seen:
                selected.append(parameterId)
                seen.add(parameterId)
        return {"action": action, "parameterIds": selected}
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
