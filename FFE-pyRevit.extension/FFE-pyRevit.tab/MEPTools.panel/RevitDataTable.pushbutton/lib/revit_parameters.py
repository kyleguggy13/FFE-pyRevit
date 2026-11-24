# -*- coding: utf-8 -*-
"""Purpose: capture every Space instance/type parameter for Loadwise coordination.
Behaviors: display-unit values, explicit empty/unavailable values and unique names.
Revit API: 2025/2026; call only in an API context, returning JSON data only.
Design: discover definitions before reading selected Spaces so duplicate names get
stable keys across group syncs. Table columns never control parameter extraction.
"""
import hashlib
import math

from revit_data_table import BuildParameterDescriptor, IdText, StringTypes


def ParameterNames(definitions):
    """Keep ordinary names usable for coordination and disambiguate collisions."""
    names, used = {}, set()
    # Prefer a built-in instance definition for the original name. Definitions
    # from other scopes retain a deterministic identity suffix instead of merging.
    ordered = sorted(definitions.items(), key=lambda pair: (
        pair[1]["isType"], not pair[1]["identity"].startswith("builtin:"), pair[0]))
    for identity, definition in ordered:
        name = definition["name"].strip() or "Unnamed parameter"
        if name in used or len(name) > 200:
            suffix = " [{0} {1}]".format("type" if definition["isType"] else "instance",
                                        hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16])
            name = name[:200 - len(suffix)] + suffix
        if name in used:
            raise ValueError("Parameter names could not be distinguished safely.")
        names[identity] = name
        used.add(name)
    return names


def ParameterUnit(parameter, descriptor, document, db):
    """Resolve the project's display unit before converting Revit's internal value."""
    if descriptor["storageType"] != "Double":
        return None, ""
    if not descriptor["dataTypeId"]:
        return None, "Revit internal units"
    spec = parameter.Definition.GetDataType()
    if hasattr(db.UnitUtils, "IsMeasurableSpec") and not db.UnitUtils.IsMeasurableSpec(spec):
        return None, ""
    options = document.GetUnits().GetFormatOptions(spec)
    unitId = options.GetUnitTypeId()
    try:
        label = db.LabelUtils.GetLabelForSymbol(options.GetSymbolTypeId())
    except Exception:
        label = ""
    if not label:
        label = db.LabelUtils.GetLabelForUnit(unitId)
    return unitId, label[:50]


def ReadSyncValue(parameter, descriptor, unitId, document, db):
    """Keep large references as strings and Yes/No values as booleans."""
    if not parameter.HasValue:
        return {"v": None, "s": ""}
    storage = descriptor["storageType"]
    if storage == "String":
        value = parameter.AsString()
        return {"v": value, "s": value or ""}
    if storage == "Double":
        value = float(parameter.AsDouble())
        if unitId is not None:
            value = db.UnitUtils.ConvertFromInternalUnits(value, unitId)
        if math.isnan(value) or math.isinf(value):
            raise ValueError("not a finite value")
    elif storage == "Integer":
        value = int(parameter.AsInteger())
        if descriptor["isBoolean"]:
            return {"v": bool(value), "s": "Yes" if value else "No"}
    elif storage == "ElementId":
        reference = parameter.AsElementId()
        value = IdText(reference)
        try:
            target = document.GetElement(reference) if reference.Value > 0 else None
            display = target.Name if target is not None else value
        except Exception:
            display = value
    else:
        raise ValueError("unsupported storage type {0}".format(storage))
    try:
        text = parameter.AsValueString()
    except Exception:
        text = None
    return {"v": value, "s": text if text is not None else (display if storage == "ElementId" else str(value))}


def CollectSyncParameters(document, rows, db, allRows=None):
    """Capture all values for the scoped Spaces, including parameters on their types."""
    definitions, entries, issues = {}, {}, []
    chosen = set(row["uniqueId"] for row in rows)
    failed, typeCache, unitCache = set(), {}, {}
    for row in (allRows if allRows is not None else rows):
        uniqueId = row["uniqueId"]
        selected = uniqueId in chosen
        try:
            space = document.GetElement(uniqueId)
            if space is None:
                raise ValueError("Space no longer exists")
            instanceParameters = list(space.Parameters)
        except Exception as error:
            if selected:
                failed.add(uniqueId)
                issues.append("Space {0}: parameters unavailable ({1}); Space skipped.".format(row["id"], error))
            continue
        owners = [(False, instanceParameters)]
        try:
            typeId = space.GetTypeId()
            if typeId.Value > 0:
                key = IdText(typeId)
                if key not in typeCache:
                    elementType = document.GetElement(typeId)
                    typeCache[key] = list(elementType.Parameters) if elementType is not None else []
                owners.append((True, typeCache[key]))
        except Exception as error:
            if selected:
                issues.append("Space {0}: type parameters unavailable ({1}).".format(row["id"], error))
        found = []
        for isType, parameters in owners:
            for parameter in parameters:
                try:
                    descriptor = BuildParameterDescriptor(parameter, db, document)
                    descriptor["name"] = descriptor["name"] or "Unnamed parameter"
                    stable = descriptor["identity"] or "id:" + descriptor["id"]
                    identity = ("type:" if isType else "instance:") + stable
                    if identity not in definitions:
                        definitions[identity] = dict(descriptor, isType=isType, identity=stable)
                    if selected:
                        found.append((identity, parameter))
                except Exception as error:
                    if selected:
                        issues.append("Space {0}: parameter definition unavailable ({1}).".format(row["id"], error))
        if selected:
            entries[uniqueId] = found
    names = ParameterNames(definitions)
    spaces, catalog, unavailable = {}, {}, 0
    for row in rows:
        uniqueId = row["uniqueId"]
        if uniqueId in failed:
            continue
        values = {}
        for identity, parameter in entries.get(uniqueId, []):
            definition = definitions[identity]
            name = names[identity]
            unitId, unit = None, ""
            try:
                if identity not in unitCache:
                    unitCache[identity] = ParameterUnit(parameter, definition, document, db)
                unitId, unit = unitCache[identity]
                value = ReadSyncValue(parameter, definition, unitId, document, db)
            except Exception as error:
                unavailable += 1
                value = {"v": None, "s": "Unavailable"}
                issues.append("Space {0}, {1}: unavailable ({2}).".format(row["id"], name, error))
            for field in ("v", "s"):
                if isinstance(value.get(field), StringTypes) and len(value[field]) > 2000:
                    value[field] = value[field][:2000]
                    issues.append("Space {0}, {1}: {2} truncated to 2000 characters.".format(row["id"], name, field))
            values[name] = value
            if name not in catalog:
                try:
                    group = db.LabelUtils.GetLabelForGroup(parameter.Definition.GetGroupTypeId())
                except Exception:
                    group = ""
                catalog[name] = {"name": name, "storageType": definition["storageType"], "unit": unit,
                                 "group": group[:100], "isShared": definition["parameterType"] == "Shared Parameter",
                                 "isType": definition["isType"]}
        spaces[uniqueId] = values
    return {"spaces": spaces, "catalog": sorted(catalog.values(), key=lambda item: item["name"]),
            "issues": issues, "skippedCount": len(failed), "unavailableCount": unavailable}
