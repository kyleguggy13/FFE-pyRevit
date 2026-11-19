# -*- coding: utf-8 -*-
"""Purpose: persist complete project-team table definitions without storing Space data.
Behaviors: stable parameter references, layout restoration, and revision-checked writes.
Revit API: 2025/2026; storage and worksharing calls require an ExternalEvent context.
Design: pure layout functions and an injectable storage boundary support offline tests.
One DataStorage element per definition keeps unrelated tables independently editable.
"""
import json
import math
import time
import uuid

try:
    StringTypes = (basestring,)
except NameError:
    StringTypes = (str,)

SCHEMA_GUID = "7d601ebb-ff8a-4268-9ba7-145ea9a2f03c"
BASE_FIELDS = ("level", "number", "name", "areaValue", "heightValue", "phase")


def RequireText(value, label, maximum=1000):
    """Bound strings before accepting them from the browser or model storage."""
    if not isinstance(value, StringTypes) or len(value) > maximum:
        raise ValueError("Invalid {0}.".format(label))
    return value


def ValidateLayout(layout):
    """Validate portable or runtime layout structure without trusting model JSON."""
    if not isinstance(layout, dict):
        raise ValueError("Invalid table layout.")
    columns = layout.get("columns")
    if not isinstance(columns, list) or not columns or len(columns) > 500:
        raise ValueError("A table must have between 1 and 500 columns.")
    fields, normalized = set(), []
    for column in columns:
        if not isinstance(column, dict):
            raise ValueError("Invalid table column.")
        field = RequireText(column.get("field"), "column reference")
        width = column.get("width")
        if (not field or field in fields or isinstance(width, bool) or
                not isinstance(width, (int, float)) or math.isnan(width) or math.isinf(width) or not 60 <= width <= 3000):
            raise ValueError("Invalid or duplicate table column/width.")
        normalized.append({"field": field, "width": width})
        fields.add(field)
    sorters = layout.get("sorters", [])
    if not isinstance(sorters, list) or len(sorters) > len(columns):
        raise ValueError("Invalid table sorting.")
    normalizedSorters, sortedFields = [], set()
    for sorter in sorters:
        if (not isinstance(sorter, dict) or sorter.get("field") not in fields or
                sorter.get("field") in sortedFields or sorter.get("dir") not in ("asc", "desc")):
            raise ValueError("Invalid table sorting.")
        sortedFields.add(sorter["field"])
        normalizedSorters.append({"field": sorter["field"], "dir": sorter["dir"]})
    # Legacy definitions contain a single groupField; normalize without writing storage.
    legacyGroup = layout.get("groupField", "")
    groups = layout.get("groupFields", [legacyGroup] if legacyGroup else [])
    if (not isinstance(groups, list) or len(groups) > len(columns) or
            any(not isinstance(group, StringTypes) or group not in fields for group in groups) or
            len(set(groups)) != len(groups) or
            ("groupFields" not in layout and not isinstance(legacyGroup, StringTypes))):
        raise ValueError("Invalid table grouping.")
    headerGroups = layout.get("headerGroups", [])
    if not isinstance(headerGroups, list) or len(headerGroups) > len(columns):
        raise ValueError("Invalid column header groups.")
    normalizedHeaders, assigned = [], set()
    for header in headerGroups:
        if not isinstance(header, dict):
            raise ValueError("Invalid column header group.")
        name = RequireText(header.get("name"), "column group name", 100).strip()
        members = header.get("fields")
        if (not name or not isinstance(members, list) or not members or len(members) > len(columns) or
                any(not isinstance(field, StringTypes) or field not in fields or field in assigned for field in members) or
                len(set(members)) != len(members)):
            raise ValueError("Invalid column header group name or membership.")
        assigned.update(members)
        normalizedHeaders.append({"name": name, "fields": list(members)})
    filters = layout.get("filters")
    if not isinstance(filters, dict):
        raise ValueError("Invalid table filters.")
    return {"columns": normalized, "sorters": normalizedSorters, "groupFields": list(groups),
            "headerGroups": normalizedHeaders,
            "filters": {"phase": RequireText(filters.get("phase"), "phase filter"),
                        "level": RequireText(filters.get("level"), "level filter"),
                        "query": RequireText(filters.get("query", ""), "search filter", 5000)}}


def CaptureLayout(layout, snapshot):
    """Translate current grid fields and filter IDs into persistent model identities."""
    layout = ValidateLayout(layout)
    references = dict((field, "base:" + field) for field in BASE_FIELDS)
    for parameter in snapshot.get("parameters", []):
        if not parameter.get("baseField") and parameter.get("identity"):
            references["parameterValues." + parameter["key"]] = parameter["identity"]
    for column in layout["columns"]:
        if column["field"] not in references:
            raise ValueError("A column no longer resolves to a parameter. Refresh and retry.")
    portable = {"columns": [{"field": references[column["field"]], "width": column["width"]}
                             for column in layout["columns"]],
                "sorters": [{"field": references[sorter["field"]], "dir": sorter["dir"]}
                            for sorter in layout["sorters"]],
                "groupFields": [references[group] for group in layout["groupFields"]],
                "headerGroups": [{"name": header["name"], "fields": [references[field] for field in header["fields"]]}
                                 for header in layout["headerGroups"]],
                "filters": dict(layout["filters"])}
    for filterName, source in (("phase", "phases"), ("level", "levels")):
        value = layout["filters"][filterName]
        if value == "all":
            continue
        matches = [item for item in snapshot[source] if item["id"] == value and item.get("uniqueId")]
        if len(matches) != 1:
            raise ValueError("The {0} filter no longer resolves. Refresh and retry.".format(filterName))
        portable["filters"][filterName] = matches[0]["uniqueId"]
    return ValidateLayout(portable)


def ResolveLayout(layout, snapshot):
    """Restore usable settings by identity; never substitute a same-named parameter."""
    layout = ValidateLayout(layout)
    fields = dict(("base:" + field, field) for field in BASE_FIELDS)
    identities = {}
    for parameter in snapshot.get("parameters", []):
        identity = parameter.get("identity")
        if identity and not parameter.get("baseField"):
            identities.setdefault(identity, []).append("parameterValues." + parameter["key"])
    fields.update((identity, values[0]) for identity, values in identities.items() if len(values) == 1)
    warnings, columns = [], []
    for column in layout["columns"]:
        field = fields.get(column["field"])
        if field:
            columns.append({"field": field, "width": column["width"]})
        else:
            warnings.append("Unavailable column: {0}".format(column["field"]))
    present = set(column["field"] for column in columns)
    for field in BASE_FIELDS:
        if field not in present:
            columns.append({"field": field, "width": 140})
            present.add(field)
    sorters = [{"field": fields[sorter["field"]], "dir": sorter["dir"]}
               for sorter in layout["sorters"] if fields.get(sorter["field"]) in present]
    if len(sorters) != len(layout["sorters"]):
        warnings.append("Sorting on unavailable columns was cleared.")
    groups = [fields[group] for group in layout["groupFields"] if fields.get(group) in present]
    if len(groups) != len(layout["groupFields"]):
        warnings.append("Grouping on unavailable columns was removed; available hierarchy levels were retained.")
    filters = dict(layout["filters"])
    for filterName, source, fallback in (("phase", "phases", snapshot["defaultPhaseId"]),
                                         ("level", "levels", "all")):
        if filters[filterName] == "all":
            continue
        matches = [item for item in snapshot[source] if item.get("uniqueId") == filters[filterName]]
        if len(matches) == 1:
            filters[filterName] = matches[0]["id"]
        else:
            filters[filterName] = fallback
            warnings.append("The saved {0} filter is unavailable; the initial default was used.".format(filterName))
    if "phase" in groups and filters["phase"] != "all":
        groups.remove("phase")
        warnings.append("Phase grouping was cleared because the Phase column is hidden.")
    headers = []
    for header in layout["headerGroups"]:
        members = [fields[field] for field in header["fields"] if fields.get(field) in present]
        if len(members) != len(header["fields"]):
            warnings.append("Unavailable columns were removed from header group: {0}".format(header["name"]))
        if members:
            headers.append({"name": header["name"], "fields": members})
    return {"columns": columns, "sorters": sorters, "groupFields": groups,
            "headerGroups": headers, "filters": filters}, warnings


def ValidateRecord(record):
    """Reject unsupported definitions without replacing or deleting their storage."""
    if not isinstance(record, dict) or record.get("version") not in (1, 2, 3):
        raise ValueError("Unsupported saved-table version.")
    for key in ("id", "revision"):
        uuid.UUID(RequireText(record.get(key), key, 36))
    name = RequireText(record.get("name"), "table name", 100).strip()
    if not name:
        raise ValueError("Enter a table name.")
    return {"version": record["version"], "id": record["id"], "revision": record["revision"], "name": name,
            "modifiedBy": RequireText(record.get("modifiedBy", ""), "author"),
            "modifiedAt": RequireText(record.get("modifiedAt", ""), "timestamp"),
            "layout": ValidateLayout(record.get("layout"))}


def TableMetadata(record):
    """Expose the catalog without shipping every table's layout to the browser."""
    return dict((key, record[key]) for key in ("id", "revision", "name", "modifiedBy", "modifiedAt"))


class RevitTableStore:
    """Read and transact one model-backed table at a time with worksharing checks."""

    def __init__(self, document, db, sdk=None):
        if sdk is None:
            from System import Guid, String
            from System.Collections.Generic import List
            from Autodesk.Revit.DB.ExtensibleStorage import (AccessLevel, DataStorage, Entity,
                                                             Schema, SchemaBuilder)
            sdk = {"Guid": Guid, "String": String, "AccessLevel": AccessLevel,
                   "DataStorage": DataStorage, "Entity": Entity, "Schema": Schema,
                   "SchemaBuilder": SchemaBuilder, "CreateIds": lambda ids: List[db.ElementId](ids)}
        self.Document = document
        self.DB = db
        self.SDK = sdk

    def GetSchema(self, create=False):
        """Reading a project never creates a schema or a DataStorage element."""
        schema = self.SDK["Schema"].Lookup(self.SDK["Guid"](SCHEMA_GUID))
        if schema is None and create:
            builder = self.SDK["SchemaBuilder"](self.SDK["Guid"](SCHEMA_GUID))
            builder.SetSchemaName("FFERevitDataTableV1")
            builder.SetReadAccessLevel(self.SDK["AccessLevel"].Public)
            builder.SetWriteAccessLevel(self.SDK["AccessLevel"].Public)
            builder.AddSimpleField("DefinitionJson", self.SDK["String"])
            schema = builder.Finish()
        return schema

    def ReadElement(self, element, schema):
        """Decode only this application's entity and bound its payload size."""
        entity = element.GetEntity(schema)
        if not entity.IsValid():
            return None
        text = entity.Get[self.SDK["String"]](schema.GetField("DefinitionJson"))
        if len(text) > 1000000:
            raise ValueError("Saved table exceeds the supported size.")
        return ValidateRecord(json.loads(text))

    def List(self):
        """Report malformed/future records individually while loading usable tables."""
        schema = self.GetSchema()
        records, issues = [], []
        if schema is None:
            return records, issues
        for element in self.DB.FilteredElementCollector(self.Document).OfClass(self.SDK["DataStorage"]):
            try:
                record = self.ReadElement(element, schema)
                if record:
                    records.append((element, record))
            except Exception as error:
                issues.append("Saved table {0}: {1}".format(element.Id.Value, error))
        return records, issues

    def Read(self, tableId):
        """Require an unambiguous UUID, never a display name or browser element ID."""
        records, issues = self.List()
        matches = [(element, record) for element, record in records if record["id"] == tableId]
        if len(matches) != 1:
            raise ValueError("The saved table is missing or ambiguous. Refresh the table list and retry.")
        return matches[0]

    def ResolveParameterIds(self, layout):
        """Resolve selected definitions before collecting current values in one pass."""
        result = []
        for column in layout["columns"]:
            reference = column["field"]
            kind, separator, value = reference.partition(":")
            try:
                if kind == "builtin":
                    result.append(str(int(value)))
                elif kind == "shared":
                    element = self.DB.SharedParameterElement.Lookup(self.Document, self.SDK["Guid"](value))
                    if element is not None:
                        result.append(str(element.Id.Value))
                elif kind == "project":
                    element = self.Document.GetElement(value)
                    if element is not None:
                        result.append(str(element.Id.Value))
            except Exception:
                # ResolveLayout reports any reference absent from the new catalog.
                continue
        return result

    def CheckEditable(self, element):
        """Use an actual checkout, then inspect central updates and the local revision."""
        if self.Document.IsReadOnly:
            raise ValueError("This model is read-only. The working table has been retained.")
        if element is not None and self.Document.IsWorkshared:
            ids = self.DB.WorksharingUtils.CheckoutElements(self.Document, self.SDK["CreateIds"]([element.Id]))
            if element.Id.Value not in set(item.Value for item in ids):
                raise ValueError("Another user owns this saved table. Your changes have been retained; retry after they relinquish it.")
            state = self.DB.WorksharingUtils.GetModelUpdatesStatus(self.Document, element.Id)
            if str(state) not in ("CurrentWithCentral", "NotYetInCentral"):
                raise ValueError("This saved table changed in central. Reload Latest, reopen the saved table, or use Save As. Your changes have been retained.")

    def Write(self, action, tableId, revision, name=None, layout=None):
        """Apply one explicit metadata operation; commit failures always roll back."""
        element, previous = (None, None) if not tableId else self.Read(tableId)
        self.CheckEditable(element)
        if element is not None:
            # Checkout can contact central; re-read the entity before comparison.
            previous = self.ReadElement(element, self.GetSchema())
            if previous is None or previous["revision"] != revision:
                raise ValueError("This saved table has a newer revision. Reopen it or use Save As; your changes have been retained.")
        if action not in ("saveTable", "renameTable", "deleteTable") or (action != "saveTable" and previous is None):
            raise ValueError("Invalid saved-table operation.")
        record = None
        if action != "deleteTable":
            record = ValidateRecord({"version": 3, "id": tableId or str(uuid.uuid4()),
                                     "revision": str(uuid.uuid4()), "name": name,
                                     "modifiedBy": self.Document.Application.Username or "",
                                     "modifiedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                                     "layout": previous["layout"] if action == "renameTable" else layout})
            if len(json.dumps(record, ensure_ascii=True, allow_nan=False)) > 1000000:
                raise ValueError("Saved table exceeds the supported size.")
        transaction = self.DB.Transaction(self.Document, {
            "saveTable": "Save Revit Data Table", "renameTable": "Rename Revit Data Table",
            "deleteTable": "Delete Revit Data Table"}[action])
        try:
            if transaction.Start() != self.DB.TransactionStatus.Started:
                raise ValueError("Revit could not start the saved-table transaction.")
            if action == "deleteTable":
                self.Document.Delete(element.Id)
            else:
                schema = self.GetSchema(create=True)
                if element is None:
                    element = self.SDK["DataStorage"].Create(self.Document)
                entity = self.SDK["Entity"](schema)
                entity.Set[self.SDK["String"]](schema.GetField("DefinitionJson"),
                    json.dumps(record, ensure_ascii=True, allow_nan=False))
                element.SetEntity(entity)
            if transaction.Commit() != self.DB.TransactionStatus.Committed:
                raise ValueError("Revit did not commit the saved-table operation.")
        except Exception:
            if transaction.GetStatus() == self.DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        finally:
            transaction.Dispose()
        return record
