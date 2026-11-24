# -*- coding: utf-8 -*-
"""Purpose: sync full Revit Space parameter snapshots with Loadwise API v3.
Behaviors: local preview, explicit sync, repeat-safe requests and project binding.
Revit API: 2025/2026; storage runs only in an ExternalEvent API context.
Design: HTTP accepts plain dictionaries on a worker; preview never writes cloud
or model data. The retired v2 model UUID is never treated as a project UUID.
"""
import uuid

from revit_spaces import SpacesSnapshot, _catalog
from saved_tables import RevitTableStore

CLOUD_PROJECT_SCHEMA_GUID = "711c7777-907c-4c60-9ea7-35e7567465de"


def LoadCatalog(client, linkedProjectId=None, preferredProjectId=None, formaContext=None):
    """List v3 projects without the retired model registry or cursor protocol."""
    if formaContext is not None:
        if not formaContext.get("available"):
            raise ValueError("The Autodesk Forma project could not be read. Reopen HVAC Data before syncing.")
        found = client.find_forma_project({"formaProjectId": formaContext["formaProjectId"]})["project"]
        if found is None:
            raise ValueError("Start this Autodesk Forma project in Loadwise, then reopen HVAC Data.")
        actualId = found["projectId"]
        if preferredProjectId and preferredProjectId != actualId:
            raise ValueError("The Autodesk Forma project connection changed. Reopen HVAC Data.")
        preferredProjectId = actualId
    if linkedProjectId and preferredProjectId and linkedProjectId != preferredProjectId:
        raise ValueError("This model is connected to a different Loadwise project than Autodesk Forma.")
    projects = [{"id": item["projectId"], "name": item["name"]}
                for item in client.list_projects()["items"]]
    projectId = preferredProjectId or linkedProjectId
    if projectId and not any(item["id"] == projectId for item in projects):
        raise ValueError("The connected Loadwise project is unavailable. Check the project in Loadwise.")
    return {"projects": projects, "projectId": projectId}


def SelectSyncRows(snapshot, groupIds=None):
    """Resolve the host-owned scope before extracting every selected Space parameter."""
    rows = snapshot["rows"]
    warnings = list(snapshot["issues"])
    scope = {"allSpaces": groupIds is None, "groups": [], "excludedCount": 0, "unavailableCount": 0}
    if groupIds is not None:
        if not groupIds:
            raise ValueError("Choose at least one Space group to sync.")
        groups = snapshot.get("spaceGroups", [])
        chosen = []
        for groupId in sorted(set(groupIds)):
            matches = [group for group in groups if group["id"] == groupId]
            if len(matches) != 1:
                raise ValueError("A selected Space group is missing or ambiguous. Refresh and select groups again.")
            chosen.append(matches[0])
        members = set(member for group in chosen for member in group["members"])
        available = set(row["uniqueId"] for row in rows)
        scope["groups"] = [{"id": group["id"], "name": group["name"]} for group in chosen]
        scope["unavailableCount"] = len(members - available)
        selectedRows = [row for row in rows if row["uniqueId"] in members]
        scope["excludedCount"] = len(rows) - len(selectedRows)
        rows = selectedRows
        if not rows:
            raise ValueError("The selected Space groups contain no available Spaces. Refresh and edit their membership.")
        if scope["unavailableCount"]:
            warnings.append("{0} group member(s) were not collected; they may have been deleted or unreadable."
                            .format(scope["unavailableCount"]))
    return rows, scope, warnings


def PrepareSync(snapshot, projectId, groupIds=None, parameterData=None):
    """Freeze a full-parameter preview and deterministic requests for safe retries."""
    rows, scope, warnings = SelectSyncRows(snapshot, groupIds)
    data = parameterData or {"spaces": {row["uniqueId"]: row.get("parameters", {}) for row in rows},
                             "catalog": snapshot.get("syncParameterCatalog", []), "issues": [],
                             "skippedCount": 0, "unavailableCount": 0}
    warnings.extend(data["issues"])
    enriched = [dict(row, parameters=data["spaces"][row["uniqueId"]])
                for row in rows if row["uniqueId"] in data["spaces"]]
    spaces = SpacesSnapshot(enriched)
    if not spaces:
        raise ValueError("No Spaces with readable parameter collections are available to sync.")
    catalog = _catalog(data["catalog"])
    if len(catalog) > 5000:
        raise ValueError("Loadwise supports at most 5000 parameter definitions per sync.")
    syncId = str(uuid.uuid4())
    projectId = str(uuid.UUID(projectId))
    skipped = snapshot["skippedCount"] + data["skippedCount"]
    # A partial/group scan must never mark other groups or unreadable Spaces missing.
    complete = scope["allSpaces"] and not skipped
    chunks = []
    for start in range(0, len(spaces), 200):
        chunk = {"projectId": projectId, "syncId": syncId, "spaces": spaces[start:start + 200],
                 "final": bool(complete and start + 200 >= len(spaces))}
        if start == 0:
            chunk["parameters"] = catalog
        chunks.append(chunk)
    return {"reviewId": str(uuid.uuid4()), "request": {"projectId": projectId, "spaces": spaces},
            "chunks": chunks, "nextChunk": 0,
            "totals": {"created": 0, "updated": 0, "unchanged": 0, "newlyMissing": []},
            "parameterCount": sum(len(space["parameters"]) for space in spaces),
            "dimensions": {row["uniqueId"]: {"area": row.get("areaValue"), "ceilingHeight": row.get("heightValue")} for row in enriched},
            "parameterDefinitionCount": len(catalog), "unavailableParameterCount": data["unavailableCount"],
            "skippedCount": skipped, "warnings": warnings, "scope": scope, "complete": bool(complete)}


def SyncReview(client, review):
    """Resume frozen chunks after an uncertain response without changing their sync ID."""
    while review["nextChunk"] < len(review["chunks"]):
        chunk = review["chunks"][review["nextChunk"]]
        result = client.push_revit_spaces(chunk)
        for key in ("created", "updated", "unchanged"):
            review["totals"][key] += result[key]
        review["totals"]["newlyMissing"] = result.get("newlyMissing", [])
        review["nextChunk"] += 1
    return dict(review["totals"], chunks=len(review["chunks"]), parameterCount=review["parameterCount"],
                complete=review["complete"])


class CloudProjectStore(RevitTableStore):
    """Keep the Loadwise project UUID on Project Information independently of bridge keys."""

    def GetSchema(self, create=False):
        """Reading a publish destination never creates model metadata."""
        schema = self.SDK["Schema"].Lookup(self.SDK["Guid"](CLOUD_PROJECT_SCHEMA_GUID))
        if schema is None and create:
            builder = self.SDK["SchemaBuilder"](self.SDK["Guid"](CLOUD_PROJECT_SCHEMA_GUID))
            builder.SetSchemaName("FFEHVACLoadwiseProjectV1")
            builder.SetReadAccessLevel(self.SDK["AccessLevel"].Public)
            builder.SetWriteAccessLevel(self.SDK["AccessLevel"].Public)
            builder.AddSimpleField("ProjectId", self.SDK["String"])
            schema = builder.Finish()
        return schema

    def Read(self):
        """Return an existing project binding, including on a read-only linked model."""
        schema = self.GetSchema()
        if schema is None:
            return None
        entity = self.Document.ProjectInformation.GetEntity(schema)
        if not entity.IsValid():
            return None
        return str(uuid.UUID(str(entity.Get[self.SDK["String"]](schema.GetField("ProjectId")))))

    def Ensure(self, projectId):
        """Persist the chosen identity once; reject checkout races and roll back failures."""
        projectId = str(uuid.UUID(projectId))
        existing = self.Read()
        if existing == projectId:
            return False
        if existing is not None:
            raise ValueError("This model's Loadwise project connection changed. Reopen Sync Spaces before retrying.")
        try:
            self.CheckEditable(self.Document.ProjectInformation)
        except ValueError as error:
            raise ValueError(str(error).replace("saved table", "Loadwise project connection"))
        # Checkout may refresh model metadata; never overwrite another user's binding.
        existing = self.Read()
        if existing is not None:
            if existing == projectId:
                return False
            raise ValueError("This model's Loadwise project connection changed. Reopen Sync Spaces before retrying.")
        transaction = self.DB.Transaction(self.Document, "Connect HVAC Data to Loadwise project")
        try:
            if transaction.Start() != self.DB.TransactionStatus.Started:
                raise ValueError("Revit could not start the Loadwise project connection transaction.")
            schema = self.GetSchema(create=True)
            entity = self.SDK["Entity"](schema)
            entity.Set[self.SDK["String"]](schema.GetField("ProjectId"), projectId)
            self.Document.ProjectInformation.SetEntity(entity)
            if transaction.Commit() != self.DB.TransactionStatus.Committed:
                raise ValueError("Revit did not commit the Loadwise project connection.")
        except Exception:
            if transaction.GetStatus() == self.DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        finally:
            transaction.Dispose()
        return True
