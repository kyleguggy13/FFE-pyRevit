# -*- coding: utf-8 -*-
"""Purpose: collect and sync Revit-owned room fields with Loadwise API v3.
Behaviors: local preview, explicit sync, repeat-safe requests and project binding.
Revit API: 2025/2026; storage runs only in an ExternalEvent API context.
Design: HTTP accepts plain dictionaries on a worker; preview never writes cloud
or model data. The retired v2 model UUID is never treated as a project UUID.
"""
import uuid

from revit_spaces import SpacesPayload
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


def PrepareSync(snapshot, projectId):
    """Freeze all collected Spaces for preview and an explicitly requested sync."""
    payload = SpacesPayload(snapshot["rows"])
    if len(payload["spaces"]) > 5000:
        raise ValueError("Loadwise supports at most 5000 Spaces per sync.")
    payload["projectId"] = str(uuid.UUID(projectId))
    return {"reviewId": str(uuid.uuid4()), "request": payload,
            "skippedCount": snapshot["skippedCount"], "warnings": list(snapshot["issues"])}


def SyncReview(client, review):
    """Send only Revit identity and geometry; repeated requests never delete rooms."""
    return client.sync_revit_spaces(review["request"])


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
