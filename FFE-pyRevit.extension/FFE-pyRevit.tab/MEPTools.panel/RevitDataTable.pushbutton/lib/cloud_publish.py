# -*- coding: utf-8 -*-
"""Purpose: publish complete HVAC snapshots using FFE-Mechanical's shared RPCs.
Behaviors: persistent model identity, explicit destinations, revision conflict checks.
Revit API: 2025/2026; CloudModelStore runs only in an ExternalEvent API context.
Design: RPC helpers accept JSON only and run on a worker; identity writes use one
short transaction before HTTP starts. No transaction remains open during network IO.
"""
import json
import uuid

from saved_tables import RevitTableStore

CLOUD_SCHEMA_GUID = "a9b98292-45e8-4539-a20c-abd263581dfb"


def PrepareSnapshot(snapshot, modelId):
    """Detach plain JSON and refuse incomplete extraction before replacing cloud data."""
    if snapshot.get("skippedCount", 0):
        raise ValueError("Some Spaces could not be collected. Resolve the skipped-Space warnings and retry Publish.")
    result = json.loads(json.dumps(snapshot, ensure_ascii=True, allow_nan=False))
    for field in ("savedTables", "tableIssues", "canSaveTables"):
        result.pop(field, None)
    result["documentKey"] = str(uuid.UUID(modelId))
    return result


def LoadCatalog(client, modelId):
    """Load projects and capture the revision the user will review before publishing."""
    projects = client.list_projects()
    if not isinstance(projects, list):
        raise ValueError("Supabase returned an invalid project list.")
    choices = []
    for project in projects:
        choices.append({"id": str(uuid.UUID(project["id"])), "name": project["name"]})
    current = client.get_snapshot(modelId)
    model = None
    if current is not None:
        model = {key: current[key] for key in ("id", "projectId", "name", "revision")}
        if model["id"] != modelId or not any(project["id"] == model["projectId"] for project in choices):
            raise ValueError("This model's cloud destination is unavailable. Reload projects and retry.")
        uuid.UUID(model["revision"])
    return {"modelId": modelId, "projects": choices, "model": model}


def PublishSnapshot(client, plan, projectId, snapshot):
    """Create only this model, then compare the reviewed revision without automatic retries."""
    if not any(project["id"] == projectId for project in plan["projects"]):
        raise ValueError("Choose an available cloud project.")
    payload = PrepareSnapshot(snapshot, plan["modelId"])
    model = plan["model"]
    if model is not None and model["projectId"] != projectId:
        raise ValueError("This model is already published in another project. Reload its destination.")
    if model is None:
        model = client.create_model(projectId, plan["modelId"], snapshot["documentTitle"][:200])
    result = client.publish_snapshot(plan["modelId"], model["revision"], payload)
    return {"modelId": plan["modelId"], "revision": result["revision"],
            "rowCount": len(payload["rows"]), "parameterCount": len(payload["selectedParameterIds"])}


class CloudModelStore(RevitTableStore):
    """Keep the cloud model UUID on Project Information independently of bridge keys."""

    def GetSchema(self, create=False):
        """Reading a publish destination never creates model metadata."""
        schema = self.SDK["Schema"].Lookup(self.SDK["Guid"](CLOUD_SCHEMA_GUID))
        if schema is None and create:
            builder = self.SDK["SchemaBuilder"](self.SDK["Guid"](CLOUD_SCHEMA_GUID))
            builder.SetSchemaName("FFEHVACCloudModelV1")
            builder.SetReadAccessLevel(self.SDK["AccessLevel"].Public)
            builder.SetWriteAccessLevel(self.SDK["AccessLevel"].Public)
            builder.AddSimpleField("ModelId", self.SDK["String"])
            schema = builder.Finish()
        return schema

    def Read(self):
        """Return an existing stable identity, including on a read-only linked model."""
        schema = self.GetSchema()
        if schema is None:
            return None
        entity = self.Document.ProjectInformation.GetEntity(schema)
        if not entity.IsValid():
            return None
        return str(uuid.UUID(str(entity.Get[self.SDK["String"]](schema.GetField("ModelId")))))

    def Ensure(self, modelId):
        """Persist the chosen identity once; reject checkout races and roll back failures."""
        modelId = str(uuid.UUID(modelId))
        existing = self.Read()
        if existing == modelId:
            return False
        if existing is not None:
            raise ValueError("This model's cloud identity changed. Reopen Publish before retrying.")
        try:
            self.CheckEditable(self.Document.ProjectInformation)
        except ValueError as error:
            raise ValueError(str(error).replace("saved table", "cloud model identity"))
        # Checkout may refresh model metadata; never overwrite another user's binding.
        existing = self.Read()
        if existing is not None:
            if existing == modelId:
                return False
            raise ValueError("This model's cloud identity changed. Reopen Publish before retrying.")
        transaction = self.DB.Transaction(self.Document, "Link HVAC Data to cloud model")
        try:
            if transaction.Start() != self.DB.TransactionStatus.Started:
                raise ValueError("Revit could not start the cloud model link transaction.")
            schema = self.GetSchema(create=True)
            entity = self.SDK["Entity"](schema)
            entity.Set[self.SDK["String"]](schema.GetField("ModelId"), modelId)
            self.Document.ProjectInformation.SetEntity(entity)
            if transaction.Commit() != self.DB.TransactionStatus.Committed:
                raise ValueError("Revit did not commit the cloud model link.")
        except Exception:
            if transaction.GetStatus() == self.DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        finally:
            transaction.Dispose()
        return True
