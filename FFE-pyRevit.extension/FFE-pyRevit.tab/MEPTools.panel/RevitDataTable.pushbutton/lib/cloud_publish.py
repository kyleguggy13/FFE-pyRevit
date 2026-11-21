# -*- coding: utf-8 -*-
"""Purpose: review and commit Space changes using FFE-Mechanical's repository API.
Behaviors: persistent model identity, explicit staging, optimistic concurrency.
Revit API: 2025/2026; CloudModelStore runs only in an ExternalEvent API context.
Design: RPC helpers accept JSON only and run on a worker; identity writes use one
short transaction before HTTP starts. No transaction remains open during network IO.
"""
import uuid

from repository_collection import from_snapshot
from saved_tables import RevitTableStore

CLOUD_SCHEMA_GUID = "a9b98292-45e8-4539-a20c-abd263581dfb"


def PrepareCollection(snapshot, capabilities):
    """Send supported observations and preserve coverage and extraction warnings."""
    supported = []
    issues = list(snapshot["issues"])
    loaded = set(snapshot["selectedParameterIds"])
    for parameter in snapshot["parameters"]:
        if parameter["baseField"] or parameter["id"] not in loaded:
            continue
        storage = parameter["storageType"]
        if (not parameter["identity"] or storage not in ("String", "Double", "Integer", "ElementId")
                or (storage == "Double" and not parameter["dataTypeId"])):
            issues.append("Cannot publish {0} ({1}): its type or stable identity is unavailable."
                          .format(parameter["name"], parameter["id"]))
        else:
            supported.append(parameter)
    return from_snapshot(dict(snapshot, parameters=supported, issues=issues), capabilities,
                         {"spaces": "full", "successful": not snapshot["skippedCount"]})


def ReadPages(operation, request):
    """Read the repository's cursor protocol with the original filters."""
    page = operation(request)
    for item in page["items"]:
        yield item
    while page["nextCursor"] is not None:
        page = operation(dict(request, cursor=page["nextCursor"]))
        for item in page["items"]:
            yield item


def LoadCatalog(client, modelId, linked, preferredProjectId=None, cloudContext=None):
    """Find the stored model's project, including links created by the old command."""
    choices = [{"id": project["projectId"], "name": project["name"], "headCommitId": project["headCommitId"]}
               for project in ReadPages(client.list_projects, {"limit": 200})]
    model = None
    if linked or preferredProjectId:
        for project in choices:
            if not linked and project["id"] != preferredProjectId:
                continue
            for candidate in ReadPages(client.list_models, {"projectId": project["id"], "limit": 200}):
                if candidate["modelId"] == modelId:
                    model = {"id": modelId, "projectId": project["id"], "name": candidate["name"]}
                    break
            if model is not None:
                break
    if model is not None and preferredProjectId and model["projectId"] != preferredProjectId:
        raise ValueError("This model is linked to a different project than Autodesk Forma. Review the existing connection.")
    return {"modelId": modelId, "projects": choices, "model": model,
            "projectId": preferredProjectId, "cloudContext": cloudContext}


def ReviewCollection(client, plan, project, collection, modelName, spaceLabels):
    """Register this model if needed and compare observations at the reviewed head."""
    projectId, head = project["id"], project["headCommitId"]
    modelId = plan["modelId"]
    if plan.get("projectId") and projectId != plan["projectId"]:
        raise ValueError("Use this model's Autodesk Forma project.")
    if plan["model"] is None:
        registration = {"projectId": projectId, "modelId": modelId, "name": modelName}
        context = plan.get("cloudContext")
        if context and context.get("available"):
            registration["metadata"] = {"formaProjectId": context["formaProjectId"], "cloudModelGuid": context.get("cloudModelGuid")}
        try:
            client.register_model(registration)
        except Exception as error:
            if getattr(error, "code", None) != "23505":
                raise
            # Another cloud copy may have registered this deterministic model ID.
            if not any(item["modelId"] == modelId for item in ReadPages(client.list_models, {"projectId": projectId, "limit": 200})):
                raise
    collectionId = str(uuid.uuid4())
    client.submit_collection({"projectId": projectId, "modelId": modelId, "collectionId": collectionId,
                              "baseCommitId": head, "collection": collection})
    candidates = {}
    incomingCount = 0
    for item in ReadPages(client.compare_collection, {"projectId": projectId, "collectionId": collectionId,
                                                       "commitId": head, "limit": 200}):
        classification = item["classification"]
        if classification == "incoming":
            incomingCount += 1
        elif classification in ("local_change", "conflict", "unbased"):
            item = dict(item, spaceLabel=spaceLabels.get(item["spaceId"], item["spaceId"]))
            candidates[str(len(candidates))] = item
    return {"collectionId": collectionId, "projectId": projectId, "headCommitId": head,
            "candidates": candidates, "incomingCount": incomingCount, "skippedCount": collection["skippedCount"],
            "warningCount": len(collection["warnings"])}


def CreateBatch(review, selection, message):
    """Accept only reviewed proposal IDs and require an explicit Space creation selection."""
    candidates = review["candidates"]
    if not selection or any(key not in candidates for key in selection):
        raise ValueError("Choose changes from the current review.")
    additions = {item["spaceId"]: key for key, item in candidates.items()
                 if item["field"] == "@space" and item["current"]["state"] == "missing"
                 and item["proposed"]["state"] == "value"}
    chosen = set(selection)
    for key in selection:
        item = candidates[key]
        required = additions.get(item["spaceId"])
        if required is not None and required not in chosen:
            raise ValueError("Select the new Space along with its values.")
    return {"selection": selection, "message": message, "draftId": str(uuid.uuid4()),
            "createId": str(uuid.uuid4()), "resolveId": str(uuid.uuid4()), "commitId": str(uuid.uuid4())}


def CommitReview(client, review, batch):
    """Stage only selected changes and reuse request IDs when a response is uncertain."""
    draft = {"projectId": review["projectId"], "draftId": batch["draftId"]}
    changes, resolutions = [], []
    for candidateId in batch["selection"]:
        item = review["candidates"][candidateId]
        key = {field: item[field] for field in ("modelId", "spaceId", "field")}
        changes.append(dict(key, staged=True, proposed=item["proposed"]))
        if item["classification"] in ("conflict", "unbased"):
            resolutions.append(dict(key, choice="draft", expectedCurrent=item["current"]))
    created = client.create_draft(dict(draft, requestId=batch["createId"],
                                       collectionId=review["collectionId"], changes=changes))
    revision = created["revision"]
    if resolutions:
        resolved = client.resolve_draft(dict(draft, requestId=batch["resolveId"],
                                             expectedDraftRevision=revision, resolutions=resolutions,
                                             expectedHeadCommitId=review["headCommitId"]))
        revision = resolved["revision"]
    return client.commit_draft(dict(draft, requestId=batch["commitId"],
                                    expectedDraftRevision=revision, message=batch["message"]))


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
