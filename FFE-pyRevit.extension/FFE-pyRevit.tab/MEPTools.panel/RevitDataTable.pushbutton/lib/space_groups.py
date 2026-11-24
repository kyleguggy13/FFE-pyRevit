# -*- coding: utf-8 -*-
"""Purpose: persist named sets of Spaces for selective Loadwise sync.
Behaviors: stable UniqueId membership, independent records and revision checks.
Revit API: 2025/2026; all storage access runs inside an API callback.
Design: one DataStorage per group avoids locking a shared catalog. Groups are
metadata only; reading never creates storage and writes never edit Space values.
"""
import json
import time
import uuid

from saved_tables import RevitTableStore, RequireText

SPACE_GROUP_SCHEMA_GUID = "aacb3874-e9c6-4c70-8cec-746c8118c0c6"
GroupActions = ("saveSpaceGroup", "deleteSpaceGroup")


def ValidateMembers(members):
    """Keep exact Space UniqueIds, with bounded size and no duplicate members."""
    if not isinstance(members, list) or not members or len(members) > 50000:
        raise ValueError("Choose between 1 and 50,000 Spaces for the group.")
    result = set()
    for member in members:
        text = RequireText(member, "Space UniqueId", 200)
        if not text or text != text.strip():
            raise ValueError("Invalid Space UniqueId in the group.")
        result.add(text)
    return sorted(result)


def ValidateGroup(record):
    """Reject corrupt/future metadata without rewriting a shared definition."""
    if not isinstance(record, dict) or record.get("version") != 1:
        raise ValueError("Unsupported Space group version.")
    name = RequireText(record.get("name"), "group name", 100).strip()
    if not name:
        raise ValueError("Enter a group name of 1 to 100 characters.")
    return {"version": 1, "id": str(uuid.UUID(record["id"])),
            "revision": str(uuid.UUID(record["revision"])), "name": name,
            "members": ValidateMembers(record.get("members")),
            "modifiedBy": RequireText(record.get("modifiedBy", ""), "author"),
            "modifiedAt": RequireText(record.get("modifiedAt", ""), "timestamp")}


class RevitSpaceGroupStore(RevitTableStore):
    """Store named groups independently of table layouts and the cloud connection."""

    def GetSchema(self, create=False):
        """Register a separate schema only on an explicit group save."""
        schema = self.SDK["Schema"].Lookup(self.SDK["Guid"](SPACE_GROUP_SCHEMA_GUID))
        if schema is None and create:
            builder = self.SDK["SchemaBuilder"](self.SDK["Guid"](SPACE_GROUP_SCHEMA_GUID))
            builder.SetSchemaName("FFEHVACSpaceGroupsV1")
            builder.SetReadAccessLevel(self.SDK["AccessLevel"].Public)
            builder.SetWriteAccessLevel(self.SDK["AccessLevel"].Public)
            builder.AddSimpleField("DefinitionJson", self.SDK["String"])
            schema = builder.Finish()
        return schema

    def ReadElement(self, element, schema):
        """Read this group's bounded JSON and validate its version and identities."""
        entity = element.GetEntity(schema)
        if not entity.IsValid():
            return None
        text = entity.Get[self.SDK["String"]](schema.GetField("DefinitionJson"))
        if len(text) > 1000000:
            raise ValueError("Space group exceeds the supported size.")
        return ValidateGroup(json.loads(text))

    def List(self):
        """Reuse best-effort collection while naming group errors accurately."""
        records, issues = RevitTableStore.List(self)
        return records, [issue.replace("Saved table", "Space group") for issue in issues]

    def Read(self, groupId):
        """Require an unambiguous ID even when two groups share the same name."""
        try:
            return RevitTableStore.Read(self, groupId)
        except ValueError:
            raise ValueError("The Space group is missing or ambiguous. Refresh and retry.")

    def Write(self, action, groupId, revision, name=None, members=None):
        """Checkout and recheck a single group, then commit or roll back metadata."""
        if action not in GroupActions:
            raise ValueError("Invalid Space group operation.")
        if action == "deleteSpaceGroup" and not groupId:
            raise ValueError("Choose a Space group first.")
        element, previous = self.Read(groupId) if groupId else (None, None)
        try:
            self.CheckEditable(element)
        except ValueError as error:
            raise ValueError(str(error).replace("saved table", "Space group")
                             .replace("use Save As", "create a new group"))
        if element is not None:
            previous = self.ReadElement(element, self.GetSchema())
            if previous is None or previous["revision"] != revision:
                raise ValueError("The Space group has a newer revision. Refresh and reopen it before saving.")
        record = None
        if action == "saveSpaceGroup":
            record = ValidateGroup({"version": 1, "id": groupId or str(uuid.uuid4()),
                                    "revision": str(uuid.uuid4()), "name": name, "members": members,
                                    "modifiedBy": self.Document.Application.Username or "",
                                    "modifiedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())})
            if len(json.dumps(record, ensure_ascii=True)) > 1000000:
                raise ValueError("Space group exceeds the supported size.")
        transaction = self.DB.Transaction(self.Document, "Save HVAC Space group" if record else "Delete HVAC Space group")
        try:
            if transaction.Start() != self.DB.TransactionStatus.Started:
                raise ValueError("Revit could not start the Space group transaction.")
            if record is None:
                self.Document.Delete(element.Id)
            else:
                schema = self.GetSchema(create=True)
                if element is None:
                    element = self.SDK["DataStorage"].Create(self.Document)
                entity = self.SDK["Entity"](schema)
                entity.Set[self.SDK["String"]](schema.GetField("DefinitionJson"), json.dumps(record, ensure_ascii=True))
                element.SetEntity(entity)
            if transaction.Commit() != self.DB.TransactionStatus.Committed:
                raise ValueError("Revit did not commit the Space group operation.")
        except Exception:
            if transaction.GetStatus() == self.DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        finally:
            transaction.Dispose()
        return record
