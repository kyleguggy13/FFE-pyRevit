# -*- coding: utf-8 -*-
"""Purpose: verify named Space groups and selective sync requests offline.
Behaviors: membership persistence, concurrency, rollback and document boundaries.
Revit API: 2025/2026 behavior uses the existing storage and event-host fixtures.
Design: exercise actual validation and storage code with fake SDK boundaries.
"""
import json
import pathlib
import sys
import unittest
import uuid
from types import SimpleNamespace as NS
from unittest.mock import Mock

import test_saved_tables as fixtures
import test_revit_data_table as hostFixtures

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / 'lib'))
from space_groups import RevitSpaceGroupStore, ValidateMembers, ValidateGroup, SPACE_GROUP_SCHEMA_GUID


class GroupStorageTests(unittest.TestCase):
    def setUp(self):
        self.host = fixtures.StoreHost()
        self.store = RevitSpaceGroupStore(self.host.document, self.host.db, self.host.sdk)

    def create(self, name='West wing', members=None):
        return self.store.Write('saveSpaceGroup', '', '', name, members or ['uid-002', 'uid-001', 'uid-002'])

    def test_reads_do_not_create_schema_and_explicit_save_persists_exact_members(self):
        self.assertEqual(([], []), self.store.List())
        self.assertEqual({}, self.host.schemas)
        self.assertEqual([], self.host.transactions)
        group = self.create()
        self.assertEqual(['uid-001', 'uid-002'], group['members'])
        self.assertEqual(group, self.store.Read(group['id'])[1])
        self.assertIn(SPACE_GROUP_SCHEMA_GUID, self.host.schemas)
        self.assertEqual(1, len(self.host.transactions))
        self.assertTrue(self.host.transactions[0].disposed)

    def test_edit_rename_delete_and_same_named_groups_are_independent(self):
        first, second = self.create(), self.create()
        self.assertNotEqual(first['id'], second['id'])
        self.host.document.IsWorkshared = True
        changed = self.store.Write('saveSpaceGroup', first['id'], first['revision'], 'New name', ['uid-003'])
        self.assertEqual('New name', changed['name'])
        self.assertEqual(['uid-003'], changed['members'])
        self.assertNotEqual(first['revision'], changed['revision'])
        self.assertEqual([1], [item.Value for item in self.host.checkout.call_args.args[1]])
        self.assertEqual(second, self.store.Read(second['id'])[1])
        self.store.Write('deleteSpaceGroup', first['id'], changed['revision'])
        self.assertEqual([second], [record for element, record in self.store.List()[0]])

    def test_stale_revision_ownership_and_central_changes_never_overwrite(self):
        record = self.create()
        self.host.document.IsWorkshared = True
        for revision, allowed, status in ((str(uuid.uuid4()), True, 'CurrentWithCentral'),
                                          (record['revision'], False, 'CurrentWithCentral'),
                                          (record['revision'], True, 'UpdatedInCentral')):
            self.host.allowCheckout, self.host.centralStatus = allowed, status
            before = self.host.elements[0].text
            with self.assertRaises(ValueError):
                self.store.Write('saveSpaceGroup', record['id'], revision, 'Unsafe', ['other'])
            self.assertEqual(before, self.host.elements[0].text)
            self.assertEqual(1, len(self.host.transactions))

    def test_checkout_race_and_readonly_preserve_members(self):
        record = self.create()
        self.host.document.IsWorkshared = True
        def checkout(document, ids):
            value = json.loads(self.host.elements[0].text)
            value['revision'] = str(uuid.uuid4())
            self.host.elements[0].text = json.dumps(value)
            return ids
        self.host.checkout.side_effect = checkout
        with self.assertRaisesRegex(ValueError, 'newer revision'):
            self.store.Write('deleteSpaceGroup', record['id'], record['revision'])
        self.host.document.IsReadOnly = True
        self.assertEqual(record['members'], self.store.Read(record['id'])[1]['members'])
        with self.assertRaises(ValueError): self.create()
        self.assertEqual(1, len(self.host.transactions))

    def test_failed_create_update_and_delete_roll_back_and_dispose(self):
        self.host.commitFails = True
        with self.assertRaises(ValueError): self.create()
        self.assertEqual([], self.host.elements)
        self.host.commitFails = False
        record = self.create()
        before = self.host.elements[0].text
        self.host.commitFails = True
        for action in ('saveSpaceGroup', 'deleteSpaceGroup'):
            with self.assertRaises(ValueError):
                self.store.Write(action, record['id'], record['revision'], 'Unsafe', ['other'])
            self.assertEqual(before, self.host.elements[0].text)
            self.assertTrue(self.host.transactions[-1].rolledBack)
            self.assertTrue(self.host.transactions[-1].disposed)

    def test_invalid_and_future_records_are_reported_and_preserved(self):
        record = self.create()
        self.host.elements[0].text = json.dumps(dict(record, version=2))
        before = self.host.elements[0].text
        records, issues = self.store.List()
        self.assertEqual([], records)
        self.assertIn('Space group', issues[0])
        self.assertEqual(before, self.host.elements[0].text)
        for members in ([], [''], [' '], [1], [' padded '], 'uid', ['x' * 201]):
            with self.subTest(members=members), self.assertRaises(ValueError): ValidateMembers(members)
        for name in ('', ' ', 'x' * 101):
            with self.assertRaises(ValueError): ValidateGroup(dict(record, name=name))


class GroupRequestTests(unittest.TestCase):
    def setUp(self):
        self.groupId, self.revision = str(uuid.uuid4()), str(uuid.uuid4())
        self.snapshot = {'rows': [{'uniqueId': 'uid-001'}], 'spaceGroups': [
            {'id': self.groupId, 'members': ['uid-deleted']}]}
        self.request = {'action': 'saveSpaceGroup', 'documentKey': 'doc', 'requestId': 'group-1',
                        'groupId': self.groupId, 'revision': self.revision,
                        'name': ' West ', 'members': ['uid-001', 'uid-deleted', 'uid-001']}

    def test_requests_are_document_bound_and_preserve_missing_existing_members(self):
        result = hostFixtures.core.ValidateRequest(self.request, 'doc', self.snapshot)
        self.assertEqual(['uid-001', 'uid-deleted'], result['members'])
        self.assertEqual('West', result['name'])
        for changes in ({'documentKey': 'other'}, {'requestId': None}, {'members': ['foreign']},
                        {'members': []}, {'name': ' '}, {'revision': ''}, {'groupId': 'invalid'}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                hostFixtures.core.ValidateRequest(dict(self.request, **changes), 'doc', self.snapshot)
        with self.assertRaises(ValueError):
            hostFixtures.core.ValidateRequest(dict(self.request, groupId='', revision='', members=['uid-deleted']), 'doc', self.snapshot)

    def test_sync_scope_requires_nonempty_valid_group_ids_and_ignores_browser_rows(self):
        request = {'action': 'publishReview', 'documentKey': 'doc', 'requestId': 'publish-1',
                   'projectId': str(uuid.uuid4()), 'groupIds': [self.groupId, self.groupId], 'spaces': ['foreign']}
        result = hostFixtures.core.ValidateRequest(request, 'doc', self.snapshot)
        self.assertEqual([self.groupId], result['groupIds'])
        self.assertNotIn('spaces', result)
        for groupIds in ([], 'all', [1], ['invalid']):
            with self.assertRaises(ValueError):
                hostFixtures.core.ValidateRequest(dict(request, groupIds=groupIds), 'doc', self.snapshot)

    def test_group_writes_are_queued_in_api_context_and_invalidate_sync_preview(self):
        api = hostFixtures.host_classes()
        window = object.__new__(api['DataTableWindow'])
        host = fixtures.StoreHost()
        window.Document = host.document
        window.GroupStore = RevitSpaceGroupStore(host.document, host.db, host.sdk)
        window.Snapshot = dict(self.snapshot)
        window.CloudBusy = False
        window.DocumentRevision = 0
        window.PublishReview = {'reviewId': 'old-review'}
        window.Send = Mock()
        window.SendResult = Mock()
        window.DocumentKey = 'doc'
        window.IndexUri = 'file:///index.html'
        window.DocumentActive = True
        window.IsClosed = False
        window.Handler = api['DataTableEventHandler']()
        window.Handler.Window = window
        window.ExternalEvent = NS(Raise=lambda: 'Accepted')
        host.document.IsValidObject = True
        host.document.Equals = lambda document: document is host.document
        window.OnWebMessage(None, NS(Source=window.IndexUri, WebMessageAsJson=json.dumps(
            dict(self.request, groupId='', revision='', members=['uid-001']))))
        self.assertEqual([], host.transactions)
        self.assertEqual(1, len(window.Handler.Requests))
        window.Handler.Execute(NS(ActiveUIDocument=NS(Document=host.document)))
        response = window.Send.call_args.args[0]
        self.assertEqual('groupResult', response['type'])
        self.assertEqual(['uid-001'], response['catalog']['spaceGroups'][0]['members'])
        self.assertEqual(1, window.DocumentRevision)
        self.assertIsNone(window.PublishReview)
        self.assertFalse(window.WritingTables)
        window.CloudBusy = True
        with self.assertRaises(ValueError): window.HandleGroupRequest(self.request)
        self.assertEqual(1, len(host.transactions))


if __name__ == '__main__':
    unittest.main()
