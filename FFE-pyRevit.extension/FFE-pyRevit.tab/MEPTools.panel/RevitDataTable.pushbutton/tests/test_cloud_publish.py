# -*- coding: utf-8 -*-
"""Purpose: verify explicit cloud publishing without network access or Revit.
Behaviors: complete JSON exports, conflict preservation, persistent model identity.
Revit API: 2025/2026 boundaries use the saved-table storage fixtures.
Design: real helpers and RPC payloads run against injected storage/transports.
"""
import importlib.util
import json
import pathlib
import sys
import unittest
import uuid
from unittest import mock

import test_saved_tables as fixtures

LIBRARY = pathlib.Path(__file__).resolve().parents[1] / 'lib'
sys.modules.setdefault('saved_tables', fixtures.saved)


def load(name):
    """Import a pure command module without loading the native Revit host."""
    spec = importlib.util.spec_from_file_location(name, LIBRARY / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


cloud = load('cloud_publish')
rpc = load('ffe_supabase')
MODEL = '00000000-0000-4000-8000-000000000001'
PROJECT = '00000000-0000-4000-8000-000000000002'
REVISION = '00000000-0000-4000-8000-000000000003'


def snapshot():
    """Include strings larger than JavaScript's safe integer range and leading zeros."""
    return {'documentKey': 'temporary-bridge', 'documentTitle': 'Mechanical',
            'rows': [{'id': '9007199254740993', 'number': '002', 'parameterValues': {'p1': 12.5}}],
            'selectedParameterIds': ['1'], 'skippedCount': 0, 'savedTables': ['local'],
            'tableIssues': [], 'canSaveTables': True}


class PublishTests(unittest.TestCase):
    def setUp(self):
        self.client = mock.Mock()
        self.client.list_projects.return_value = [{'id': PROJECT, 'name': 'FFE Mechanical'}]
        self.model = {'id': MODEL, 'projectId': PROJECT, 'name': 'Mechanical', 'revision': REVISION}
        self.client.get_snapshot.return_value = dict(self.model, snapshot={'old': 'data'})
        self.client.create_model.return_value = self.model
        self.client.publish_snapshot.return_value = dict(self.model, revision=str(uuid.uuid4()))

    def test_catalog_is_read_only_and_captures_reviewed_revision(self):
        plan = cloud.LoadCatalog(self.client, MODEL)
        self.assertEqual(self.model, plan['model'])
        self.assertNotIn('snapshot', plan['model'])
        self.client.create_model.assert_not_called()
        self.client.publish_snapshot.assert_not_called()

    def test_snapshot_copy_preserves_ids_numbers_and_values_without_local_metadata(self):
        original = snapshot()
        payload = cloud.PrepareSnapshot(original, MODEL)
        self.assertEqual(MODEL, payload['documentKey'])
        self.assertEqual('temporary-bridge', original['documentKey'])
        self.assertEqual('9007199254740993', payload['rows'][0]['id'])
        self.assertEqual('002', payload['rows'][0]['number'])
        self.assertEqual(12.5, payload['rows'][0]['parameterValues']['p1'])
        self.assertFalse(set(('savedTables', 'tableIssues', 'canSaveTables')) & set(payload))
        payload['rows'].clear()
        self.assertEqual(1, len(original['rows']))

    def test_incomplete_or_nonfinite_snapshot_does_not_create_or_publish(self):
        plan = cloud.LoadCatalog(self.client, MODEL)
        for update in ({'skippedCount': 1}, {'rows': [{'value': float('nan')}]}):
            with self.subTest(update=update), self.assertRaises(ValueError):
                cloud.PublishSnapshot(self.client, plan, PROJECT, dict(snapshot(), **update))
        self.client.create_model.assert_not_called()
        self.client.publish_snapshot.assert_not_called()

    def test_existing_model_uses_reviewed_revision_and_never_refetches_before_write(self):
        plan = cloud.LoadCatalog(self.client, MODEL)
        self.client.get_snapshot.reset_mock()
        result = cloud.PublishSnapshot(self.client, plan, PROJECT, snapshot())
        args = self.client.publish_snapshot.call_args.args
        self.assertEqual((MODEL, REVISION), args[:2])
        self.assertEqual((1, 1), (result['rowCount'], result['parameterCount']))
        self.client.get_snapshot.assert_not_called()
        self.client.create_model.assert_not_called()

    def test_first_publish_registers_stable_model_then_publishes_empty_dataset(self):
        self.client.get_snapshot.return_value = None
        plan = cloud.LoadCatalog(self.client, MODEL)
        result = cloud.PublishSnapshot(self.client, plan, PROJECT, dict(snapshot(), rows=[]))
        self.client.create_model.assert_called_once_with(PROJECT, MODEL, 'Mechanical')
        self.assertEqual(0, result['rowCount'])

    def test_conflicts_are_not_retried_and_wrong_project_cannot_write(self):
        plan = cloud.LoadCatalog(self.client, MODEL)
        self.client.publish_snapshot.side_effect = rpc.RpcError('Changed', '40001', 400)
        with self.assertRaises(rpc.RpcError):
            cloud.PublishSnapshot(self.client, plan, PROJECT, snapshot())
        self.assertEqual(1, self.client.publish_snapshot.call_count)
        other = str(uuid.uuid4())
        plan['projects'].append({'id': other, 'name': 'Other'})
        for project in (other, str(uuid.uuid4())):
            with self.assertRaises(ValueError):
                cloud.PublishSnapshot(self.client, plan, project, snapshot())
        self.assertEqual(1, self.client.publish_snapshot.call_count)

    def test_rpc_uses_publishable_key_and_database_signature_without_auth_token(self):
        calls = []

        def transport(url, headers, payload):
            calls.append((url, headers, json.loads(payload)))
            return 200, json.dumps(self.model)

        client = rpc.Client(modified_by='Autodesk user', transport=transport)
        client.publish_snapshot(MODEL, REVISION, snapshot())
        url, headers, payload = calls[0]
        self.assertTrue(url.endswith('/rest/v1/rpc/ffe_publish_snapshot'))
        self.assertEqual({'apikey': rpc.PUBLISHABLE_KEY}, headers)
        self.assertEqual({'p_model_id': MODEL, 'p_expected_revision': REVISION,
                          'p_snapshot': snapshot(), 'p_modified_by': 'Autodesk user'}, payload)


class IdentityTests(unittest.TestCase):
    def setUp(self):
        self.host = fixtures.StoreHost()
        info = fixtures.DataStorage(42)
        self.host.elements.append(info)
        self.host.document.ProjectInformation = info
        self.store = cloud.CloudModelStore(self.host.document, self.host.db, self.host.sdk)

    def test_read_creates_nothing_and_ensure_writes_only_once(self):
        self.assertIsNone(self.store.Read())
        self.assertEqual({}, self.host.schemas)
        self.assertEqual([], self.host.transactions)
        self.assertTrue(self.store.Ensure(MODEL))
        self.assertEqual(MODEL, self.store.Read())
        self.assertFalse(self.store.Ensure(MODEL))
        self.assertEqual(1, len(self.host.transactions))
        self.assertTrue(self.host.transactions[0].disposed)

    def test_readonly_unlinked_model_and_other_owner_cannot_create_identity(self):
        self.host.document.IsReadOnly = True
        with self.assertRaises(ValueError):
            self.store.Ensure(MODEL)
        self.host.document.IsReadOnly = False
        self.host.document.IsWorkshared = True
        self.host.allowCheckout = False
        with self.assertRaises(ValueError):
            self.store.Ensure(MODEL)
        self.assertEqual([], self.host.transactions)

    def test_commit_failure_rolls_back_identity_and_disposes_transaction(self):
        self.host.commitFails = True
        with self.assertRaises(ValueError):
            self.store.Ensure(MODEL)
        self.assertIsNone(self.store.Read())
        self.assertTrue(self.host.transactions[0].rolledBack)
        self.assertTrue(self.host.transactions[0].disposed)

    def test_linked_readonly_model_reuses_identity_but_never_rebinds(self):
        self.store.Ensure(MODEL)
        self.host.document.IsReadOnly = True
        self.assertFalse(self.store.Ensure(MODEL))
        with self.assertRaises(ValueError):
            self.store.Ensure(str(uuid.uuid4()))
        self.assertEqual(MODEL, self.store.Read())


if __name__ == '__main__':
    unittest.main()
