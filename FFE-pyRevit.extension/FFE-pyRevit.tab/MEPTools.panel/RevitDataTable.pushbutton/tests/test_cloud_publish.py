# -*- coding: utf-8 -*-
"""Purpose: verify Loadwise v3 sync, field ownership and project storage offline.
Behaviors: no preview writes, exact IDs, repeatable sync and transaction rollback.
Revit API: 2025/2026 storage uses the saved-table fixtures.
Design: real command helpers run with injected transports and storage boundaries.
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
sys.path.insert(0, str(LIBRARY))
sys.modules.setdefault('saved_tables', fixtures.saved)


def load(name):
    spec = importlib.util.spec_from_file_location(name, LIBRARY / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


cloud = load('cloud_publish')
rpc = load('ffe_supabase')
PROJECT = '00000000-0000-4000-8000-000000000002'
HEAD = '00000000-0000-4000-8000-000000000004'
SPACE = 'space-unique'


def snapshot():
    return {'documentKey': 'temporary-bridge', 'documentTitle': 'Mechanical',
            'rows': [{'id': '9007199254740993', 'uniqueId': SPACE, 'number': '002', 'name': 'Office',
                      'level': 'Level 2', 'areaValue': 120.000004, 'heightValue': 9.0,
                      'parameterValues': {'p1': 12.5}, 'lpd': 99, 'occupants': 100}],
            'selectedParameterIds': ['1'], 'skippedCount': 0, 'issues': [], 'savedTables': ['local']}


class PublishTests(unittest.TestCase):
    def test_catalog_uses_v3_without_registry_or_pagination_and_honors_binding(self):
        client = mock.Mock()
        client.list_projects.return_value = {'ok': True, 'items': [{'projectId': PROJECT, 'name': 'Mechanical'}]}
        self.assertEqual({'projects': [{'id': PROJECT, 'name': 'Mechanical'}], 'projectId': PROJECT},
                         cloud.LoadCatalog(client, PROJECT, PROJECT))
        client.list_projects.assert_called_once_with()
        self.assertEqual([mock.call.list_projects()], client.mock_calls)
        with self.assertRaises(ValueError): cloud.LoadCatalog(client, PROJECT, HEAD)
        with self.assertRaises(ValueError): cloud.LoadCatalog(client, HEAD)

    def test_cloud_catalog_rechecks_forma_and_never_offers_a_foreign_manual_project(self):
        client = mock.Mock()
        client.list_projects.return_value = {'items': [{'projectId': PROJECT, 'name': 'Mechanical'}]}
        context = {'available': True, 'formaProjectId': HEAD}
        client.find_forma_project.return_value = {'project': {'projectId': PROJECT}}
        self.assertEqual(PROJECT, cloud.LoadCatalog(client, formaContext=context)['projectId'])
        client.find_forma_project.assert_called_with({'formaProjectId': HEAD})
        with self.assertRaises(ValueError): cloud.LoadCatalog(client, HEAD, formaContext=context)
        with self.assertRaises(ValueError): cloud.LoadCatalog(client, preferredProjectId=HEAD, formaContext=context)
        client.find_forma_project.return_value = {'project': None}
        with self.assertRaisesRegex(ValueError, 'Start this Autodesk Forma project'):
            cloud.LoadCatalog(client, formaContext=context)
        with self.assertRaises(ValueError): cloud.LoadCatalog(client, formaContext={'available': False})
        client.find_forma_project.side_effect = OSError('Offline')
        with self.assertRaises(OSError): cloud.LoadCatalog(client, formaContext=context)
        client.ensure_forma_project.assert_not_called()

    def test_preview_has_only_owned_fields_and_preserves_zero_missing_ids_and_warnings(self):
        data = snapshot()
        data['rows'].append({'uniqueId': 'unplaced', 'areaValue': 0, 'heightValue': None})
        data.update(skippedCount=1, issues=['Skipped unreadable Space'])
        review = cloud.PrepareSync(data, PROJECT)
        self.assertEqual(1, review['skippedCount'])
        self.assertEqual(['Skipped unreadable Space'], review['warnings'])
        self.assertEqual({'uniqueId': SPACE, 'elementId': '9007199254740993', 'number': '002',
                          'name': 'Office', 'level': 'Level 2', 'area': 120.0, 'ceilingHeight': 9.0},
                         review['request']['spaces'][0])
        self.assertEqual({'uniqueId': 'unplaced', 'number': '', 'name': '', 'level': '', 'area': 0.0},
                         review['request']['spaces'][1])
        self.assertEqual(PROJECT, review['request']['projectId'])
        uuid.UUID(review['reviewId'])
        self.assertEqual(12.5, data['rows'][0]['parameterValues']['p1'])

    def test_invalid_geometry_and_duplicate_ids_fail_before_network(self):
        for change in ({'areaValue': float('inf')}, {'heightValue': float('nan')}, {'heightValue': 201},
                       {'areaValue': 1000001}, {'id': 9007199254740993}, {'uniqueId': ''}):
            data = snapshot()
            data['rows'][0].update(change)
            with self.subTest(change=change), self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT)
        data = snapshot(); data['rows'] *= 2
        with self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT)
        data['rows'] = [dict(snapshot()['rows'][0], uniqueId=str(i)) for i in range(5001)]
        with self.assertRaisesRegex(ValueError, '5000'): cloud.PrepareSync(data, PROJECT)

    def test_sync_and_retry_use_the_same_v3_payload_and_return_missing_rooms(self):
        calls = []
        result = {'ok': True, 'created': 1, 'updated': 2, 'unchanged': 3, 'linked': 1, 'notInRevit': [HEAD]}
        def transport(url, headers, payload):
            calls.append((url, headers, json.loads(payload)))
            if len(calls) == 1: raise OSError('Connection closed')
            return 200, json.dumps(result)
        client = rpc.Client(actor_name='Autodesk user', transport=transport)
        review = cloud.PrepareSync(snapshot(), PROJECT)
        self.assertEqual([], calls)
        with self.assertRaises(OSError): cloud.SyncReview(client, review)
        self.assertEqual(result, cloud.SyncReview(client, review))
        self.assertEqual(calls[0], calls[1])
        url, headers, payload = calls[0]
        self.assertTrue(url.endswith('/rest/v1/rpc/ffe_v3_sync_revit_spaces'))
        self.assertEqual({'apikey': rpc.PUBLISHABLE_KEY}, headers)
        self.assertEqual(3, payload['p_request']['schemaVersion'])
        self.assertEqual({'name': 'Autodesk user'}, payload['p_request']['actor'])
        self.assertEqual({'projectId', 'spaces', 'schemaVersion', 'actor'}, set(payload['p_request']))
        with self.assertRaises(ValueError): client.call('commit_draft')


class IdentityTests(unittest.TestCase):
    def setUp(self):
        self.host = fixtures.StoreHost()
        info = fixtures.DataStorage(42)
        self.host.elements.append(info)
        self.host.document.ProjectInformation = info
        self.store = cloud.CloudProjectStore(self.host.document, self.host.db, self.host.sdk)

    def test_read_creates_nothing_and_ensure_writes_only_once(self):
        self.assertIsNone(self.store.Read())
        self.assertEqual({}, self.host.schemas)
        self.assertEqual([], self.host.transactions)
        self.assertTrue(self.store.Ensure(PROJECT))
        self.assertEqual(PROJECT, self.store.Read())
        self.assertFalse(self.store.Ensure(PROJECT))
        self.assertEqual(1, len(self.host.transactions))
        self.assertTrue(self.host.transactions[0].disposed)

    def test_readonly_unlinked_model_and_other_owner_cannot_create_identity(self):
        self.host.document.IsReadOnly = True
        with self.assertRaises(ValueError):
            self.store.Ensure(PROJECT)
        self.host.document.IsReadOnly = False
        self.host.document.IsWorkshared = True
        self.host.allowCheckout = False
        with self.assertRaises(ValueError):
            self.store.Ensure(PROJECT)
        self.assertEqual([], self.host.transactions)

    def test_commit_failure_rolls_back_identity_and_disposes_transaction(self):
        self.host.commitFails = True
        with self.assertRaises(ValueError):
            self.store.Ensure(PROJECT)
        self.assertIsNone(self.store.Read())
        self.assertTrue(self.host.transactions[0].rolledBack)
        self.assertTrue(self.host.transactions[0].disposed)

    def test_linked_readonly_model_reuses_identity_but_never_rebinds(self):
        self.store.Ensure(PROJECT)
        self.host.document.IsReadOnly = True
        self.assertFalse(self.store.Ensure(PROJECT))
        with self.assertRaises(ValueError):
            self.store.Ensure(str(uuid.uuid4()))
        self.assertEqual(PROJECT, self.store.Read())


if __name__ == '__main__':
    unittest.main()
