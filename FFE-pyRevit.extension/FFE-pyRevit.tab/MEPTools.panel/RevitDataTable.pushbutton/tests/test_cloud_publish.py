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
                      'parameters': {'Hidden': {'v': 12.5, 's': '12.5 CFM'}, 'Unset': {'v': None}}, 'parameterValues': {'p1': 12.5}, 'lpd': 99, 'occupants': 100}],
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

    def test_preview_keeps_all_parameters_and_preserves_ids_and_warnings(self):
        data = snapshot()
        data['rows'].append({'uniqueId': 'unplaced', 'areaValue': 0, 'heightValue': None})
        data.update(skippedCount=1, issues=['Skipped unreadable Space'])
        review = cloud.PrepareSync(data, PROJECT)
        self.assertEqual(1, review['skippedCount'])
        self.assertEqual(['Skipped unreadable Space'], review['warnings'])
        self.assertEqual({'uniqueId': SPACE, 'elementId': '9007199254740993', 'number': '002',
                          'name': 'Office', 'level': 'Level 2', 'parameters': data['rows'][0]['parameters']},
                         review['request']['spaces'][0])
        self.assertEqual({'uniqueId': 'unplaced', 'number': '', 'name': '', 'level': '', 'parameters': {}},
                         review['request']['spaces'][1])
        self.assertEqual(PROJECT, review['request']['projectId'])
        uuid.UUID(review['reviewId'])
        self.assertEqual(12.5, data['rows'][0]['parameterValues']['p1'])

    def test_group_selection_filters_on_host_and_deduplicates_overlapping_members(self):
        data = snapshot()
        data['rows'] += [dict(data['rows'][0], uniqueId='other', areaValue=float('inf')),
                         dict(data['rows'][0], uniqueId='second', number='003')]
        data['spaceGroups'] = [
            {'id': PROJECT, 'name': 'West', 'members': [SPACE, 'gone']},
            {'id': HEAD, 'name': 'Offices', 'members': [SPACE, 'second']}]
        review = cloud.PrepareSync(data, PROJECT, [PROJECT, HEAD, PROJECT])
        self.assertEqual([SPACE, 'second'], [space['uniqueId'] for space in review['request']['spaces']])
        self.assertEqual({'projectId', 'spaces'}, set(review['request']))
        self.assertFalse(review['scope']['allSpaces'])
        self.assertEqual(1, review['scope']['excludedCount'])
        self.assertEqual(1, review['scope']['unavailableCount'])
        self.assertEqual(2, len(review['scope']['groups']))
        self.assertIn('1 group member', review['warnings'][0])
        # Membership is frozen along with the payload, including across retries.
        data['spaceGroups'][0]['members'] = ['other']
        client = mock.Mock()
        client.push_revit_spaces.return_value = {'created': 2, 'updated': 0, 'unchanged': 0, 'newlyMissing': []}
        cloud.SyncReview(client, review)
        cloud.SyncReview(client, review)
        self.assertEqual(1, client.push_revit_spaces.call_count)
        self.assertEqual([SPACE, 'second'], [space['uniqueId'] for space in client.push_revit_spaces.call_args.args[0]['spaces']])

    def test_empty_unknown_ambiguous_and_unavailable_group_scopes_never_sync_all(self):
        data = snapshot()
        data['spaceGroups'] = [{'id': HEAD, 'name': 'Deleted Spaces', 'members': ['gone']}]
        for ids in ([], [PROJECT], [HEAD]):
            with self.subTest(ids=ids), self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT, ids)
        data['spaceGroups'] *= 2
        with self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT, [HEAD])
        self.assertTrue(cloud.PrepareSync(snapshot(), PROJECT)['scope']['allSpaces'])

    def test_invalid_parameter_values_and_duplicate_ids_fail_before_network(self):
        for change in ({'parameters': {'Bad': {'v': float('inf')}}}, {'parameters': {'Bad': {'v': []}}},
                       {'id': 9007199254740993}, {'uniqueId': ''}):
            data = snapshot(); data['rows'][0].update(change)
            with self.subTest(change=change), self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT)
        data = snapshot(); data['rows'] *= 2
        with self.assertRaises(ValueError): cloud.PrepareSync(data, PROJECT)

    def test_chunk_retry_resumes_the_frozen_batch_and_partial_scope_never_finalizes(self):
        data = snapshot()
        data['rows'] = [dict(data['rows'][0], uniqueId=str(i)) for i in range(401)]
        review = cloud.PrepareSync(data, PROJECT)
        self.assertEqual([200, 200, 1], [len(chunk['spaces']) for chunk in review['chunks']])
        self.assertEqual([False, False, True], [chunk['final'] for chunk in review['chunks']])
        calls = []
        def transport(url, headers, payload):
            request = json.loads(payload)['p_request']
            calls.append((url, headers, request))
            if len(calls) == 2: raise OSError('Connection closed')
            return 200, json.dumps({'ok': True, 'created': len(request['spaces']), 'updated': 0,
                                    'unchanged': 0, 'newlyMissing': [HEAD] if request['final'] else []})
        client = rpc.Client(actor_name='Autodesk user', transport=transport)
        with self.assertRaises(rpc.RpcError) as raised: cloud.SyncReview(client, review)
        self.assertEqual("NETWORK_ERROR", raised.exception.code)
        result = cloud.SyncReview(client, review)
        self.assertEqual(calls[1], calls[2])
        self.assertEqual(401, result['created'])
        self.assertEqual(802, result['parameterCount'])
        self.assertEqual([HEAD], result['newlyMissing'])
        self.assertTrue(calls[0][0].endswith('/rest/v1/rpc/ffe_v3_push_revit_spaces'))
        self.assertEqual({'apikey': rpc.PUBLISHABLE_KEY}, calls[0][1])
        self.assertEqual({'name': 'Autodesk user'}, calls[0][2]['actor'])
        data['spaceGroups'] = [{'id': HEAD, 'name': 'Subset', 'members': ['0']}]
        self.assertFalse(cloud.PrepareSync(data, PROJECT, [HEAD])['chunks'][0]['final'])
        data['skippedCount'] = 1
        self.assertFalse(cloud.PrepareSync(data, PROJECT)['chunks'][-1]['final'])



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
