# -*- coding: utf-8 -*-
"""Purpose: verify repository publishing without network access or Revit.
Behaviors: typed observations, explicitly staged commits, and persistent model identity.
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


sys.modules.setdefault('repository_collection', load('repository_collection'))
cloud = load('cloud_publish')
rpc = load('ffe_supabase')
MODEL = '00000000-0000-4000-8000-000000000001'
PROJECT = '00000000-0000-4000-8000-000000000002'
REVISION = '00000000-0000-4000-8000-000000000003'
HEAD = '00000000-0000-4000-8000-000000000004'
SPACE = 'space-unique'


def snapshot():
    """Include IDs larger than JavaScript's safe integer range and leading zeros."""
    return {'documentKey': 'temporary-bridge', 'documentTitle': 'Mechanical',
            'rows': [{'id': '9007199254740993', 'uniqueId': SPACE, 'number': '002', 'name': 'Office',
                      'parameterValues': {'p1': 12.5}, 'parameterTexts': {'p1': '12.5 L/s'}}],
            'parameters': [{'id': '1', 'baseField': '', 'identity': 'builtin:-200', 'name': 'Airflow', 'key': 'p1',
                            'storageType': 'Double', 'dataTypeId': 'autodesk.spec.aec:airFlow-2.0.0',
                            'unitTypeId': 'autodesk.unit.unit:litersPerSecond-1.0.1'}],
            'selectedParameterIds': ['1'], 'skippedCount': 0, 'issues': [], 'savedTables': ['local']}


def candidate(field, classification='local_change', current=None, value='002'):
    """Create the comparison item returned for a single Space field."""
    return {'modelId': MODEL, 'spaceId': SPACE, 'field': field, 'classification': classification,
            'current': current or {'state': 'missing'}, 'proposed': {'state': 'value', 'value': value}, 'metadata': {}}


class PublishTests(unittest.TestCase):
    def setUp(self):
        self.client = mock.Mock()
        self.project = {'id': PROJECT, 'name': 'FFE Mechanical', 'headCommitId': HEAD}
        self.client.list_projects.return_value = {
            'items': [{'projectId': PROJECT, 'name': 'FFE Mechanical', 'headCommitId': HEAD}], 'nextCursor': None}
        self.client.list_models.return_value = {'items': [{'modelId': MODEL, 'name': 'Mechanical'}], 'nextCursor': None}
        self.client.compare_collection.return_value = {
            'items': [candidate('@space', value={'elementId': '9007199254740993', 'context': {}}),
                      candidate('base:number'), candidate('base:name', 'incoming')], 'nextCursor': None}
        self.client.create_draft.return_value = {'revision': REVISION}
        self.resolved = str(uuid.uuid4())
        self.client.resolve_draft.return_value = {'revision': self.resolved}
        self.client.commit_draft.return_value = {'ok': True, 'changeCount': 2}

    def review(self, linked=True):
        plan = cloud.LoadCatalog(self.client, MODEL, linked)
        return cloud.ReviewCollection(self.client, plan, self.project,
                                      cloud.PrepareCollection(snapshot(), {}), 'Mechanical', {SPACE: '002 Office'})

    def test_catalog_pages_projects_and_finds_saved_model_without_writes(self):
        cursor = {'offset': 1}
        lastPage = self.client.list_projects.return_value
        self.client.list_projects.side_effect = [{'items': [], 'nextCursor': cursor}, lastPage]
        plan = cloud.LoadCatalog(self.client, MODEL, True)
        self.assertEqual({'id': MODEL, 'projectId': PROJECT, 'name': 'Mechanical'}, plan['model'])
        self.assertEqual([mock.call({'limit': 200}), mock.call({'limit': 200, 'cursor': cursor})],
                         self.client.list_projects.call_args_list)
        self.client.register_model.assert_not_called()
        self.client.submit_collection.assert_not_called()

    def test_collection_keeps_typed_values_and_suppresses_removals_when_scan_skips_spaces(self):
        capabilities = {SPACE: {'builtin:-200': {'writable': True}}}
        payload = cloud.PrepareCollection(dict(snapshot(), skippedCount=1), capabilities)
        self.assertEqual({'spaces': 'full', 'successful': False}, payload['coverage'])
        space = payload['spaces'][0]
        self.assertEqual('9007199254740993', space['elementId'])
        self.assertEqual(SPACE, space['uniqueId'])
        cells = {cell['parameter']: cell for cell in space['cells']}
        self.assertEqual('002', cells['base:number']['state']['value'])
        self.assertEqual(12.5, cells['builtin:-200']['state']['value'])
        self.assertEqual('12.5 L/s', cells['builtin:-200']['displayText'])
        self.assertTrue(cells['builtin:-200']['writable'])
        self.assertEqual('unavailable', cells['base:areaValue']['availability'])
        self.assertNotIn('savedTables', payload)

    def test_first_review_registers_and_submits_at_pinned_head_without_committing(self):
        review = self.review(linked=False)
        self.client.register_model.assert_called_once_with({'projectId': PROJECT, 'modelId': MODEL, 'name': 'Mechanical'})
        submitted = self.client.submit_collection.call_args.args[0]
        self.assertEqual(HEAD, submitted['baseCommitId'])
        self.assertEqual(review['collectionId'], submitted['collectionId'])
        self.assertEqual(HEAD, self.client.compare_collection.call_args.args[0]['commitId'])
        self.assertEqual(1, review['incomingCount'])
        self.assertEqual('002 Office', review['candidates']['1']['spaceLabel'])
        self.client.create_draft.assert_not_called()

    def test_review_validates_selection_and_only_stages_chosen_cells(self):
        review = self.review()
        for selection in (['1'], ['0', 'not-reviewed']):
            with self.assertRaises(ValueError):
                cloud.CreateBatch(review, selection, 'Publish')
        batch = cloud.CreateBatch(review, ['0'], 'Create Space')
        result = cloud.CommitReview(self.client, review, batch)
        self.assertTrue(result['ok'])
        self.client.register_model.assert_not_called()
        staged = self.client.create_draft.call_args.args[0]['changes']
        self.assertEqual(['@space'], [change['field'] for change in staged])
        self.assertTrue(staged[0]['staged'])
        self.client.resolve_draft.assert_not_called()
        self.assertEqual('Create Space', self.client.commit_draft.call_args.args[0]['message'])

    def test_explicitly_selected_unbased_value_resolves_against_reviewed_head(self):
        old = {'state': 'value', 'value': '001'}
        self.client.compare_collection.return_value['items'] = [candidate('base:number', 'unbased', old)]
        review = self.review()
        batch = cloud.CreateBatch(review, ['0'], 'Revise Space number')
        cloud.CommitReview(self.client, review, batch)
        resolved = self.client.resolve_draft.call_args.args[0]
        self.assertEqual(HEAD, resolved['expectedHeadCommitId'])
        self.assertEqual(old, resolved['resolutions'][0]['expectedCurrent'])
        self.assertEqual('draft', resolved['resolutions'][0]['choice'])
        self.assertEqual(self.resolved, self.client.commit_draft.call_args.args[0]['expectedDraftRevision'])

    def test_retry_reuses_mutation_ids_and_does_not_retry_a_conflict_automatically(self):
        review = self.review()
        batch = cloud.CreateBatch(review, ['0', '1'], 'Publish')
        self.client.commit_draft.side_effect = [OSError('Connection closed'), {'ok': False, 'code': 'CONFLICT'}]
        with self.assertRaises(OSError):
            cloud.CommitReview(self.client, review, batch)
        self.assertEqual(1, self.client.commit_draft.call_count)
        self.assertFalse(cloud.CommitReview(self.client, review, batch)['ok'])
        self.assertEqual(self.client.create_draft.call_args_list[0], self.client.create_draft.call_args_list[1])
        self.assertEqual(self.client.commit_draft.call_args_list[0], self.client.commit_draft.call_args_list[1])

    def test_rpc_uses_publishable_key_and_versioned_envelope(self):
        calls = []

        def transport(url, headers, payload):
            calls.append((url, headers, json.loads(payload)))
            return 200, '{"ok":true}'

        client = rpc.Client(client_id=MODEL, actor_name='Autodesk user', transport=transport)
        client.commit_draft({'projectId': PROJECT, 'draftId': REVISION, 'requestId': HEAD})
        url, headers, payload = calls[0]
        self.assertTrue(url.endswith('/rest/v1/rpc/ffe_v2_commit_draft'))
        self.assertEqual({'apikey': rpc.PUBLISHABLE_KEY}, headers)
        self.assertEqual({'p_request': {'schemaVersion': 2, 'projectId': PROJECT, 'draftId': REVISION,
                                       'requestId': HEAD, 'actor': {'clientId': MODEL, 'source': 'revit',
                                                                   'name': 'Autodesk user', 'userId': ''}}}, payload)


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
