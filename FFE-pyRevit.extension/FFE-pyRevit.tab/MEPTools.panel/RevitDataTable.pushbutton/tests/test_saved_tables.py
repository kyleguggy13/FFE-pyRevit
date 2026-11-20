# -*- coding: utf-8 -*-
"""Purpose: verify shared layouts, stable identities, and model-storage transactions.
Behaviors: round trips, missing references, revision/ownership failures, and rollback.
Revit API notes: injected SDK boundaries exercise the real storage implementation.
Design: no Revit process or document mutation is needed for automated validation.
"""
import importlib.util
import json
import pathlib
import types
import unittest
import uuid
from unittest import mock

path = pathlib.Path(__file__).resolve().parents[1] / 'lib/saved_tables.py'
spec = importlib.util.spec_from_file_location('saved_tables_test_module', path)
saved = importlib.util.module_from_spec(spec)
spec.loader.exec_module(saved)
NS = types.SimpleNamespace
SHARED_GUID = '6cba6914-a10e-4be7-900a-b8a85fb041d8'


def snapshot():
    return {'defaultPhaseId': '10', 'phases': [{'id': '10', 'uniqueId': 'phase-uid'}],
            'levels': [{'id': '20', 'uniqueId': 'level-uid'}],
            'parameters': [{'id': '201', 'key': 'p201', 'name': 'Comments',
                            'identity': 'shared:' + SHARED_GUID}]}


def runtime():
    return {'columns': [{'field': field, 'width': 120 + index}
                        for index, field in enumerate(saved.BASE_FIELDS)] + [{'field': 'parameterValues.p201', 'width': 210}],
            'sorters': [{'field': 'number', 'dir': 'desc'}, {'field': 'parameterValues.p201', 'dir': 'asc'}],
            'groupFields': ['level', 'parameterValues.p201', 'name'],
            'headerGroups': [{'name': 'Identity', 'fields': ['number', 'name']},
                             {'name': 'Notes', 'fields': ['parameterValues.p201']}],
            'filters': {'phase': '10', 'level': '20', 'query': 'Office'}}


class LayoutTests(unittest.TestCase):
    def test_colors_round_trip_with_stable_parameter_identities(self):
        layout = runtime()
        layout['columns'][-1]['color'] = '#123456'
        layout['headerGroups'][1]['color'] = '#abcdef'
        portable = saved.CaptureLayout(layout, snapshot())
        current = snapshot()
        current['parameters'][0].update(id='999', key='p999')
        restored, warnings = saved.ResolveLayout(portable, current)
        self.assertEqual('#123456', restored['columns'][-1]['color'])
        self.assertEqual('parameterValues.p999', restored['columns'][-1]['field'])
        self.assertEqual({'name': 'Notes', 'fields': ['parameterValues.p999'], 'color': '#abcdef'},
                         restored['headerGroups'][1])
        self.assertEqual([], warnings)

    def test_invalid_colors_are_rejected_on_columns_and_groups(self):
        for value in (None, 123, {}, '#fff', 'red', '#123456\n', '#123456; color:red'):
            for key in ('columns', 'headerGroups'):
                layout = runtime()
                layout[key][0]['color'] = value
                with self.subTest(value=value, key=key), self.assertRaises(ValueError):
                    saved.ValidateLayout(layout)
        layout = runtime()
        layout['columns'][0]['color'] = '#AABBCC'
        self.assertEqual('#aabbcc', saved.ValidateLayout(layout)['columns'][0]['color'])

    def test_complete_layout_round_trip_preserves_order_width_sort_group_and_filters(self):
        portable = saved.CaptureLayout(runtime(), snapshot())
        self.assertEqual('phase-uid', portable['filters']['phase'])
        self.assertEqual('level-uid', portable['filters']['level'])
        self.assertEqual(['base:level', 'shared:' + SHARED_GUID, 'base:name'], portable['groupFields'])
        self.assertEqual({'name': 'Notes', 'fields': ['shared:' + SHARED_GUID]}, portable['headerGroups'][1])
        restored, warnings = saved.ResolveLayout(portable, snapshot())
        self.assertEqual(runtime(), restored)
        self.assertEqual([], warnings)

    def test_rebound_shared_parameter_resolves_new_element_id_without_name_matching(self):
        portable = saved.CaptureLayout(runtime(), snapshot())
        current = snapshot()
        current['parameters'][0].update(id='999', key='p999', name='Renamed')
        restored, warnings = saved.ResolveLayout(portable, current)
        self.assertEqual(['level', 'parameterValues.p999', 'name'], restored['groupFields'])
        self.assertEqual('parameterValues.p999', restored['columns'][-1]['field'])
        self.assertEqual({'name': 'Notes', 'fields': ['parameterValues.p999']}, restored['headerGroups'][1])
        self.assertEqual([], warnings)

    def test_missing_project_parameter_clears_dependent_sort_and_group_but_keeps_definition(self):
        original = snapshot()
        original['parameters'][0]['identity'] = 'project:original-parameter'
        portable = saved.CaptureLayout(runtime(), original)
        before = json.dumps(portable, sort_keys=True)
        current = snapshot()
        current['parameters'][0]['identity'] = 'project:replacement-with-same-name'
        restored, warnings = saved.ResolveLayout(portable, current)
        self.assertEqual(['level', 'name'], restored['groupFields'])
        self.assertEqual([{'name': 'Identity', 'fields': ['number', 'name']}], restored['headerGroups'])
        self.assertEqual([{'field': 'number', 'dir': 'desc'}], restored['sorters'])
        self.assertEqual(6, len(restored['columns']))
        self.assertGreaterEqual(len(warnings), 3)
        self.assertEqual(before, json.dumps(portable, sort_keys=True))

    def test_missing_phase_and_level_use_initial_defaults_with_warnings(self):
        portable = saved.CaptureLayout(runtime(), snapshot())
        current = snapshot()
        current['phases'][0]['uniqueId'] = 'new-phase'
        current['levels'] = []
        restored, warnings = saved.ResolveLayout(portable, current)
        self.assertEqual({'phase': '10', 'level': 'all', 'query': 'Office'}, restored['filters'])
        self.assertEqual(2, len(warnings))

    def test_layout_validation_rejects_duplicate_unknown_and_nonfinite_inputs(self):
        for value in (None, {}, {'columns': []}):
            with self.subTest(value=value), self.assertRaises(ValueError):
                saved.ValidateLayout(value)
        for width in (float('nan'), float('inf'), True, 0, 5000):
            layout = runtime()
            layout['columns'][0]['width'] = width
            with self.subTest(width=width), self.assertRaises(ValueError):
                saved.ValidateLayout(layout)
        layout = runtime()
        layout['columns'].append(dict(layout['columns'][0]))
        with self.assertRaises(ValueError):
            saved.ValidateLayout(layout)
        layout = runtime()
        layout['columns'][-1]['field'] = 'unknown'
        layout['sorters'] = []
        layout['groupFields'] = []
        layout['headerGroups'] = []
        with self.assertRaises(ValueError):
            saved.CaptureLayout(layout, snapshot())

    def test_legacy_single_group_normalizes_without_mutating_definition(self):
        layout = runtime()
        del layout['groupFields']
        layout['groupField'] = 'parameterValues.p201'
        before = json.dumps(layout, sort_keys=True)
        portable = saved.CaptureLayout(layout, snapshot())
        self.assertEqual(['shared:' + SHARED_GUID], portable['groupFields'])
        legacy = dict(portable)
        legacy['groupField'] = legacy.pop('groupFields')[0]
        restored, warnings = saved.ResolveLayout(legacy, snapshot())
        self.assertEqual(['parameterValues.p201'], restored['groupFields'])
        self.assertEqual([], warnings)
        self.assertEqual(before, json.dumps(layout, sort_keys=True))

    def test_invalid_group_hierarchies_are_rejected(self):
        for groups in ('level', ['level', 'level'], ['missing'], [None], [{}], list(saved.BASE_FIELDS) * 2):
            layout = runtime()
            layout['groupFields'] = groups
            with self.subTest(groups=groups), self.assertRaises(ValueError):
                saved.ValidateLayout(layout)

    def test_hiding_phase_removes_only_that_hierarchy_level(self):
        layout = runtime()
        layout['groupFields'] = ['level', 'phase', 'parameterValues.p201']
        portable = saved.CaptureLayout(layout, snapshot())
        restored, warnings = saved.ResolveLayout(portable, snapshot())
        self.assertEqual(['level', 'parameterValues.p201'], restored['groupFields'])
        self.assertEqual(1, len(warnings))

    def test_invalid_header_names_and_memberships_are_rejected(self):
        for headers in (None, {}, [{'name': '', 'fields': ['name']}],
                        [{'name': '   ', 'fields': ['name']}],
                        [{'name': 'x' * 101, 'fields': ['name']}],
                        [{'name': 'Identity', 'fields': []}],
                        [{'name': 'Identity', 'fields': ['name', 'name']}],
                        [{'name': 'Identity', 'fields': [{}]}],
                        [{'name': 'Identity', 'fields': ['missing']}],
                        [{'name': 'A', 'fields': ['name']}, {'name': 'B', 'fields': ['name']}]):
            layout = runtime()
            layout['headerGroups'] = headers
            with self.subTest(headers=headers), self.assertRaises(ValueError):
                saved.ValidateLayout(layout)

    def test_legacy_layouts_default_to_no_bands_and_hidden_members_remain_bound(self):
        layout = runtime()
        del layout['headerGroups']
        self.assertEqual([], saved.ValidateLayout(layout)['headerGroups'])
        layout['headerGroups'] = [{'name': 'Context', 'fields': ['level', 'phase']}]
        restored, warnings = saved.ResolveLayout(saved.CaptureLayout(layout, snapshot()), snapshot())
        self.assertEqual(layout['headerGroups'], restored['headerGroups'])
        self.assertEqual([], warnings)


class Generic:
    def __init__(self, function):
        self.function = function

    def __getitem__(self, value):
        return self.function


class Entity:
    def __init__(self, schema, text=None, valid=True):
        self.text = text
        self.valid = valid
        self.Get = Generic(lambda field: self.text)
        self.Set = Generic(lambda field, value: setattr(self, 'text', value))

    def IsValid(self):
        return self.valid


class DataStorage:
    def __init__(self, value):
        self.Id = NS(Value=value)
        self.text = None

    @staticmethod
    def Create(document):
        storage = DataStorage(len(document._host.elements) + 1)
        document._host.elements.append(storage)
        return storage

    def GetEntity(self, schema):
        return Entity(schema, self.text, self.text is not None)

    def SetEntity(self, entity):
        self.text = entity.text


class StoreHost:
    """Simulate transaction/checkout boundaries, leaving validation and writes real."""
    def __init__(self):
        self.elements = []
        self.schemas = {}
        self.transactions = []
        self.commitFails = False
        self.allowCheckout = True
        self.centralStatus = 'CurrentWithCentral'
        self.definitions = {}
        self.document = NS(_host=self, IsReadOnly=False, IsWorkshared=False,
                           Application=NS(Username='Tester'), GetElement=lambda key: self.definitions.get(key))
        self.document.Delete = lambda elementId: self.elements.remove(next(
            element for element in self.elements if element.Id.Value == elementId.Value))
        host = self

        class Builder:
            def __init__(self, guid):
                self.guid = guid

            def SetSchemaName(self, name): pass
            def SetReadAccessLevel(self, level): pass
            def SetWriteAccessLevel(self, level): pass
            def AddSimpleField(self, name, value): pass

            def Finish(self):
                schema = NS(GetField=lambda name: name)
                host.schemas[self.guid] = schema
                return schema

        class Transaction:
            def __init__(self, document, name):
                self.state = 'Uninitialized'
                self.name = name
                self.disposed = False
                self.rolledBack = False
                host.transactions.append(self)

            def Start(self):
                self.backup = [(element, element.text) for element in host.elements]
                self.state = 'Started'
                return self.state

            def Commit(self):
                if host.commitFails:
                    raise ValueError('commit failed')
                self.state = 'Committed'
                return self.state

            def GetStatus(self): return self.state

            def RollBack(self):
                host.elements[:] = [element for element, text in self.backup]
                for element, text in self.backup:
                    element.text = text
                self.rolledBack = True
                self.state = 'RolledBack'

            def Dispose(self): self.disposed = True

        self.checkout = mock.Mock(side_effect=lambda document, ids: ids if self.allowCheckout else [])
        self.db = NS(FilteredElementCollector=lambda document: NS(OfClass=lambda cls: list(self.elements)),
                     Transaction=Transaction, TransactionStatus=NS(Started='Started', Committed='Committed'),
                     WorksharingUtils=NS(CheckoutElements=self.checkout,
                                        GetModelUpdatesStatus=lambda document, elementId: self.centralStatus),
                     SharedParameterElement=NS(Lookup=lambda document, guid: self.definitions.get(guid)))
        self.sdk = {'Guid': lambda value: str(uuid.UUID(value)), 'String': str,
                    'Schema': NS(Lookup=lambda guid: self.schemas.get(guid)), 'SchemaBuilder': Builder,
                    'AccessLevel': NS(Public='Public'), 'DataStorage': DataStorage, 'Entity': Entity,
                    'CreateIds': list}
        self.store = saved.RevitTableStore(self.document, self.db, self.sdk)

    def create(self, name='Review'):
        return self.store.Write('saveTable', '', '', name, saved.CaptureLayout(runtime(), snapshot()))


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.host = StoreHost()

    def test_reading_empty_project_creates_no_storage_schema_or_transactions(self):
        self.assertEqual(([], []), self.host.store.List())
        self.assertEqual([], self.host.elements)
        self.assertEqual({}, self.host.schemas)
        self.assertEqual([], self.host.transactions)

    def test_legacy_storage_reads_without_writing_and_upgrades_on_explicit_save(self):
        record = self.host.create()
        element = self.host.elements[0]
        legacy = json.loads(element.text)
        legacy['version'] = 1
        legacy['layout']['groupField'] = legacy['layout'].pop('groupFields')[1]
        del legacy['layout']['headerGroups']
        element.text = json.dumps(legacy)
        before = element.text
        transactionCount = len(self.host.transactions)
        restored = self.host.store.Read(record['id'])[1]
        self.assertEqual(['shared:' + SHARED_GUID], restored['layout']['groupFields'])
        self.assertEqual([], restored['layout']['headerGroups'])
        self.assertEqual(before, element.text)
        self.assertEqual(transactionCount, len(self.host.transactions))
        updated = self.host.store.Write('saveTable', restored['id'], restored['revision'],
                                        restored['name'], restored['layout'])
        self.assertEqual(4, updated['version'])
        self.assertEqual(['shared:' + SHARED_GUID], updated['layout']['groupFields'])

    def test_save_reopen_rename_and_delete_are_separate_metadata_transactions(self):
        record = self.host.create()
        self.assertEqual(1, len(self.host.elements))
        self.assertEqual(record, self.host.store.Read(record['id'])[1])
        renamed = self.host.store.Write('renameTable', record['id'], record['revision'], 'New name')
        self.assertEqual(record['layout'], renamed['layout'])
        self.assertEqual('New name', renamed['name'])
        self.assertNotEqual(record['revision'], renamed['revision'])
        self.host.store.Write('deleteTable', record['id'], renamed['revision'])
        self.assertEqual([], self.host.elements)
        self.assertEqual(3, len(self.host.transactions))
        self.assertTrue(all(item.state == 'Committed' and item.disposed for item in self.host.transactions))

    def test_save_as_creates_independent_table_and_does_not_lock_another_table(self):
        first = self.host.create()
        second = self.host.create('Second')
        self.assertNotEqual(first['id'], second['id'])
        self.host.document.IsWorkshared = True
        self.host.store.Write('renameTable', first['id'], first['revision'], 'Renamed')
        self.assertEqual([1], [item.Value for item in self.host.checkout.call_args.args[1]])
        self.assertEqual(second, self.host.store.Read(second['id'])[1])

    def test_stale_revision_and_other_owner_never_begin_a_transaction(self):
        record = self.host.create()
        self.host.document.IsWorkshared = True
        before = self.host.elements[0].text
        for revision, allowed, status in ((str(uuid.uuid4()), True, 'CurrentWithCentral'),
                                          (record['revision'], False, 'CurrentWithCentral'),
                                          (record['revision'], True, 'UpdatedInCentral'),
                                          (record['revision'], True, 'DeletedInCentral')):
            with self.subTest(revision=revision, allowed=allowed, status=status):
                self.host.allowCheckout = allowed
                self.host.centralStatus = status
                with self.assertRaises(ValueError):
                    self.host.store.Write('renameTable', record['id'], revision, 'Unsafe')
                self.assertEqual(1, len(self.host.transactions))
                self.assertEqual(before, self.host.elements[0].text)

    def test_readonly_model_never_begins_storage_transaction(self):
        self.host.document.IsReadOnly = True
        with self.assertRaisesRegex(ValueError, 'read-only'):
            self.host.create()
        self.assertEqual([], self.host.transactions)

    def test_revision_is_checked_again_after_checkout(self):
        record = self.host.create()
        self.host.document.IsWorkshared = True

        def checkout(document, ids):
            changed = json.loads(self.host.elements[0].text)
            changed['revision'] = str(uuid.uuid4())
            self.host.elements[0].text = json.dumps(changed)
            return ids

        self.host.checkout.side_effect = checkout
        with self.assertRaisesRegex(ValueError, 'newer revision'):
            self.host.store.Write('renameTable', record['id'], record['revision'], 'Unsafe')
        self.assertEqual(1, len(self.host.transactions))
        self.assertEqual('Review', json.loads(self.host.elements[0].text)['name'])

    def test_save_as_remains_possible_when_original_is_owned_or_stale(self):
        first = self.host.create()
        self.host.document.IsWorkshared = True
        self.host.allowCheckout = False
        self.host.centralStatus = 'UpdatedInCentral'
        second = self.host.create('My working copy')
        self.assertNotEqual(first['id'], second['id'])
        self.host.checkout.assert_not_called()
        self.assertEqual(2, len(self.host.elements))

    def test_readonly_model_can_read_existing_definitions(self):
        record = self.host.create()
        self.host.document.IsReadOnly = True
        self.assertEqual(record, self.host.store.Read(record['id'])[1])
        self.assertEqual(1, len(self.host.transactions))

    def test_failed_commit_rolls_back_creation_update_and_delete(self):
        self.host.commitFails = True
        with self.assertRaisesRegex(ValueError, 'commit failed'):
            self.host.create()
        self.assertEqual([], self.host.elements)
        self.assertTrue(self.host.transactions[-1].rolledBack)
        self.host.commitFails = False
        record = self.host.create()
        before = self.host.elements[0].text
        self.host.commitFails = True
        for action in ('renameTable', 'deleteTable'):
            with self.assertRaises(ValueError):
                self.host.store.Write(action, record['id'], record['revision'], 'Changed')
            self.assertEqual(before, self.host.elements[0].text)
            self.assertTrue(self.host.transactions[-1].rolledBack)
            self.assertTrue(self.host.transactions[-1].disposed)

    def test_future_or_corrupt_record_is_reported_without_replacing_it(self):
        record = self.host.create()
        unsupported = dict(record, id=str(uuid.uuid4()), version=5)
        element = DataStorage.Create(self.host.document)
        element.text = json.dumps(unsupported)
        records, issues = self.host.store.List()
        self.assertEqual(1, len(records))
        self.assertIn('Unsupported', issues[0])
        self.assertEqual(unsupported, json.loads(element.text))
        with self.assertRaises(ValueError):
            self.host.store.Write('deleteTable', unsupported['id'], unsupported['revision'])
        self.assertEqual(2, len(self.host.elements))

    def test_identity_resolution_uses_guid_and_uniqueid_instead_of_old_parameter_ids(self):
        self.host.definitions[SHARED_GUID] = NS(Id=NS(Value=999))
        self.host.definitions['project-uid'] = NS(Id=NS(Value=1000))
        layout = {'columns': [{'field': 'shared:' + SHARED_GUID}, {'field': 'project:project-uid'},
                              {'field': 'builtin:-123'}, {'field': 'project:missing'}]}
        self.assertEqual(['999', '1000', '-123'], self.host.store.ResolveParameterIds(layout))


if __name__ == '__main__':
    unittest.main()
