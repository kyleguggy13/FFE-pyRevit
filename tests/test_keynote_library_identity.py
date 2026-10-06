"""Identity lookup and Revit association transactions, without loading Revit."""
import copy
import types
import unittest
from test_keynote_annotation_library import load_functions, row

LIBRARY_ID = '11111111-1111-4111-8111-111111111111'
OTHER_ID = '22222222-2222-4222-8222-222222222222'
FIELDS = {'libraryId': 'LibraryId', 'projectUrl': 'ProjectUrl', 'libraryKey': 'LibraryKey',
          'storageMode': 'StorageMode', 'documentKey': 'DocumentKey', 'keynotePath': 'KeynotePath',
          'annotationKey': 'AnnotationKey'}


class LibraryIdentityTests(unittest.TestCase):
    def setUp(self):
        self.api = load_functions(with_associations=True)
        self.events = []
        self.rpc_calls = []
        self.schema = None
        self.fail_commit = False
        self.doc = types.SimpleNamespace(storages=[], document_key='central.rvt', path='old.txt')
        self.settings = {'configured': True, 'url': 'https://project.supabase.co/', 'anonKey': 'never-store-this'}
        outer = self
        class GenericCall:
            def __init__(self, function): self.function = function
            def __getitem__(self, type_arg): return self.function
        class Entity:
            def __init__(self, schema, valid=True):
                self.data = {}
                self.valid = valid
                self.Get = GenericCall(lambda field: self.data.get(field, ''))
                self.Set = GenericCall(lambda field, value: self.data.__setitem__(field, value))
            def IsValid(self): return self.valid
        class Storage:
            fail_write = False
            def __init__(self): self.entity = Entity(None, False)
            def GetEntity(self, schema): return self.entity
            def SetEntity(self, entity):
                if self.fail_write: raise Exception('Storage is owned by another user')
                self.entity = entity
        def create_storage(doc):
            storage = Storage()
            doc.storages.append(storage)
            return storage
        class Builder:
            def __init__(self, guid): self.fields = []
            def SetSchemaName(self, name): outer.schema_name = name
            def SetReadAccessLevel(self, level): pass
            def SetWriteAccessLevel(self, level): pass
            def AddSimpleField(self, name, field_type): self.fields.append(name)
            def Finish(self):
                outer.schema = types.SimpleNamespace(GetField=lambda field: field)
                outer.schema_fields = self.fields
                return outer.schema
        class Collector:
            def __init__(self, doc): self.doc = doc
            def OfClass(self, cls): return self
            def ToElements(self): return self.doc.storages
        class Transaction:
            def __init__(self, doc, name): self.doc, self.status = doc, None
            def Start(self):
                self.before = (list(self.doc.storages), [storage.entity for storage in self.doc.storages])
                self.status = 'started'
                outer.events.append('start')
            def Commit(self):
                if outer.fail_commit: return 'failed'
                self.status = 'committed'
                outer.events.append('commit')
                return self.status
            def GetStatus(self): return self.status
            def RollBack(self):
                self.doc.storages[:] = self.before[0]
                for storage, entity in zip(self.doc.storages, self.before[1]): storage.entity = entity
                self.status = 'rolled-back'
                outer.events.append('rollback')
        class Group(Transaction):
            def Start(self):
                super().Start()
                self.path_before = self.doc.path
            def Assimilate(self): return self.Commit()
            def RollBack(self):
                super().RollBack()
                self.doc.path = self.path_before
        self.api.update({'LIBRARY_ASSOCIATION_SCHEMA_GUID': 'test-schema-guid', 'LIBRARY_ASSOCIATION_FIELDS': FIELDS,
            'Schema': types.SimpleNamespace(Lookup=lambda guid: self.schema), 'SchemaBuilder': Builder,
            'Guid': lambda value: value, 'String': str, 'Entity': Entity,
            'DataStorage': types.SimpleNamespace(Create=create_storage), 'FilteredElementCollector': Collector,
            'AccessLevel': types.SimpleNamespace(Public='public'), 'Transaction': Transaction, 'TransactionGroup': Group,
            'TransactionStatus': types.SimpleNamespace(Started='started', Committed='committed'),
            'load_supabase_settings': lambda: self.settings,
            'normalize_path': lambda path: path.replace('\\', '/').lower(),
            'get_document_analytics_identity': lambda doc: {'documentKey': doc.document_key, 'documentKeySource': 'centralPath'},
            'get_document_title': lambda doc: 'Test Project',
            'read_user_settings': lambda: {}, 'save_storage_mode': lambda doc, mode: self.events.append('mode:' + mode),
            'build_base_payload': lambda doc, status, message: {'status': status, 'message': message, 'preferences': {}, 'issues': []}})
        self.snapshot = {'status': 'ready', 'libraryId': LIBRARY_ID, 'libraryKey': 'old.txt',
                         'sourceType': 'file', 'datasetVersion': 7, 'fileHash': 'hash', 'entries': [row('A', 'Original')]}
        def rpc(name, arguments, **kwargs):
            self.rpc_calls.append((name, arguments))
            return copy.deepcopy(self.snapshot)
        self.api['keynote_supabase_rpc'] = rpc
        self.record = {'libraryId': LIBRARY_ID, 'projectUrl': 'https://project.supabase.co', 'libraryKey': 'old.txt',
                       'storageMode': 'file', 'documentKey': 'central.rvt', 'keynotePath': 'old.txt',
                       'annotationKey': 'annotation:' + LIBRARY_ID}

    def bind(self):
        self.api['write_library_association'](self.doc, self.record)
        self.events.clear()

    def file_payload(self, path='new.txt'):
        return {'status': 'ready', 'storageMode': 'file', 'libraryKey': path, 'fileLibraryKey': path,
                'keynotePath': path, 'fileHash': 'hash', 'supabase': self.settings, 'issues': []}

    def test_schema_and_binding_store_identity_only_and_repeated_reads_writes_are_noops(self):
        self.bind()
        self.assertEqual(sorted(FIELDS.values()), self.schema_fields)
        self.assertEqual(self.record, self.api['read_library_association'](self.doc))
        self.assertFalse(self.api['write_library_association'](self.doc, self.record))
        self.assertEqual([], self.events)
        self.assertNotIn('never-store-this', repr(self.doc.storages[0].entity.data))

    def test_first_successful_path_lookup_adopts_existing_uuid(self):
        payload = self.file_payload('old.txt')
        self.api['resolve_file_library_association'](self.doc, payload)
        self.assertEqual('get_keynote_snapshot', self.rpc_calls[0][0])
        self.assertEqual(LIBRARY_ID, self.api['read_library_association'](self.doc)['libraryId'])
        self.assertEqual('libraryAssociationSaved', payload['issues'][0]['code'])

    def test_file_move_resolves_by_uuid_and_links_new_path_without_creating_a_library(self):
        self.bind()
        payload = self.file_payload()
        self.api['resolve_file_library_association'](self.doc, payload)
        self.assertEqual(['get_keynote_snapshot_by_id', 'link_keynote_library_path'], [call[0] for call in self.rpc_calls])
        self.assertEqual(LIBRARY_ID, payload['libraryId'])
        self.assertEqual('old.txt', payload['libraryKey'])
        self.assertEqual('new.txt', self.api['read_library_association'](self.doc)['keynotePath'])

    def test_model_using_an_older_file_must_reconnect_without_overwriting_the_active_file(self):
        self.bind()
        self.snapshot.update({'fileLibraryKey': 'new.txt', 'displayPath': 'New.txt'})
        with self.assertRaisesRegex(Exception, 'Reconnect Keynote File'):
            self.api['resolve_file_library_association'](self.doc, self.file_payload('old.txt'))
        self.assertEqual(['get_keynote_snapshot_by_id'], [call[0] for call in self.rpc_calls])
        self.assertEqual([], self.events)

    def test_corrupt_or_conflicting_metadata_can_only_be_replaced_explicitly(self):
        self.bind()
        storage = self.doc.storages[0]
        storage.entity.data['LibraryId'] = 'corrupt'
        with self.assertRaises(Exception): self.api['read_library_association'](self.doc)
        with self.assertRaises(Exception): self.api['write_library_association'](self.doc, self.record)
        self.api['write_library_association'](self.doc, self.record, replace_existing=True)
        self.assertEqual(self.record, self.api['read_library_association'](self.doc))

    def test_failed_first_write_removes_the_new_storage_element(self):
        self.fail_commit = True
        with self.assertRaisesRegex(Exception, 'did not commit'):
            self.api['write_library_association'](self.doc, self.record)
        self.assertEqual([], self.doc.storages)
        self.assertIsNone(self.api['read_library_association'](self.doc))

    def test_annotation_load_uses_uuid_even_when_canonical_key_is_a_file_path(self):
        self.bind()
        self.snapshot['sourceType'] = 'annotation'
        payload = self.api['build_annotation_keynote_payload'](self.doc, include_model_health=False)
        self.assertEqual('ready', payload['status'])
        self.assertEqual(['get_keynote_snapshot_by_id'], [call[0] for call in self.rpc_calls])
        self.assertEqual('annotation', self.api['read_library_association'](self.doc)['storageMode'])

    def test_missing_associated_uuid_never_falls_back_to_path_lookup_or_creation(self):
        self.bind()
        self.snapshot = {'status': 'error', 'message': 'Associated library missing'}
        with self.assertRaisesRegex(Exception, 'Associated library missing'):
            self.api['resolve_file_library_association'](self.doc, self.file_payload())
        self.assertEqual(['get_keynote_snapshot_by_id'], [call[0] for call in self.rpc_calls])
        self.assertEqual([], self.events)

    def test_wrong_supabase_project_is_blocked_before_any_lookup(self):
        self.bind()
        self.settings['url'] = 'https://other.supabase.co'
        with self.assertRaisesRegex(Exception, 'different Supabase project'):
            self.api['resolve_file_library_association'](self.doc, self.file_payload())
        self.assertEqual([], self.rpc_calls)

    def test_offline_file_load_retains_cached_uuid_without_new_path_fallback(self):
        self.bind()
        def offline(name, arguments, **kwargs):
            self.rpc_calls.append((name, arguments))
            raise Exception('Could not access the Supabase keynote library. Offline')
        self.api['keynote_supabase_rpc'] = offline
        payload = self.file_payload()
        self.api['resolve_file_library_association'](self.doc, payload)
        self.assertEqual(LIBRARY_ID, payload['libraryId'])
        self.assertEqual('old.txt', payload['libraryKey'])
        self.assertEqual('libraryAssociationLookupFailed', payload['issues'][0]['code'])
        self.assertEqual([], self.events)

    def test_model_rename_can_keep_association_and_cancellation_changes_nothing(self):
        self.bind()
        self.doc.document_key = 'renamed-central.rvt'
        self.api['forms'] = types.SimpleNamespace(CommandSwitchWindow=types.SimpleNamespace(show=lambda *args, **kwargs: None))
        with self.assertRaisesRegex(Exception, 'canceled'):
            self.api['reconcile_model_library_association'](self.doc)
        self.assertEqual([], self.events)
        self.api['forms'].CommandSwitchWindow.show = lambda *args, **kwargs: 'Keep existing library'
        self.api['reconcile_model_library_association'](self.doc)
        record = self.api['read_library_association'](self.doc)
        self.assertEqual(LIBRARY_ID, record['libraryId'])
        self.assertEqual('renamed-central.rvt', record['documentKey'])
        self.assertEqual('annotation:' + LIBRARY_ID, self.api['annotation_library_key'](self.doc))
        self.events.clear()
        self.api['reconcile_model_library_association'](self.doc)
        self.assertEqual([], self.events)

    def test_failed_storage_writes_and_commit_roll_back_and_remain_visible(self):
        for failure in ('ownership', 'commit'):
            with self.subTest(failure=failure):
                self.setUp()
                self.bind()
                if failure == 'ownership': self.doc.storages[0].fail_write = True
                else: self.fail_commit = True
                payload = self.file_payload()
                self.api['remember_library_association'](self.doc, payload, self.snapshot)
                self.assertEqual(self.record, self.api['read_library_association'](self.doc))
                self.assertEqual('rollback', self.events[-1])
                self.assertEqual('libraryAssociationWriteFailed', payload['issues'][0]['code'])

    def test_unsaved_copy_is_blocked_and_same_central_locals_do_not_prompt(self):
        self.bind()
        self.api['forms'] = types.SimpleNamespace(CommandSwitchWindow=types.SimpleNamespace(
            show=lambda *args, **kwargs: self.fail('Same central identity prompted')))
        self.api['reconcile_model_library_association'](self.doc)
        self.api['get_document_analytics_identity'] = lambda doc: {'documentKey': 'title:Copy', 'documentKeySource': 'title'}
        with self.assertRaisesRegex(Exception, 'Save this project'):
            self.api['reconcile_model_library_association'](self.doc)

    def configure_fork(self, mode='file', cancel=False, failure=None):
        self.bind()
        self.snapshot['sourceType'] = mode
        self.api['forms'] = types.SimpleNamespace(save_file=lambda **kwargs: None if cancel else 'copy.txt')
        def raw_load(doc, **kwargs):
            self.assertFalse(kwargs['resolve_association'])
            path = kwargs.get('source_path') or doc.path
            return {'status': 'ready', 'storageMode': 'file', 'libraryKey': path, 'displayPath': path,
                    'keynotePath': path, 'fileHash': 'copy-hash', 'entries': [row('A', 'Latest file text')]}
        self.api['build_keynote_payload'] = raw_load
        def assign(doc, path, entries): doc.path = path
        self.api['write_and_assign_keynote_file'] = assign
        def rpc(name, arguments, **kwargs):
            self.rpc_calls.append((name, arguments))
            if name == 'get_keynote_snapshot_by_id': return copy.deepcopy(self.snapshot)
            self.assertEqual('fork_keynote_library', name)
            if failure == 'rpc': raise Exception('Concurrent library edit')
            if failure == 'binding': self.doc.storages[0].fail_write = True
            return dict(self.snapshot, libraryId=OTHER_ID, libraryKey=arguments['p_library_key'])
        self.api['keynote_supabase_rpc'] = rpc

    def test_independent_file_library_uses_new_file_and_uuid_and_latest_file_contents(self):
        self.configure_fork()
        result = self.api['fork_model_keynote_library'](self.doc, self.record)
        self.assertEqual(OTHER_ID, result['libraryId'])
        self.assertEqual('copy.txt', self.doc.path)
        self.assertEqual(OTHER_ID, self.api['read_library_association'](self.doc)['libraryId'])
        request = self.rpc_calls[-1][1]
        self.assertEqual('Latest file text', request['p_seed_entries'][0]['text'])
        self.assertEqual(7, request['p_base_dataset_version'])
        self.assertEqual('mode:file', self.events[-1])

    def test_independent_annotation_library_copies_saved_cloud_notes_without_file_dialog(self):
        self.configure_fork(mode='annotation')
        self.api['forms'].save_file = lambda **kwargs: self.fail('Annotation fork opened a file dialog')
        result = self.api['fork_model_keynote_library'](self.doc, self.record)
        self.assertEqual(OTHER_ID, result['libraryId'])
        self.assertTrue(result['libraryKey'].startswith('annotation:'))
        self.assertIsNone(self.rpc_calls[-1][1]['p_seed_entries'])
        self.assertEqual('old.txt', self.doc.path)

    def test_fork_cancel_or_failure_keeps_original_association_and_assignment(self):
        for failure in ('cancel', 'rpc', 'binding'):
            with self.subTest(failure=failure):
                self.setUp()
                self.configure_fork(cancel=failure == 'cancel', failure=failure)
                if failure == 'cancel':
                    self.assertIsNone(self.api['fork_model_keynote_library'](self.doc, self.record))
                else:
                    with self.assertRaises(Exception): self.api['fork_model_keynote_library'](self.doc, self.record)
                self.assertEqual(self.record, self.api['read_library_association'](self.doc))
                self.assertEqual('old.txt', self.doc.path)
                self.assertNotIn('mode:file', self.events)

    def configure_library_choice(self, cancel=False, fail_write=False):
        self.bind()
        selected = dict(self.snapshot, libraryId=OTHER_ID, libraryKey='selected.txt', displayPath='Selected.txt')
        self.api['forms'] = types.SimpleNamespace(SelectFromList=types.SimpleNamespace(
            show=lambda options, **kwargs: None if cancel else options[0]))
        self.api['os'] = types.SimpleNamespace(path=types.SimpleNamespace(exists=lambda path: True))
        def rpc(name, arguments, **kwargs):
            self.rpc_calls.append((name, arguments))
            if name == 'list_keynote_libraries': return {'status': 'ready', 'libraries': [selected]}
            return copy.deepcopy(selected)
        self.api['keynote_supabase_rpc'] = rpc
        self.api['assign_keynote_file'] = lambda doc, path: setattr(doc, 'path', path)
        self.api['build_keynote_payload'] = lambda doc, **kwargs: dict(self.file_payload(doc.path), entries=[row('A')])
        self.doc.storages[0].fail_write = fail_write

    def test_change_library_repairs_corrupt_metadata_only_after_verifying_selected_library(self):
        self.configure_library_choice()
        self.doc.storages[0].entity.data['LibraryId'] = 'corrupt'
        result = self.api['change_model_library_association'](self.doc, 'choose')
        self.assertEqual('ready', result['status'])
        self.assertEqual(OTHER_ID, self.api['read_library_association'](self.doc)['libraryId'])
        self.assertEqual('Selected.txt', self.doc.path)
        self.assertEqual(['list_keynote_libraries', 'get_keynote_snapshot_by_id', 'link_keynote_library_path'],
                         [name for name, arguments in self.rpc_calls])
        self.assertIn('Save/Sync', result['message'])

    def test_change_library_cancel_or_failed_association_write_restores_model(self):
        for cancel in (True, False):
            with self.subTest(cancel=cancel):
                self.setUp()
                self.configure_library_choice(cancel=cancel, fail_write=not cancel)
                result = self.api['change_model_library_association'](self.doc, 'choose')
                self.assertEqual('canceled' if cancel else 'error', result['status'])
                self.assertEqual(self.record, self.api['read_library_association'](self.doc))
                self.assertEqual('old.txt', self.doc.path)
                self.assertNotIn('mode:file', self.events)


if __name__ == '__main__':
    unittest.main()
