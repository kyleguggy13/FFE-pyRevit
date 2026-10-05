"""Behavior tests for keynote storage without loading Revit or starting the UI.

The actual pure functions and storage workflow are loaded from script.py via AST.
Revit transactions are replaced only at the boundary to verify atomic rollback.
Run: python -m unittest discover -s tests -p 'test_keynote_annotation_library.py'
"""
import ast
import codecs
import copy
import ntpath
import pathlib
import types
import unittest
import uuid


ROOT = pathlib.Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'FFE-pyRevit.extension/FFE-pyRevit.tab/DesignTools.panel/Keynotes.pushbutton/script.py'
TREE = ast.parse(SCRIPT.read_text(encoding='utf-8'))


def load_functions():
    namespace = {'unicode': str, 'uuid': uuid, 'codecs': codecs}
    definitions = [node for node in TREE.body if isinstance(node, ast.FunctionDef)]
    exec(compile(ast.Module(body=definitions, type_ignores=[]), str(SCRIPT), 'exec'), namespace)
    return namespace


def row(key, description='', parent='', identity=None):
    return {'id': identity or key, 'key': key, 'text': description, 'parentKey': parent}


class KeynoteStorageTests(unittest.TestCase):
    def setUp(self):
        self.api = load_functions()
        self.api['get_storage_mode'] = lambda doc: 'file'
        self.api['get_document_analytics_identity'] = lambda doc: {
            'documentKey': 'central.rvt', 'documentKeySource': 'centralPath'}

    def configure_conversion(self, canceled=False, failure=None):
        events = []
        rows = [row('DIVISION 22', 'PLUMBING'), row('22.01', '', 'DIVISION 22'),
                row('22.01A', 'Caf\u00e9 fixture', '22.01'), row('22.02', 'Unplaced', 'DIVISION 22')]
        payload = {'status': 'ready', 'message': 'Loaded file', 'storageMode': 'file',
                   'libraryKey': 'new-notes.txt', 'keynotePath': 'new-notes.txt',
                   'entries': rows, 'encoding': 'utf-16', 'lineEnding': '\r\n',
                   'fileHash': 'export-hash', 'lastWriteUtc': 123, 'supabase': {'clientId': 'user'}}
        class Forms:
            @staticmethod
            def save_file(**kwargs):
                events.append('dialog')
                return None if canceled else 'new-notes.txt'
        def read_cloud(*args, **kwargs):
            events.append('read-cloud')
            if failure == 'read': raise Exception('Supabase unavailable')
            return {'status': 'ready', 'entries': rows, 'libraryId': 'same-library-id', 'datasetVersion': 7}
        def assign(doc, path, entries):
            events.append('assign')
            self.assertEqual(rows, entries)
            if failure == 'assign': raise Exception('Assignment failed')
            content = self.api['canonicalize_entries'](entries, '\r\n', path, 'utf-16')
            parsed, issues = self.api['parse_keynote_text'](content)
            self.assertEqual([], issues)
            self.assertEqual([(r['key'], r['text'], r['parentKey']) for r in rows],
                             [(r['key'], r['text'], r['parentKey']) for r in parsed])
        def sync(name, args):
            events.append('sync')
            self.assertEqual('convert_annotation_keynote_library_to_file', name)
            self.assertEqual('annotation:project', args['p_annotation_key'])
            self.assertEqual(7, args['p_base_dataset_version'])
            self.assertEqual('new-notes.txt', args['p_file_key'])
            self.assertNotIn('p_entries', args)
            self.assertEqual('export-hash', args['p_file_hash'])
            self.assertEqual(123, args['p_last_write_utc'])
            self.assertEqual('utf-16', args['p_encoding'])
            if failure == 'sync': raise Exception('Connection interrupted')
            return {'status': 'ready', 'libraryId': 'same-library-id'}
        class Group:
            def __init__(self, *args): pass
            def Start(self): events.append('start')
            def Assimilate(self):
                events.append('commit')
                return 'committed'
            def GetStatus(self): return 'started'
            def RollBack(self): events.append('rollback')
        class Status:
            Started = 'started'
            Committed = 'committed'
        self.api.update({'forms': Forms, 'get_storage_mode': lambda doc: 'annotation',
                         'annotation_library_key': lambda doc: 'annotation:project',
                         'TransactionGroup': Group, 'TransactionStatus': Status,
                         'template_entries': lambda request: self.fail('Conversion requested the division template'),
                         'build_annotation_keynote_payload': read_cloud, 'write_and_assign_keynote_file': assign,
                         'save_storage_mode': lambda doc, mode: events.append('mode:' + mode),
                         'build_keynote_payload': lambda *args, **kwargs: copy.deepcopy(payload),
                         'keynote_supabase_rpc': sync})
        return events

    def test_annotation_to_file_exports_all_rows_and_syncs_new_file_metadata(self):
        for create_file in (False, True):
            events = self.configure_conversion()
            result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file', 'createFile': create_file})
            self.assertEqual('ready', result['status'])
            self.assertEqual(['dialog', 'read-cloud', 'start', 'assign', 'sync', 'commit', 'mode:file'], events)

    def test_canceled_annotation_export_keeps_source_and_does_not_write(self):
        events = self.configure_conversion(canceled=True)
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file'})
        self.assertEqual('canceled', result['status'])
        self.assertEqual(['dialog'], events)

    def test_annotation_export_read_or_assignment_failure_keeps_mode_and_database(self):
        for failure in ('read', 'assign'):
            events = self.configure_conversion(failure=failure)
            result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file'})
            self.assertEqual('error', result['status'])
            self.assertNotIn('mode:file', events)
            self.assertNotIn('sync', events)

    def test_annotation_export_rolls_back_assignment_when_conversion_is_unconfirmed(self):
        events = self.configure_conversion(failure='sync')
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file'})
        self.assertEqual('error', result['status'])
        self.assertNotIn('payload', result)
        self.assertIn('assignment was rolled back', result['message'])
        self.assertEqual(['dialog', 'read-cloud', 'start', 'assign', 'sync', 'rollback'], events)

    def test_manual_file_recovery_exports_existing_cloud_notes_even_with_a_file_preference(self):
        events = self.configure_conversion()
        convert_rpc = self.api['keynote_supabase_rpc']
        def rpc(name, arguments, allow_missing=False):
            if name == 'get_annotation_keynote_snapshot':
                events.append('lookup')
                return {'status': 'ready', 'libraryId': 'same-library-id'}
            return convert_rpc(name, arguments)
        self.api['get_storage_mode'] = lambda doc: 'file'
        self.api['keynote_supabase_rpc'] = rpc
        self.api['build_project_keynote_payload'] = lambda *args, **kwargs: self.fail('Manual recovery used automatic detection')
        result = self.api['setup_keynote_storage'](object(), {
            'storageMode': 'file', 'createFile': True, 'recoverySetup': True})
        self.assertEqual('ready', result['status'])
        self.assertEqual(['lookup', 'dialog', 'read-cloud', 'start', 'assign', 'sync', 'commit', 'mode:file'], events)

    def test_template_has_exact_divisions_and_no_notes(self):
        migration = (ROOT / 'supabase/keynote_annotation_library.sql').read_text(encoding='utf-8')
        content = migration.split('$template$')[1]
        rows, issues = self.api['parse_keynote_text'](content)
        self.assertEqual([], issues)
        self.assertEqual(21, len(rows))
        self.assertTrue(all(not entry['parentKey'] for entry in rows))
        self.assertEqual(['DIVISION ' + number for number in
                          ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11',
                           '12', '13', '14', '22', '23', '26', '28', '31', '32', '33']],
                         [entry['key'] for entry in rows])

    def test_text_roundtrip_preserves_hierarchy_unicode_and_blank_descriptions(self):
        entries = [row('DIVISION 22', 'PLUMBING'), row('22.01', '', 'DIVISION 22'),
                   row('22.01A', 'Caf\u00e9 fixture', '22.01')]
        content = self.api['canonicalize_entries'](entries, '\r\n', 'notes.txt', 'utf-16')
        encoded = self.api['encode_keynote_text'](content, 'utf-16')
        self.assertTrue(encoded.startswith(codecs.BOM_UTF16_LE))
        actual, issues = self.api['parse_keynote_text'](encoded.decode('utf-16'))
        self.assertEqual([], issues)
        self.assertFalse(self.api['has_error_issues'](self.api['validate_entries'](actual, issues)))
        self.assertEqual([(r['key'], r['text'], r['parentKey']) for r in entries],
                         [(r['key'], r['text'], r['parentKey']) for r in actual])

    def test_three_way_merge_rejects_stale_edits(self):
        baseline = [row('A', 'Original')]
        actual, issues = self.api['merge_keynote_entries'](
            [row('A', 'Other user')], baseline, [row('A', 'My edit')])
        self.assertIsNone(actual)
        self.assertEqual('rowConflict', issues[0]['code'])

    def test_invalid_hierarchy_blocks_save(self):
        for rows in ([row('A'), row('A')], [row('A', parent='missing')],
                     [row('A', parent='B'), row('B', parent='A')]):
            self.assertTrue(self.api['has_error_issues'](self.api['validate_entries'](rows)))

    def test_canceled_file_dialog_creates_nothing_and_keeps_source(self):
        class Forms:
            @staticmethod
            def save_file(**kwargs):
                return None
        self.api['forms'] = Forms
        self.api['save_storage_mode'] = lambda *args: self.fail('Cancellation changed the source')
        result = self.api['setup_keynote_storage'](object(), {
            'storageMode': 'file', 'createFile': True, 'template': {'content': 'A\tAlpha\n'}})
        self.assertEqual('canceled', result['status'])

    def test_both_setup_modes_require_a_saved_project_before_any_side_effect(self):
        self.api['get_document_analytics_identity'] = lambda doc: {'documentKeySource': 'title'}
        self.api['keynote_supabase_rpc'] = lambda *args: self.fail('Unsaved setup accessed Supabase')
        self.api['template_entries'] = lambda *args: self.fail('Unsaved setup requested a template')
        self.api['save_storage_mode'] = lambda *args: self.fail('Unsaved setup changed storage mode')
        for mode in ('file', 'annotation'):
            result = self.api['setup_keynote_storage'](object(), {'storageMode': mode, 'createFile': True})
            self.assertEqual('error', result['status'])
            self.assertIn('Save the Revit project', result['message'])

    def test_setup_does_not_persist_mode_when_resulting_source_fails_to_load(self):
        self.api['build_keynote_payload'] = lambda *args, **kwargs: {'status': 'error', 'message': 'Load failed'}
        self.api['save_storage_mode'] = lambda *args: self.fail('Failed loading persisted mode')
        self.api['get_document_title'] = lambda doc: 'Project'
        self.api['keynote_supabase_rpc'] = lambda *args: {'status': 'ready'}
        for mode in ('file', 'annotation'):
            result = self.api['setup_keynote_storage'](object(), {'storageMode': mode})
            self.assertEqual('error', result['status'])
            self.assertEqual('Load failed', result['message'])
            self.assertNotIn('payload', result)

    def test_unmatched_base_family_does_not_require_arrowhead_for_library_save(self):
        self.api['get_generic_annotation_keynote_family'] = lambda doc: 'family'
        self.api['get_family_symbols'] = lambda doc, family: ['base']
        self.api['get_generic_annotation_symbol_key'] = lambda symbol: 'Default'
        self.api['get_element_name'] = lambda symbol: 'Default'
        self.api['get_leader_arrowhead_type'] = lambda doc: self.fail('Unrelated base type requested an arrowhead')
        summary = self.api['sync_generic_annotation_types'](object(), [row('A')], {}, [], skip_unmatched=True)
        self.assertEqual(0, summary['failedCount'])

    def test_annotation_analytics_excludes_native_keynote_tags(self):
        class Collector:
            def __init__(self, doc): pass
            def OfCategory(self, category): return self
            def WhereElementIsNotElementType(self): return self
            def __iter__(self): return iter(['native-tag'])

        class Category:
            OST_KeynoteTags = 1

        scanned = []
        self.api.update({
            'FilteredElementCollector': Collector, 'BuiltInCategory': Category,
            'build_view_sheet_lookup': lambda doc: {},
            'iter_generic_annotation_keynote_instances': lambda doc: [('annotation', 'type')],
            'get_generic_annotation_symbol_key': lambda symbol: 'A',
            'get_keynote_tag_key': lambda tag: self.fail('Model scan included native tag'),
            'record_keynote_analytics_placement': lambda rows, entries, key, source, *args: scanned.append(source) or True,
            'finalize_keynote_analytics_rows': lambda rows: [],
            'summarize_keynote_analytics_rows': lambda rows: {},
            'get_document_analytics_identity': lambda doc: {},
            'get_generated_at': lambda: 'test'
        })
        result = self.api['collect_keynote_analytics'](object(), {'storageMode': 'annotation', 'entries': []})
        self.assertEqual(['genericAnnotation'], scanned)
        self.assertEqual(0, result['userKeynoteScannedCount'])
        self.assertEqual(1, result['genericAnnotationScannedCount'])


    def test_no_rvt_library_storage_code_remains(self):
        source = SCRIPT.read_text(encoding='utf-8')
        for forbidden in ('DataStorage', 'ExtensibleStorage', 'storageRevision', 'storageAncestors'):
            self.assertNotIn(forbidden, source)

    def test_connection_popup_saves_both_fields_and_preserves_client_identity(self):
        saved = {'clientId': 'existing-user', 'otherSetting': True}
        self.api.update({
            'get_supabase_settings_path': lambda: 'settings.json',
            'read_json_file': lambda path: copy.deepcopy(saved),
            'write_json_file': lambda path, value: saved.update(value),
            'get_client_name': lambda: 'Tester'
        })
        self.api['save_supabase_settings']({'url': ' https://project.supabase.co/ ', 'anonKey': ' sb_publishable_example '})
        self.assertEqual('https://project.supabase.co', saved['url'])
        self.assertEqual('sb_publishable_example', saved['anonKey'])
        self.assertEqual('existing-user', saved['clientId'])
        self.assertTrue(saved['otherSetting'])

    def test_connection_popup_rejects_invalid_values_before_writing(self):
        self.api['write_json_file'] = lambda *args: self.fail('Invalid connection was saved')
        for values in ({'url': 'not-a-url', 'anonKey': 'key'},
                       {'url': 'https://project.supabase.co', 'anonKey': ''},
                       {'url': 'https://project.supabase.co', 'anonKey': 'two words'}):
            with self.assertRaisesRegex(Exception, 'project URL|publishable key'):
                self.api['save_supabase_settings'](values)

    def test_connection_popup_reports_failed_persistence(self):
        self.api.update({
            'get_supabase_settings_path': lambda: 'settings.json',
            'read_json_file': lambda path: {},
            'write_json_file': lambda *args: None,
            'get_client_name': lambda: 'Tester'
        })
        with self.assertRaisesRegex(Exception, 'Could not save Supabase settings'):
            self.api['save_supabase_settings']({'url': 'https://project.supabase.co', 'anonKey': 'example'})

    def test_missing_connection_loads_without_opening_a_native_form(self):
        self.api.update({
            'get_supabase_settings_path': lambda: 'settings.json',
            'read_json_file': lambda path: {},
            'load_shared_supabase_settings': lambda: {},
            'get_client_name': lambda: 'Tester'
        })
        self.assertFalse(self.api['load_supabase_settings']()['configured'])

    def test_library_identity_uses_shared_document_path_and_rejects_unsaved_models(self):
        self.api['get_document_analytics_identity'] = lambda doc: {'documentKey': 'central.rvt', 'documentKeySource': 'centralPath'}
        self.assertEqual('annotation:central.rvt', self.api['annotation_library_key'](object()))
        self.api['get_document_analytics_identity'] = lambda doc: {'documentKey': 'title:project', 'documentKeySource': 'title'}
        with self.assertRaisesRegex(Exception, 'Save the Revit project'):
            self.api['annotation_library_key'](object())

    def test_file_save_bridge_cannot_save_annotation_library(self):
        self.api['get_storage_mode'] = lambda doc: 'annotation'
        result = self.api['save_keynote_payload'](object(), {'storageMode': 'annotation'})
        self.assertEqual('error', result['status'])
        self.assertIn('directly to Supabase', result['message'])

    def test_placement_rejects_stale_supabase_snapshot_before_touching_revit(self):
        result = self.api['place_generic_annotation_keynote'](None, None,
            {'libraryKey': 'annotation:x', 'datasetVersion': 1, 'key': 'A'},
            {'storageMode': 'annotation', 'status': 'ready', 'libraryKey': 'annotation:x', 'datasetVersion': 2})
        self.assertEqual('warning', result['status'])
        self.assertIn('Refresh', result['message'])

    def test_family_failure_does_not_undo_supabase_save(self):
        events = []
        class Group:
            def __init__(self, *args): pass
            def Start(self): events.append('start')
            def GetStatus(self): return 'started'
            def RollBack(self): events.append('rollback-family')
        class Status:
            Started = 'started'
        self.api.update({
            'get_storage_mode': lambda doc: 'annotation', 'TransactionGroup': Group, 'TransactionStatus': Status,
            'build_annotation_keynote_payload': lambda *args, **kwargs: {'status': 'ready', 'libraryKey': 'annotation:x',
                'datasetVersion': 2, 'entries': [row('A', 'Saved in cloud')]},
            'sync_generic_annotation_types': lambda *args, **kwargs: {'failedCount': 1, 'failures': ['Type owned by colleague']},
            'keynote_supabase_rpc': lambda *args: self.fail('Family update attempted a database write')
        })
        result = self.api['sync_annotation_family_payload'](object(), {
            'libraryKey': 'annotation:x', 'datasetVersion': 2, 'baselineEntries': [row('A', 'Old')]})
        self.assertEqual('warning', result['status'])
        self.assertIn('remains saved', result['message'])
        self.assertEqual(['start', 'rollback-family'], events)

    def test_cloud_unavailable_never_falls_back_to_file_or_rvt(self):
        self.api['build_base_payload'] = lambda *args: {'preferences': {}, 'entries': []}
        self.api['get_document_title'] = lambda doc: 'Project'
        self.api['annotation_library_key'] = lambda doc: 'annotation:project'
        def offline(*args): raise Exception('offline')
        self.api['keynote_supabase_rpc'] = offline
        result = self.api['build_annotation_keynote_payload'](object())
        self.assertEqual('error', result['status'])
        self.assertEqual([], result['entries'])
        self.assertFalse(result['writeAvailable'])


class ProjectKeynoteSetupTests(unittest.TestCase):
    def setUp(self):
        self.api = load_functions()
        self.reference_adapter = self.api['get_keynote_reference']
        self.identity = {'documentKey': 'central.rvt', 'documentKeySource': 'centralPath'}
        self.settings = {}
        self.reference_path = ''
        self.calls = []
        self.saved_modes = []
        self.snapshot = {'status': 'error', 'message': 'Keynote library was not found.', 'entries': []}
        self.api.update({
            'get_document_analytics_identity': lambda doc: self.identity.copy(),
            'read_user_settings': lambda: self.settings,
            'get_document_title': lambda doc: 'Project',
            'get_keynote_reference': lambda *args, **kwargs: ('table', None, self.reference_path),
            'build_base_payload': self.base_payload,
            'keynote_supabase_rpc': self.rpc,
            'save_storage_mode': lambda doc, mode: self.saved_modes.append(mode),
            'normalize_path': lambda path: path.lower(),
            'os': types.SimpleNamespace(path=types.SimpleNamespace(exists=lambda path: False, isabs=ntpath.isabs)),
        })

    def base_payload(self, doc, status, message):
        return dict(self.identity, status=status, message=message, storageMode='file', entries=[],
                    preferences={}, libraryKey='', keynotePath='', writeAvailable=False, issues=[],
                    projectSetup={'state': 'complete', 'canContinue': False, 'message': ''})

    def rpc(self, name, arguments, allow_missing=False):
        self.calls.append((name, arguments))
        if isinstance(self.snapshot, Exception):
            raise self.snapshot
        return self.api['check_keynote_rpc_result'](copy.deepcopy(self.snapshot), allow_missing)

    def load(self):
        return self.api['build_project_keynote_payload'](object(), include_model_health=False)

    def existing_library(self, source='annotation'):
        self.snapshot = {'status': 'ready', 'sourceType': source, 'libraryId': 'existing-library',
                         'datasetVersion': 9, 'entries': [row('A', 'Existing cloud edit')]}

    def test_new_saved_project_shows_setup_after_confirmed_absence(self):
        result = self.load()
        self.assertEqual('setupRequired', result['status'])
        self.assertEqual('required', result['projectSetup']['state'])
        self.assertTrue(result['projectSetup']['canContinue'])
        self.assertEqual('annotation:central.rvt', self.calls[0][1]['p_library_key'])
        self.assertEqual([], self.saved_modes)

    def test_unsaved_project_shows_instruction_without_cloud_lookup(self):
        self.identity = {'documentKey': 'title:Project', 'documentKeySource': 'title'}
        result = self.load()
        self.assertEqual('required', result['projectSetup']['state'])
        self.assertFalse(result['projectSetup']['canContinue'])
        self.assertIn('Save the Revit project', result['message'])
        self.assertEqual([], self.calls)

    def test_missing_or_unsupported_assigned_file_keeps_existing_error_view(self):
        for path, status in (('missing.txt', 'missingFile'), ('https://example/keynotes', 'unsupported')):
            self.reference_path = path
            result = self.load()
            self.assertEqual(status, result['status'])
            self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)

    def test_assigned_unreadable_or_invalid_file_is_never_setup(self):
        self.reference_path = 'assigned.txt'
        self.api['os'].path.exists = lambda path: True
        self.api['read_binary_file'] = lambda path: b'invalid'
        self.api['decode_keynote_bytes'] = lambda raw: ('A\tAlpha\nA\tDuplicate\n', 'utf-8')
        self.api['get_file_state'] = lambda path: {}
        self.api['check_file_write_available'] = lambda path: (True, '')
        self.assertEqual('invalidFormat', self.load()['status'])
        def unreadable(path): raise Exception('Access denied')
        self.api['read_binary_file'] = unreadable
        result = self.load()
        self.assertEqual('error', result['status'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)

    def test_existing_cloud_library_is_recovered_without_reseeding(self):
        self.existing_library()
        result = self.load()
        self.assertEqual('ready', result['status'])
        self.assertEqual('annotation', result['storageMode'])
        self.assertEqual('genericAnnotation', result['preferences']['placementMode'])
        self.assertEqual('Existing cloud edit', result['entries'][0]['text'])
        self.assertEqual('existing-library', result['libraryId'])
        self.assertEqual(['annotation'], self.saved_modes)
        self.assertEqual(['get_annotation_keynote_snapshot'], [call[0] for call in self.calls])

    def test_explicit_file_preference_is_preserved(self):
        self.settings = {'storageModes': {'central.rvt': 'file'}}
        self.existing_library()
        result = self.load()
        self.assertEqual('file', result['storageMode'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertIn('existing Supabase', result['message'])
        self.assertEqual([], self.saved_modes)

    def test_explicit_annotation_preference_overrides_an_assigned_text_file(self):
        self.settings = {'storageModes': {'central.rvt': 'annotation'}}
        self.reference_path = 'missing.txt'
        self.existing_library()
        result = self.load()
        self.assertEqual('ready', result['status'])
        self.assertEqual('annotation', result['storageMode'])
        self.assertEqual([], self.saved_modes)

    def test_configured_annotations_do_not_depend_on_native_reference_inspection(self):
        self.settings = {'storageModes': {'central.rvt': 'annotation'}}
        self.existing_library()
        def inaccessible(*args, **kwargs): raise Exception('Cannot inspect native reference')
        self.api['get_keynote_reference'] = inaccessible
        result = self.load()
        self.assertEqual('ready', result['status'])
        self.assertEqual('annotation', result['storageMode'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.saved_modes)

    def test_converted_library_uses_existing_recovery_view(self):
        self.existing_library('file')
        result = self.load()
        self.assertEqual('error', result['status'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertIn('converted to Text File', result['message'])
        self.assertEqual([], self.saved_modes)

    def test_failed_or_malformed_cloud_check_blocks_creation(self):
        for snapshot in (Exception('Offline'), {'status': 'error', 'message': 'Migration missing'},
                         {}, [], {'notFound': True}):
            self.snapshot = snapshot
            result = self.load()
            self.assertEqual('error', result['status'])
            self.assertEqual('blocked', result['projectSetup']['state'])
            self.assertFalse(result['projectSetup']['canContinue'])
        self.assertEqual([], self.saved_modes)

    def test_only_explicit_not_found_response_can_mean_absent(self):
        with self.assertRaisesRegex(Exception, 'not found'):
            self.api['check_keynote_rpc_result'](self.snapshot)
        self.assertTrue(self.api['check_keynote_rpc_result'](self.snapshot, allow_missing=True)['notFound'])

    def test_reference_read_failures_and_pathless_references_are_not_new_projects(self):
        # Exercise the real Revit reference adapter, with only the API boundary mocked.
        self.api['get_keynote_reference'] = self.reference_adapter
        self.api['get_reference_path'] = lambda ref: ''
        for references in ({'keynote': object()}, object(), None):
            table = types.SimpleNamespace(GetExternalResourceReferences=lambda: references)
            self.api['KeynoteTable'] = types.SimpleNamespace(GetKeynoteTable=lambda doc: table)
            result = self.load()
            self.assertEqual('error', result['status'])
            self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)
        table = types.SimpleNamespace(GetExternalResourceReferences=lambda: {})
        self.api['KeynoteTable'] = types.SimpleNamespace(GetKeynoteTable=lambda doc: table)
        self.assertEqual('required', self.load()['projectSetup']['state'])

    def reference_table(self, resource, stored_path=None, absolute_path=None):
        file_ref = types.SimpleNamespace(GetPath=lambda: stored_path, GetAbsolutePath=lambda: absolute_path)
        table = types.SimpleNamespace(
            GetExternalResourceReferences=lambda: {'keynote': resource},
            IsExternalFileReference=lambda: stored_path is not None,
            GetExternalFileReference=lambda: file_ref)
        self.api['get_keynote_reference'] = self.reference_adapter
        self.api['KeynoteTable'] = types.SimpleNamespace(GetKeynoteTable=lambda doc: table)
        self.api['ModelPathUtils'] = types.SimpleNamespace(ConvertModelPathToUserVisiblePath=lambda path: path)
        return table

    def blank_resource(self, info=None, server='bd4f0f53-394a-4468-b37e-1e7949013382'):
        return types.SimpleNamespace(InSessionPath='', ServerId=server,
            HasValidDisplayPath=lambda: False,
            GetReferenceInformation=lambda: info if info is not None else {},
            IsValidReference=lambda resource_type: False)

    def test_blank_builtin_reference_opens_automatic_setup(self):
        for info in ({}, {'Path': ''}, {'Path': '', 'PathType': 'Absolute'},
                     {'Path': '  ', 'PathType': 'Relative to Library Locations'}):
            self.reference_table(self.blank_resource(info), stored_path='')
            result = self.load()
            self.assertEqual('required', result['projectSetup']['state'])
            self.assertTrue(result['projectSetup']['canContinue'])
        self.assertEqual([], self.saved_modes)

    def test_null_or_unassigned_server_reference_opens_automatic_setup(self):
        for resource in (None, self.blank_resource(server='00000000-0000-0000-0000-000000000000')):
            self.reference_table(resource)
            self.assertEqual('required', self.load()['projectSetup']['state'])

    def test_blank_builtin_reference_in_unsaved_project_requests_save_and_refresh(self):
        self.identity = {'documentKey': 'title:Project', 'documentKeySource': 'title'}
        self.reference_table(self.blank_resource({'Path': '', 'PathType': 'Absolute'}), stored_path='')
        result = self.load()
        self.assertEqual('required', result['projectSetup']['state'])
        self.assertFalse(result['projectSetup']['canContinue'])
        self.assertIn('Save the Revit project', result['message'])
        self.assertEqual([], self.calls)

    def test_blank_builtin_reference_recovers_existing_cloud_library(self):
        self.reference_table(self.blank_resource({'Path': '', 'PathType': 'Absolute'}), stored_path='')
        self.existing_library()
        result = self.load()
        self.assertEqual('ready', result['status'])
        self.assertEqual('existing-library', result['libraryId'])
        self.assertEqual('Existing cloud edit', result['entries'][0]['text'])

    def test_pathless_configured_resource_is_not_mistaken_for_a_blank_reference(self):
        resources = [self.blank_resource({'Path': 'missing.txt', 'PathType': 'Absolute'}),
                     self.blank_resource({'ResourceId': 'cloud-notes'}, server='custom-server'),
                     self.blank_resource(server='custom-server')]
        valid_display = self.blank_resource()
        valid_display.HasValidDisplayPath = lambda: True
        resources.append(valid_display)
        for resource in resources:
            self.reference_table(resource)
            result = self.load()
            self.assertEqual('error', result['status'])
            self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)

    def test_stored_file_reference_recovers_path_without_a_session_path(self):
        self.reference_table(self.blank_resource({'Path': 'notes.txt', 'PathType': 'Relative to Central Model'}),
                             stored_path='notes.txt', absolute_path='C:/Project/notes.txt')
        result = self.load()
        self.assertEqual('missingFile', result['status'])
        self.assertEqual('C:/Project/notes.txt', result['keynotePath'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)

    def test_unresolved_stored_file_path_and_native_inspection_failure_never_open_setup(self):
        table = self.reference_table(self.blank_resource(), stored_path='notes.txt')
        unresolved = self.load()
        self.assertEqual('error', unresolved['status'])
        self.assertEqual('complete', unresolved['projectSetup']['state'])
        self.assertIn('Could not resolve', unresolved['message'])
        def failed(): raise Exception('Reference access failed')
        table.GetExternalFileReference = failed
        result = self.load()
        self.assertEqual('error', result['status'])
        self.assertEqual('complete', result['projectSetup']['state'])
        self.assertEqual([], self.calls)

    def test_stored_absolute_path_is_preserved_when_absolute_resolution_fails(self):
        self.reference_table(self.blank_resource({'Path': 'C:/Project/notes.txt'}),
                             stored_path='C:/Project/notes.txt')
        result = self.load()
        self.assertEqual('missingFile', result['status'])
        self.assertEqual('C:/Project/notes.txt', result['keynotePath'])
        self.assertEqual('complete', result['projectSetup']['state'])

    def test_resource_short_display_name_is_not_used_as_a_file_path(self):
        resource = self.blank_resource({'Path': 'notes.txt'})
        resource.HasValidDisplayPath = lambda: True
        resource.GetResourceShortDisplayName = lambda: 'notes.txt'
        self.assertEqual('', self.api['get_reference_path'](resource))

    def test_setup_preflight_blocks_side_effects_if_project_was_configured_elsewhere(self):
        self.existing_library()
        self.api['template_entries'] = lambda *args: self.fail('Configured project requested template')
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file', 'createFile': True, 'projectSetup': True})
        self.assertEqual('error', result['status'])
        self.assertEqual(['get_annotation_keynote_snapshot'], [call[0] for call in self.calls])

    def test_setup_preflight_blocks_creation_after_cloud_check_failure(self):
        self.snapshot = Exception('Offline')
        self.api['template_entries'] = lambda *args: self.fail('Offline setup requested template')
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'file', 'createFile': True, 'projectSetup': True})
        self.assertEqual('error', result['status'])
        self.assertEqual([], self.saved_modes)

    def test_revit_setup_requests_template_then_assigns_and_loads_before_persisting_mode(self):
        request = {'storageMode': 'file', 'createFile': True, 'projectSetup': True}
        first = self.api['setup_keynote_storage'](object(), request)
        self.assertEqual('needsTemplate', first['status'])
        self.assertTrue(first['request']['projectSetup'])
        events = []
        def dialog(**kwargs):
            self.assertEqual('RevitKeynotes.txt', kwargs['default_name'])
            events.append('dialog')
            return 'project/RevitKeynotes.txt'
        def assign(doc, path, entries):
            events.append('assign')
            self.assertEqual('project/RevitKeynotes.txt', path)
            content = self.api['canonicalize_entries'](entries, '\r\n', path, 'utf-16')
            data = self.api['encode_keynote_text'](content, 'utf-16')
            self.assertTrue(data.startswith(codecs.BOM_UTF16_LE))
            self.assertEqual(21, len(entries))
        def load(*args, **kwargs):
            events.append('load')
            result = self.base_payload(object(), 'ready', 'Loaded file')
            result['keynotePath'] = 'project/RevitKeynotes.txt'
            return result
        self.api['forms'] = types.SimpleNamespace(save_file=dialog)
        self.api['write_and_assign_keynote_file'] = assign
        self.api['build_keynote_payload'] = load
        self.api['save_storage_mode'] = lambda doc, mode: events.append('mode:' + mode)
        template = (ROOT / 'supabase/keynote_annotation_library.sql').read_text(encoding='utf-8').split('$template$')[1]
        request['template'] = {'content': template}
        result = self.api['setup_keynote_storage'](object(), request)
        self.assertEqual('ready', result['status'])
        self.assertEqual(['dialog', 'assign', 'load', 'mode:file'], events)

    def test_generic_setup_initializes_cloud_only_and_loads_before_persisting_mode(self):
        def rpc(name, arguments, allow_missing=False):
            if name == 'ensure_annotation_keynote_library':
                self.assertIsNone(arguments['p_seed_entries'])
                self.existing_library()
            return self.rpc(name, arguments, allow_missing)
        self.api['keynote_supabase_rpc'] = rpc
        self.api['build_model_health'] = lambda *args: {'placedKeyMap': {}}
        self.api['write_and_assign_keynote_file'] = lambda *args: self.fail('Annotation setup created a file')
        self.api['sync_generic_annotation_types'] = lambda *args: self.fail('Setup prepared a family')
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'annotation', 'projectSetup': True})
        self.assertEqual('ready', result['status'])
        self.assertEqual('genericAnnotation', result['payload']['preferences']['placementMode'])
        self.assertEqual(['annotation'], self.saved_modes)
        self.assertEqual(['get_annotation_keynote_snapshot', 'ensure_annotation_keynote_library',
                          'get_annotation_keynote_snapshot'], [call[0] for call in self.calls])

    def test_new_file_setup_cancel_or_assignment_failure_keeps_mode_for_retry(self):
        request = {'storageMode': 'file', 'createFile': True, 'projectSetup': True,
                   'template': {'content': 'A\tAlpha\n'}}
        assigned = []
        def failed_assignment(*args):
            assigned.append(True)
            raise Exception('File created at project/notes.txt, but assignment failed')
        self.api['write_and_assign_keynote_file'] = failed_assignment
        self.api['forms'] = types.SimpleNamespace(save_file=lambda **kwargs: None)
        canceled = self.api['setup_keynote_storage'](object(), request)
        self.assertEqual('canceled', canceled['status'])
        self.assertEqual([], assigned)
        self.api['forms'].save_file = lambda **kwargs: 'project/notes.txt'
        failed = self.api['setup_keynote_storage'](object(), request)
        self.assertEqual('error', failed['status'])
        self.assertIn('project/notes.txt', failed['message'])
        self.assertEqual([True], assigned)
        self.assertEqual([], self.saved_modes)

    def test_new_file_setup_rejects_invalid_template_before_opening_save_as(self):
        self.api['forms'] = types.SimpleNamespace(save_file=lambda **kwargs: self.fail('Invalid template opened Save As'))
        result = self.api['setup_keynote_storage'](object(), {
            'storageMode': 'file', 'createFile': True, 'projectSetup': True,
            'template': {'content': 'A\tAlpha\nA\tDuplicate\n'}})
        self.assertEqual('error', result['status'])
        self.assertEqual([], self.saved_modes)

    def test_manual_recovery_can_request_template_when_reference_detection_fails(self):
        self.api['build_project_keynote_payload'] = lambda *args, **kwargs: self.fail('Manual recovery used automatic detection')
        self.api['get_storage_mode'] = lambda doc: 'annotation'
        result = self.api['setup_keynote_storage'](object(), {
            'storageMode': 'file', 'createFile': True, 'recoverySetup': True})
        self.assertEqual('needsTemplate', result['status'])
        self.assertTrue(result['request']['recoverySetup'])
        self.assertEqual([], self.saved_modes)

    def test_manual_recovery_cloud_check_failure_does_not_create_or_replace_a_library(self):
        self.snapshot = Exception('Offline')
        self.api['forms'] = types.SimpleNamespace(save_file=lambda **kwargs: self.fail('Failed lookup opened Save As'))
        result = self.api['setup_keynote_storage'](object(), {
            'storageMode': 'file', 'createFile': True, 'recoverySetup': True})
        self.assertEqual('error', result['status'])
        self.assertEqual([], self.saved_modes)

    def test_manual_annotation_recovery_preserves_existing_cloud_notes_despite_broken_reference(self):
        self.existing_library()
        def broken(*args, **kwargs): raise Exception('The assigned keynote reference does not expose a readable file path.')
        self.api['get_keynote_reference'] = broken
        self.api['build_project_keynote_payload'] = lambda *args, **kwargs: self.fail('Manual recovery used automatic detection')
        def rpc(name, arguments, allow_missing=False):
            if name == 'ensure_annotation_keynote_library':
                self.assertIsNone(arguments['p_seed_entries'])
            return self.rpc(name, arguments, allow_missing)
        self.api['keynote_supabase_rpc'] = rpc
        self.api['build_model_health'] = lambda *args: {'placedKeyMap': {}}
        result = self.api['setup_keynote_storage'](object(), {'storageMode': 'annotation', 'recoverySetup': True})
        self.assertEqual('ready', result['status'])
        self.assertEqual('existing-library', result['payload']['libraryId'])
        self.assertEqual('Existing cloud edit', result['payload']['entries'][0]['text'])
        self.assertEqual(['annotation'], self.saved_modes)

    def test_explicit_file_source_can_load_without_reference_discovery(self):
        self.api['get_keynote_reference'] = lambda *args, **kwargs: self.fail('Selected path still used reference discovery')
        self.api['os'].path.exists = lambda path: True
        self.api['read_binary_file'] = lambda path: 'A\tAlpha\n'.encode('utf-16')
        self.api['get_file_state'] = lambda path: {}
        self.api['check_file_write_available'] = lambda path: (True, '')
        result = self.api['build_keynote_payload'](object(), include_model_health=False,
                                                  storage_mode='file', source_path='selected.txt')
        self.assertEqual('ready', result['status'])
        self.assertEqual('selected.txt', result['keynotePath'])
        self.assertEqual('Alpha', result['entries'][0]['text'])


class KeynoteFileReconnectTests(unittest.TestCase):
    def setUp(self):
        self.api = load_functions()
        self.events = []
        self.content = 'A\tAlpha\nA.1\tCaf\u00e9\tA\n'.encode('utf-16')
        self.selected = 'project/notes.txt'
        self.payload = {'status': 'ready', 'storageMode': 'file', 'entries': [row('A', 'Alpha')],
                        'projectSetup': {'state': 'complete'}, 'keynotePath': self.selected}
        self.assignment_failure = False
        self.commit_failure = False
        events = self.events
        class Group:
            def __init__(self, *args): pass
            def Start(self): events.append('start')
            def Assimilate(group):
                events.append('commit')
                return 'failed' if self.commit_failure else 'committed'
            def GetStatus(self): return 'started'
            def RollBack(self): events.append('rollback')
        def pick(**kwargs):
            self.assertEqual('txt', kwargs['file_ext'])
            events.append('picker')
            return self.selected
        def read(path):
            self.assertEqual(self.selected, path)
            events.append('read')
            if isinstance(self.content, Exception): raise self.content
            return self.content
        def assign(doc, path):
            self.assertEqual(self.selected, path)
            events.append('assign')
            if self.assignment_failure: raise Exception('Revit assignment failed')
        def load(doc, **kwargs):
            self.assertEqual('file', kwargs['storage_mode'])
            self.assertEqual(self.selected, kwargs['source_path'])
            events.append('load')
            return copy.deepcopy(self.payload)
        self.api.update({
            'get_document_analytics_identity': lambda doc: {'documentKeySource': 'path'},
            'forms': types.SimpleNamespace(pick_file=pick),
            'read_binary_file': read, 'assign_keynote_file': assign,
            'build_keynote_payload': load, 'TransactionGroup': Group,
            'TransactionStatus': types.SimpleNamespace(Started='started', Committed='committed'),
            'save_storage_mode': lambda doc, mode: events.append('mode:' + mode),
            'write_and_assign_keynote_file': lambda *args: self.fail('Reconnect rewrote existing file'),
            'keynote_supabase_rpc': lambda *args, **kwargs: self.fail('Reconnect required Supabase'),
        })

    def reconnect(self):
        return self.api['reconnect_keynote_file'](object())

    def test_reconnect_validates_assigns_and_loads_before_committing_mode_without_file_writes(self):
        original = self.content
        result = self.reconnect()
        self.assertEqual('ready', result['status'])
        self.assertEqual(self.selected, result['payload']['keynotePath'])
        self.assertEqual(original, self.content)
        self.assertEqual(['picker', 'read', 'start', 'assign', 'load', 'commit', 'mode:file'], self.events)

    def test_cancel_preserves_original_assignment_and_mode(self):
        self.selected = None
        self.assertEqual('canceled', self.reconnect()['status'])
        self.assertEqual(['picker'], self.events)

    def test_unsaved_project_cannot_open_picker(self):
        self.api['get_document_analytics_identity'] = lambda doc: {'documentKeySource': 'title'}
        result = self.reconnect()
        self.assertEqual('error', result['status'])
        self.assertIn('Save the Revit project', result['message'])
        self.assertEqual([], self.events)

    def test_invalid_or_unreadable_file_is_rejected_before_assignment(self):
        for content in ('A\tAlpha\nA\tDuplicate\n'.encode('utf-16'), Exception('File not found')):
            self.content = content
            self.events.clear()
            self.assertEqual('error', self.reconnect()['status'])
            self.assertEqual(['picker', 'read'], self.events)

    def test_assignment_failure_rolls_back_and_preserves_mode(self):
        self.assignment_failure = True
        self.assertEqual('error', self.reconnect()['status'])
        self.assertEqual(['picker', 'read', 'start', 'assign', 'rollback'], self.events)

    def test_failed_load_rolls_back_assignment_and_preserves_mode(self):
        self.payload = {'status': 'error', 'message': 'File changed during load'}
        self.assertEqual('error', self.reconnect()['status'])
        self.assertEqual(['picker', 'read', 'start', 'assign', 'load', 'rollback'], self.events)

    def test_failed_commit_never_persists_mode(self):
        self.commit_failure = True
        self.assertEqual('error', self.reconnect()['status'])
        self.assertNotIn('mode:file', self.events)

    def test_assignment_adapter_checks_revit_load_results_and_rolls_back_on_failure(self):
        api = load_functions()
        events = []
        class Transaction:
            def __init__(self, *args): pass
            def Start(self): events.append('start')
            def Commit(self): events.append('commit'); return 'committed'
            def GetStatus(self): return 'started'
            def RollBack(self): events.append('rollback')
        api.update({'Transaction': Transaction,
                    'TransactionStatus': types.SimpleNamespace(Started='started', Committed='committed'),
                    'ExternalResourceTypes': types.SimpleNamespace(BuiltInExternalResourceTypes=types.SimpleNamespace(KeynoteTable='keynote')),
                    'ExternalResourceReference': types.SimpleNamespace(CreateLocalResource=lambda *args: 'reference'),
                    'ModelPathUtils': types.SimpleNamespace(ConvertUserVisiblePathToModelPath=lambda path: path),
                    'PathType': types.SimpleNamespace(Absolute='absolute')})
        for result, failures in [('Success', []), ('Failed', []), ('Success', ['Bad rows'])]:
            events.clear()
            api['KeynoteTable'] = types.SimpleNamespace(GetKeynoteTable=lambda doc: types.SimpleNamespace(
                LoadFrom=lambda *args: result))
            api['KeyBasedTreeEntriesLoadResults'] = lambda: types.SimpleNamespace(GetFailureMessages=lambda: failures)
            if result == 'Success' and not failures:
                api['assign_keynote_file'](object(), self.selected)
                self.assertEqual(['start', 'commit'], events)
            else:
                with self.assertRaisesRegex(Exception, 'could not load'):
                    api['assign_keynote_file'](object(), self.selected)
                self.assertEqual(['start', 'rollback'], events)


if __name__ == '__main__':
    unittest.main()
