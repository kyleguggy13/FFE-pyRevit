"""Behavior tests for keynote storage without loading Revit or starting the UI.

The actual pure functions and storage workflow are loaded from script.py via AST.
Revit transactions are replaced only at the boundary to verify atomic rollback.
Run: python -m unittest discover -s tests -p 'test_keynote_annotation_library.py'
"""
import ast
import codecs
import copy
import pathlib
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


if __name__ == '__main__':
    unittest.main()
