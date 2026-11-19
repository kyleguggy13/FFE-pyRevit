# -*- coding: utf-8 -*-
"""Purpose: verify Space snapshots, bridge identity, navigation, and modeless safety.
Behaviors: exercise real helpers against fake API boundaries without opening Revit.
Revit API notes: fake IDs deliberately exceed JavaScript's exact integer range.
Design: test observable data/navigation behavior rather than implementation shape.
"""
import ast
import collections
import importlib.util
import json
import pathlib
import types
import unittest
from unittest import mock

BUTTON = pathlib.Path(__file__).resolve().parents[1]
ROOT = BUTTON.parents[3]
LIBRARY = BUTTON / 'lib'
COMMAND = BUTTON / 'script.py'
# The compatibility scenarios are shared with Keynotes, whose fixture remains
# in the repository tests directory. Load it without modifying sys.path.
fixtureSpec = importlib.util.spec_from_file_location(
    'revit_data_table_loader_fixture', ROOT / 'tests/test_keynote_webview2.py')
loader_fixture = importlib.util.module_from_spec(fixtureSpec)
fixtureSpec.loader.exec_module(loader_fixture)
spec = importlib.util.spec_from_file_location('revit_data_table', LIBRARY / 'revit_data_table.py')
core = importlib.util.module_from_spec(spec)
spec.loader.exec_module(core)
NS = types.SimpleNamespace


class Id:
    """Expose only the supported 64-bit API."""
    def __init__(self, value):
        self.Value = value


class Space:
    def __init__(self, value, phase=10, level=20, area=100, height=9, placed=True):
        self.Id = Id(value)
        self.UniqueId = 'uid-' + str(value)
        self.LevelId = Id(level)
        self.SpacePhaseId = Id(phase)
        self.CreatedPhaseId = Id(-1)
        self.Area = area
        self.UnboundedHeight = height
        self.Number = '010'
        self.Name = 'Combined name 010'
        self.Location = object() if placed else None
        self.Category = NS(Id=Id(-3600))

    def get_Parameter(self, parameter):
        if parameter in ('spacePhaseId', 'spacePhase'):
            return NS(AsElementId=lambda: self.SpacePhaseId)
        return NS(AsString=lambda: 'Office' if parameter == 'name' else self.Number)

    def get_BoundingBox(self, view):
        return object() if self.Location is not None else None


class Collector:
    def __init__(self, document):
        self.document = document

    def OfCategory(self, category):
        assert category == -3600
        return self

    def WhereElementIsNotElementType(self):
        return self.document.spaces


class SpaceSnapshotTests(unittest.TestCase):
    def setUp(self):
        self.space = Space(9007199254740993)
        self.elements = {10: NS(Id=Id(10), Name='Existing'),
                         11: NS(Id=Id(11), Name='New Construction'),
                         20: NS(Id=Id(20), Name='Level 1', Elevation=0)}
        self.document = NS(
            Title='HVAC Model', spaces=[self.space],
            Phases=[self.elements[10], self.elements[11]],
            ActiveView=NS(get_Parameter=lambda key: NS(AsElementId=lambda: Id(10))),
            GetUnits=lambda: 'project-units',
            GetElement=lambda key: self.elements.get(key.Value) if isinstance(key, Id) else
            next((space for space in self.document.spaces if space.UniqueId == key), None))
        self.formatCalls = []

        def format_value(units, spec, value, editing):
            self.formatCalls.append((units, spec, value, editing))
            return '{} {}'.format(value, 'm²' if spec == 'area' else 'm')

        self.db = NS(FilteredElementCollector=Collector,
                     BuiltInCategory=NS(OST_MEPSpaces=-3600),
                     BuiltInParameter=NS(VIEW_PHASE='viewPhase', ROOM_NAME='name', ROOM_NUMBER='number',
                                         ROOM_PHASE_ID='spacePhaseId', ROOM_PHASE='spacePhase'),
                     SpecTypeId=NS(Area='area', Length='length'),
                     UnitFormatUtils=NS(Format=format_value))

    def snapshot(self):
        return core.BuildSnapshot(self.document, 'document-key', self.db)

    def test_snapshot_keeps_large_ids_strings_and_raw_measurements(self):
        snapshot = self.snapshot()
        row = snapshot['rows'][0]
        self.assertEqual('9007199254740993', row['id'])
        self.assertEqual('Office', row['name'])
        self.assertEqual('010', row['number'])
        self.assertEqual((100, 9), (row['areaValue'], row['heightValue']))
        self.assertEqual(('100.0 m²', '9.0 m'), (row['areaText'], row['heightText']))
        self.assertEqual(('10', '20', 0), (row['phaseId'], row['levelId'], row['levelElevation']))
        self.assertEqual('10', snapshot['defaultPhaseId'])
        self.assertTrue(all(call[0] == 'project-units' for call in self.formatCalls))
        json.dumps(snapshot, allow_nan=False)

    def test_no_view_phase_uses_last_project_phase(self):
        self.document.ActiveView = NS(get_Parameter=lambda key: None)
        self.assertEqual('11', self.snapshot()['defaultPhaseId'])

    def test_space_phase_parameter_takes_precedence_over_element_creation_phase(self):
        self.space.CreatedPhaseId = Id(11)
        row = self.snapshot()['rows'][0]
        self.assertEqual(('10', 'Existing'), (row['phaseId'], row['phase']))
        self.assertFalse(any('phase is unavailable' in warning for warning in row['warnings']))

    def test_space_phase_falls_back_to_room_phase_when_phase_id_cannot_be_read(self):
        readParameter = self.space.get_Parameter
        for failure in ('missing', 'invalid', 'throws', 'unknown'):
            with self.subTest(failure=failure):
                def parameter(parameterId):
                    if parameterId != 'spacePhaseId':
                        return readParameter(parameterId)
                    if failure == 'missing':
                        return None
                    if failure == 'invalid':
                        return NS(AsElementId=lambda: Id(-1))
                    if failure == 'unknown':
                        return NS(AsElementId=lambda: Id(999))
                    raise ValueError('parameter unavailable')
                self.space.get_Parameter = parameter
                snapshot = self.snapshot()
                self.assertEqual('10', snapshot['rows'][0]['phaseId'])
                self.assertEqual(0, snapshot['skippedCount'])

    def test_unresolved_space_phase_keeps_row_without_assigning_view_or_creation_phase(self):
        readParameter = self.space.get_Parameter
        self.space.CreatedPhaseId = Id(11)
        self.space.get_Parameter = lambda parameterId: (
            None if parameterId in ('spacePhaseId', 'spacePhase') else readParameter(parameterId))
        snapshot = self.snapshot()
        row = snapshot['rows'][0]
        self.assertIsNone(row['phaseId'])
        self.assertEqual('—', row['phase'])
        self.assertTrue(any('phase is unavailable' in warning for warning in row['warnings']))
        self.assertEqual(0, snapshot['skippedCount'])

    def test_unplaced_and_zero_area_spaces_are_retained(self):
        self.document.spaces = [Space(1, area=0, placed=False), Space(2, area=0)]
        rows = self.snapshot()['rows']
        self.assertEqual(2, len(rows))
        self.assertEqual(0, rows[0]['areaValue'])
        self.assertIn('Unplaced', rows[0]['warnings'][0])
        self.assertIn('Zero area', rows[1]['warnings'][0])

    def test_missing_level_and_measurement_stay_distinct_from_zero(self):
        self.space.LevelId = Id(-1)
        self.space.UnboundedHeight = None
        row = self.snapshot()['rows'][0]
        self.assertEqual('—', row['level'])
        self.assertIsNone(row['heightValue'])
        self.assertEqual('—', row['heightText'])
        self.assertTrue(any('Height unavailable' in warning for warning in row['warnings']))

    def test_nonfinite_value_is_not_serialized_as_nan(self):
        self.space.Area = float('nan')
        snapshot = self.snapshot()
        self.assertIsNone(snapshot['rows'][0]['areaValue'])
        json.dumps(snapshot, allow_nan=False)

    def test_one_unreadable_space_does_not_hide_other_rows(self):
        broken = Space(2)
        broken.get_Parameter = mock.Mock(side_effect=ValueError('bad parameter'))
        self.document.spaces.append(broken)
        snapshot = self.snapshot()
        self.assertEqual(1, len(snapshot['rows']))
        self.assertEqual(1, snapshot['skippedCount'])
        self.assertIn('Space 2: bad parameter', snapshot['issues'][0])

    def test_empty_model_is_a_valid_snapshot(self):
        self.document.spaces = []
        self.assertEqual([], self.snapshot()['rows'])

    def test_navigation_skips_deleted_rows_and_selects_remaining_spaces(self):
        other = Space(2)
        self.document.spaces.append(other)
        snapshot = self.snapshot()
        self.document.spaces.remove(other)
        select = mock.Mock()
        uiDocument = NS(Document=self.document, Selection=NS(SetElementIds=select))
        result = core.NavigateSpaces(uiDocument, snapshot, {
            'action': 'select', 'elementIds': [str(self.space.Id.Value), '2']}, self.db, list)
        self.assertEqual([self.space.Id], select.call_args.args[0])
        self.assertEqual('warning', result['kind'])
        self.assertIn('Skipped 1', result['message'])

    def test_show_skips_unplaced_spaces(self):
        self.space.Location = None
        uiDocument = NS(Document=self.document, ShowElements=mock.Mock())
        result = core.NavigateSpaces(uiDocument, self.snapshot(), {
            'action': 'show', 'elementIds': [str(self.space.Id.Value)]}, self.db, list)
        uiDocument.ShowElements.assert_not_called()
        self.assertIn('no viewable geometry', result['message'])

    def test_replaced_element_identity_cannot_select_another_element(self):
        snapshot = self.snapshot()
        replacement = Space(22)
        self.document.GetElement = lambda key: replacement
        select = mock.Mock()
        result = core.NavigateSpaces(NS(Document=self.document, Selection=NS(SetElementIds=select)),
                                     snapshot, {'action': 'select', 'elementIds': [snapshot['rows'][0]['id']]},
                                     self.db, list)
        select.assert_not_called()
        self.assertIn('deleted or replaced', result['message'])

    def test_show_failure_reports_how_to_retry(self):
        uiDocument = NS(Document=self.document, ShowElements=mock.Mock(side_effect=ValueError('no view')))
        result = core.NavigateSpaces(uiDocument, self.snapshot(), {
            'action': 'show', 'elementIds': [str(self.space.Id.Value)]}, self.db, list)
        self.assertEqual('warning', result['kind'])
        self.assertIn('suitable plan or 3D view', result['message'])


class RequestTests(unittest.TestCase):
    def test_rejects_malformed_foreign_and_unknown_element_requests(self):
        invalid = [None, [], {'action': 'save'}, {'action': 'refresh', 'documentKey': 'other'},
                   {'action': 'select', 'documentKey': 'doc', 'elementIds': [1]},
                   {'action': 'show', 'documentKey': 'doc', 'elementIds': ['unknown']},
                   {'action': 'select', 'documentKey': 'doc', 'elementIds': []},
                   {'action': 'select', 'documentKey': 'doc', 'elementIds': '1'}]
        for message in invalid:
            with self.subTest(message=message), self.assertRaises(ValueError):
                core.ValidateRequest(message, 'doc', {'rows': [{'id': '1'}]})

    def test_deduplicates_known_string_ids(self):
        request = core.ValidateRequest({'action': 'select', 'documentKey': 'doc',
                                        'elementIds': ['1', '1']}, 'doc', {'rows': [{'id': '1'}]})
        self.assertEqual(['1'], request['elementIds'])


def host_classes():
    """Load real host classes with CLR and WPF construction replaced at the boundary."""
    tree = ast.parse(COMMAND.read_text(encoding='utf-8'))
    classes = [node for node in tree.body if isinstance(node, ast.ClassDef)]
    api = {'Window': object, 'IExternalEventHandler': object, 'collections': collections,
           'json': json, 'ValidateRequest': core.ValidateRequest, 'Logger': mock.Mock(),
           'Visibility': NS(Collapsed=0, Visible=1), 'DB': object(), '_DATA_TABLE_WINDOWS': []}
    exec(compile(ast.Module(body=classes, type_ignores=[]), str(COMMAND), 'exec'), api)
    return api


class HostSafetyTests(unittest.TestCase):
    def setUp(self):
        self.api = host_classes()
        self.handler = self.api['DataTableEventHandler']()
        self.window = object.__new__(self.api['DataTableWindow'])
        self.window.DocumentKey = 'doc'
        self.window.Snapshot = {'rows': [{'id': '1'}]}
        self.window.IndexUri = 'file:///index.html'
        self.window.DocumentActive = True
        self.window.IsClosed = False
        self.window.Handler = self.handler
        self.window.ExternalEvent = NS(Raise=lambda: 'Accepted')
        self.window.SendResult = mock.Mock()
        self.handler.Window = self.window

    def send(self, message, source='file:///index.html'):
        self.window.OnWebMessage(None, NS(Source=source, WebMessageAsJson=json.dumps(message)))

    def test_queue_retains_distinct_requests_without_overwriting(self):
        self.send({'action': 'refresh', 'documentKey': 'doc'})
        self.send({'action': 'select', 'documentKey': 'doc', 'elementIds': ['1']})
        self.assertEqual(['refresh', 'select'], [request['action'] for request in self.handler.Requests])

    def test_foreign_page_never_queues_an_action(self):
        self.send({'action': 'refresh', 'documentKey': 'doc'}, 'https://example.com')
        self.assertEqual(0, len(self.handler.Requests))

    def test_inactive_document_is_rejected_before_queueing(self):
        self.window.DocumentActive = False
        self.send({'action': 'refresh', 'documentKey': 'doc'})
        self.assertEqual(0, len(self.handler.Requests))
        self.assertIn('Activate', self.window.SendResult.call_args.args[1])

    def test_denied_raise_removes_unaccepted_request(self):
        self.window.ExternalEvent = NS(Raise=lambda: 'Denied')
        self.send({'action': 'refresh', 'documentKey': 'doc'})
        self.assertEqual(0, len(self.handler.Requests))
        self.assertIn('Retry', self.window.SendResult.call_args.args[1])

    def test_document_switch_between_raise_and_execute_prevents_navigation(self):
        self.window.Document = NS(IsValidObject=True, Equals=lambda doc: False)
        self.handler.Requests.append({'action': 'select', 'elementIds': ['1']})
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=object())))
        self.assertEqual(0, len(self.handler.Requests))
        self.assertIn('Activate', self.window.SendResult.call_args.args[1])

    def test_model_change_marks_stale_without_replacing_snapshot(self):
        snapshot = self.window.Snapshot
        self.window.Document = object()
        self.window.SendContext = mock.Mock()
        self.window.OnDocumentChanged(None, NS(GetDocument=lambda: NS(Equals=lambda doc: True)))
        self.assertTrue(self.window.Outdated)
        self.assertIs(snapshot, self.window.Snapshot)
        self.window.SendContext.assert_called_once()

    def test_close_event_closes_only_invalid_bound_document(self):
        self.window.Close = mock.Mock()
        self.window.Document = NS(IsValidObject=True)
        self.window.OnDocumentClosed(None, None)
        self.window.Close.assert_not_called()
        self.window.Document.IsValidObject = False
        self.window.OnDocumentClosed(None, None)
        self.window.Close.assert_called_once()

    def test_close_detaches_events_disposes_resources_and_clears_retained_references(self):
        class Event:
            def __init__(self):
                self.handlers = []

            def __iadd__(self, handler):
                self.handlers.append(handler)
                return self

            def __isub__(self, handler):
                self.handlers.remove(handler)
                return self

        documentChanged, documentClosed, viewActivated = Event(), Event(), Event()
        self.window.Application = NS(Application=NS(DocumentChanged=documentChanged,
                                                    DocumentClosed=documentClosed),
                                     ViewActivated=viewActivated)
        self.window.Subscribe()
        coreEvents = [Event() for _ in range(4)]
        self.window.Core = NS(WebMessageReceived=coreEvents[0], NavigationStarting=coreEvents[1],
                              NavigationCompleted=coreEvents[2], NewWindowRequested=coreEvents[3])
        handlers = [self.window.OnWebMessage, self.window.OnNavigationStarting,
                    self.window.OnNavigationCompleted, self.window.OnNewWindowRequested]
        for event, handler in zip(coreEvents, handlers):
            event += handler
        self.window.Browser = NS(CoreWebView2InitializationCompleted=Event(), Dispose=mock.Mock())
        self.window.Browser.CoreWebView2InitializationCompleted += self.window.OnCoreInitialized
        self.window.Loaded = Event()
        self.window.Loaded += self.window.OnLoaded
        self.window.Closed = Event()
        self.window.Closed += self.window.HandleClosed
        self.window.ExternalEvent = NS(Dispose=mock.Mock())
        self.api['_DATA_TABLE_WINDOWS'].append(self.window)
        self.handler.Requests.append({'action': 'refresh'})
        self.window.HandleClosed(None, None)
        self.window.HandleClosed(None, None)  # Cleanup is idempotent.
        for event in [documentChanged, documentClosed, viewActivated, *coreEvents,
                      self.window.Loaded, self.window.Closed,
                      self.window.Browser.CoreWebView2InitializationCompleted]:
            self.assertEqual([], event.handlers)
        self.window.Browser.Dispose.assert_called_once()
        self.window.ExternalEvent.Dispose.assert_called_once()
        self.assertIsNone(self.handler.Window)
        self.assertEqual(0, len(self.handler.Requests))
        self.assertEqual([], self.api['_DATA_TABLE_WINDOWS'])


class DataTableWebView2Tests(loader_fixture.KeynoteWebView2Tests):
    """Run the shared compatibility scenarios against this button's loader."""
    def setUp(self):
        with mock.patch.object(loader_fixture, 'SCRIPT', LIBRARY / 'ffe_webview2.py'):
            super().setUp()

    def test_async_failure_includes_versions_paths_and_original_exception(self):
        coreAssembly, wpf = self.host.add_pair(loader_fixture.PYREVIT, loader_fixture.NEW)
        self.host.loaded.extend([coreAssembly, wpf])
        tree = ast.parse(COMMAND.read_text(encoding='utf-8'))
        window = next(node for node in tree.body if isinstance(node, ast.ClassDef)
                      and node.name == 'DataTableWindow')
        method = next(node for node in window.body if isinstance(node, ast.FunctionDef)
                      and node.name == 'OnCoreInitialized')
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(COMMAND), 'exec'), self.api)
        failure = mock.Mock()
        args = NS(IsSuccess=False, InitializationException='Method not found: constructor')
        self.api['OnCoreInitialized'](NS(IsClosed=False, ShowFailure=failure), None, args)
        message = failure.call_args.args[0]
        for value in ('Method not found: constructor', loader_fixture.NEW, coreAssembly.Location,
                      wpf.Location, 'Fully close and restart Revit'):
            self.assertIn(value, message)


if __name__ == '__main__':
    unittest.main()
