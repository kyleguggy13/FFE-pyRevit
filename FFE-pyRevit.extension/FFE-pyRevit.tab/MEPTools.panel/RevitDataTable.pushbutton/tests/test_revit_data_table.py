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
savedSpec = importlib.util.spec_from_file_location('saved_tables_host_tests', LIBRARY / 'saved_tables.py')
saved = importlib.util.module_from_spec(savedSpec)
savedSpec.loader.exec_module(saved)
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
        self.Parameters = []

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


class Parameter:
    """Expose the parameter API, recording which values were actually read."""
    def __init__(self, valueId, name, storageType, value, dataTypeId='', text=None, isShared=False):
        self.Id = Id(valueId)
        self.Definition = NS(Name=name, GetDataType=lambda: NS(TypeId=dataTypeId))
        self.StorageType = storageType
        self.HasValue = value is not None
        self.IsShared = isShared
        self.AsString = mock.Mock(return_value=value)
        self.AsDouble = mock.Mock(return_value=value)
        self.AsInteger = mock.Mock(return_value=value)
        self.AsElementId = mock.Mock(return_value=Id(value))
        self.AsValueString = mock.Mock(return_value=text)


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
                     SpecTypeId=NS(Area='area', Length='length', Boolean=NS(YesNo=NS(TypeId='spec:yesno'))),
                     LabelUtils=NS(GetLabelForSpec=lambda spec: {
                         'spec:flow': 'Air Flow', 'spec:text': 'Text', 'spec:yesno': 'Yes/No'}[spec.TypeId]),
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

    def test_repository_capabilities_keep_references_portable_and_track_clear_support(self):
        reference = Parameter(-30, 'Reference', 'ElementId', 20, text='Level 1')
        reference.IsReadOnly = True
        self.elements[20].UniqueId = 'level-unique'
        notes = Parameter(31, 'Notes', 'String', None, isShared=True)
        notes.IsReadOnly = False
        notes.GUID = '00000000-0000-4000-8000-000000000001'
        self.elements[31] = NS(ShouldHideWhenNoValue=lambda: True)
        self.space.Parameters = [reference, notes]
        capabilities = {}
        core.BuildSnapshot(self.document, 'document-key', self.db, ['-30', '31'], capabilities)
        cells = capabilities[self.space.UniqueId]
        self.assertEqual({'writable': False, 'hasValue': True, 'referenceUniqueId': 'level-unique'},
                         cells['builtin:-30'])
        self.assertEqual({'writable': True, 'hasValue': False, 'canUnset': True},
                         cells['shared:' + notes.GUID])

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

    def test_parameter_catalog_unions_all_spaces_and_keeps_same_named_definitions_distinct(self):
        flow = Parameter(-200, 'Supply Airflow', 'Double', 125, 'spec:flow', '125 CFM')
        custom = Parameter(9007199254740993, 'Supply Airflow', 'String', 'custom', 'spec:text')
        other = Space(2)
        self.space.Parameters = [flow]
        other.Parameters = [flow, custom]
        self.document.spaces.append(other)
        snapshot = self.snapshot()
        self.assertEqual({'-200', '9007199254740993'}, {item['id'] for item in snapshot['parameters']})
        self.assertEqual({'Air Flow', 'Text'}, {item['dataType'] for item in snapshot['parameters']})
        self.assertEqual({}, snapshot['rows'][0]['parameterValues'])
        flow.AsDouble.assert_not_called()
        custom.AsString.assert_not_called()

    def test_selected_parameter_values_keep_units_zero_boolean_and_large_reference_ids(self):
        self.space.Parameters = [
            Parameter(-200, 'Supply Airflow', 'Double', 125, 'spec:flow', '125 CFM'),
            Parameter(201, 'Count', 'Integer', 0),
            Parameter(202, 'Conditioned', 'Integer', 1, 'spec:yesno'),
            Parameter(203, 'Reference', 'ElementId', 9007199254740993, text='Referenced Space'),
        ]
        snapshot = core.BuildSnapshot(self.document, 'document-key', self.db, ['-200', '201', '202', '203'])
        row = snapshot['rows'][0]
        self.assertEqual({'p-200': 125, 'p201': 0, 'p202': 1, 'p203': '9007199254740993'}, row['parameterValues'])
        self.assertEqual({'p-200': '125 CFM', 'p201': '0', 'p202': 'Yes', 'p203': 'Referenced Space'}, row['parameterTexts'])
        self.assertEqual(['-200', '201', '202', '203'], snapshot['selectedParameterIds'])
        json.dumps(snapshot, allow_nan=False)

    def test_unset_missing_and_unreadable_parameters_do_not_hide_spaces(self):
        broken = Parameter(201, 'Broken', 'Double', 1, text='1')
        broken.AsDouble.side_effect = ValueError('read failed')
        unset = Parameter(202, 'Unset', 'String', None)
        self.space.Parameters = [broken, unset]
        snapshot = core.BuildSnapshot(self.document, 'document-key', self.db, ['201', '202', 'removed'])
        self.assertEqual(0, snapshot['skippedCount'])
        self.assertEqual(['201', '202'], snapshot['selectedParameterIds'])
        row = snapshot['rows'][0]
        self.assertEqual({'p201': None, 'p202': None}, row['parameterValues'])
        self.assertEqual({'p201': '—', 'p202': '—'}, row['parameterTexts'])
        self.assertTrue(any('Broken unavailable' in warning for warning in row['warnings']))
        unset.AsString.assert_not_called()

    def test_parameter_missing_on_another_space_remains_empty(self):
        self.space.Parameters = [Parameter(201, 'Comments', 'String', '<literal text>', 'spec:text')]
        self.document.spaces.append(Space(2))
        snapshot = core.BuildSnapshot(self.document, 'document-key', self.db, ['201'])
        self.assertEqual('<literal text>', snapshot['rows'][0]['parameterTexts']['p201'])
        self.assertEqual({}, snapshot['rows'][1]['parameterValues'])

    def test_empty_or_unsupported_spec_uses_storage_type_label(self):
        parameter = Parameter(201, 'Associated Level', 'ElementId', 20)
        descriptor = core.BuildParameterDescriptor(parameter, self.db)
        self.assertEqual('Element reference', descriptor['dataType'])
        value, text = core.ReadParameterValue(parameter, descriptor, self.document, {})
        self.assertEqual(('20', 'Level 1'), (value, text))

    def test_starter_columns_match_builtin_ids_instead_of_parameter_names(self):
        for name, parameterId, field in (
                ('ROOM_LEVEL_ID', -1006916, 'level'), ('ROOM_NUMBER', -1006901, 'number'),
                ('ROOM_NAME', -1006900, 'name'), ('ROOM_AREA', -1006902, 'areaValue'),
                ('ROOM_HEIGHT', -1006920, 'heightValue'), ('ROOM_PHASE_ID', -1012112, 'phase'),
                ('ROOM_PHASE', -1012113, 'phase')):
            with self.subTest(name=name):
                setattr(self.db.BuiltInParameter, name, parameterId)
                parameter = Parameter(parameterId, 'Localized label', 'String', 'value')
                self.assertEqual(field, core.BuildParameterDescriptor(parameter, self.db)['baseField'])
        custom = Parameter(900, 'Name', 'String', 'custom')
        self.assertEqual('', core.BuildParameterDescriptor(custom, self.db)['baseField'])

    def test_parameter_origin_distinguishes_builtin_shared_project_global_and_other(self):
        class ProjectDefinition:
            pass

        class SharedDefinition(ProjectDefinition):
            pass

        class GlobalDefinition(ProjectDefinition):
            pass

        self.db.ParameterElement = ProjectDefinition
        self.db.SharedParameterElement = SharedDefinition
        self.db.GlobalParameter = GlobalDefinition
        self.elements.update({202: ProjectDefinition(), 203: GlobalDefinition(), 205: SharedDefinition()})
        self.space.Parameters = [
            Parameter(-200, 'Built-in', 'String', 'value'),
            Parameter(201, 'Shared', 'String', 'value', isShared=True),
            Parameter(202, 'Project', 'String', 'value'),
            Parameter(203, 'Global', 'String', 'value'),
            Parameter(204, 'Other', 'String', 'value'),
            Parameter(205, 'Shared definition', 'String', 'value'),
        ]
        snapshot = self.snapshot()
        self.assertEqual({'-200': 'Built-In', '201': 'Shared Parameter', '202': 'Project Parameter',
                          '203': 'Global Parameter', '204': 'Other', '205': 'Shared Parameter'},
                         {item['id']: item['parameterType'] for item in snapshot['parameters']})
        self.assertEqual(0, snapshot['skippedCount'])

    def test_unresolved_parameter_origin_does_not_guess_project_parameter(self):
        parameter = Parameter(201, 'Unknown', 'String', 'value')
        self.document.GetElement = mock.Mock(side_effect=ValueError('definition inaccessible'))
        self.assertEqual('Other', core.BuildParameterDescriptor(parameter, self.db, self.document)['parameterType'])
        parameter.Id = Id(-1)
        self.assertEqual('Other', core.GetParameterType(parameter, self.document, self.db))

    def test_parameter_and_filter_identities_are_persistent_and_not_display_names(self):
        builtin = Parameter(-200, 'Builtin', 'String', 'value')
        shared = Parameter(201, 'Same name', 'String', 'value', isShared=True)
        shared.GUID = '6CBA6914-A10E-4BE7-900A-B8A85FB041D8'
        project = Parameter(202, 'Same name', 'String', 'value')
        self.elements[202] = NS(UniqueId='project-definition-uid')
        self.elements[10].UniqueId = 'phase-uid'
        self.elements[20].UniqueId = 'level-uid'
        self.space.Parameters = [builtin, shared, project]
        snapshot = self.snapshot()
        self.assertEqual({'-200': 'builtin:-200', '201': 'shared:6cba6914-a10e-4be7-900a-b8a85fb041d8',
                          '202': 'project:project-definition-uid'},
                         {item['id']: item['identity'] for item in snapshot['parameters']})
        self.assertEqual('phase-uid', snapshot['phases'][0]['uniqueId'])
        self.assertEqual('level-uid', snapshot['levels'][0]['uniqueId'])

    def test_bad_parameter_definition_is_reported_without_skipping_row(self):
        parameter = Parameter(201, 'Good', 'String', 'value')
        broken = NS(Id=Id(202), StorageType='Double', Definition=None)
        self.space.Parameters = [broken, parameter]
        snapshot = self.snapshot()
        self.assertEqual(['201'], [item['id'] for item in snapshot['parameters']])
        self.assertEqual(0, snapshot['skippedCount'])
        self.assertTrue(any('Parameter definition unavailable' in warning for warning in snapshot['rows'][0]['warnings']))

    def test_nonfinite_selected_parameter_is_reported_without_invalid_json(self):
        self.space.Parameters = [Parameter(201, 'Invalid Number', 'Double', float('inf'), text='infinity')]
        snapshot = core.BuildSnapshot(self.document, 'document-key', self.db, ['201'])
        self.assertIsNone(snapshot['rows'][0]['parameterValues']['p201'])
        self.assertEqual(0, snapshot['skippedCount'])
        json.dumps(snapshot, allow_nan=False)

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
    def test_publish_requests_are_correlated_document_bound_and_require_valid_project(self):
        import uuid
        project = str(uuid.uuid4())
        message = {'action': 'publishReview', 'requestId': 'publish-1',
                   'documentKey': 'doc', 'projectId': project}
        self.assertEqual(project, core.ValidateRequest(message, 'doc', {})['projectId'])
        self.assertEqual('publishCatalog', core.ValidateRequest(dict(message, action='publishCatalog'), 'doc', {})['action'])
        for changes in ({'documentKey': 'other'}, {'requestId': None}, {'requestId': 'x' * 101},
                        {'projectId': None}, {'projectId': 'not-a-uuid'}, {'projectId': []}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                core.ValidateRequest(dict(message, **changes), 'doc', {})

    def test_sync_uses_only_the_review_identity_and_rejects_browser_rows(self):
        import uuid
        request = {'action': 'publishCommit', 'requestId': 'publish-2', 'documentKey': 'doc',
                   'reviewId': str(uuid.uuid4()), 'spaces': [{'lpd': 99}], 'projectId': str(uuid.uuid4())}
        result = core.ValidateRequest(request, 'doc', {})
        self.assertEqual({'action': 'publishCommit', 'requestId': 'publish-2', 'reviewId': request['reviewId']}, result)
        for reviewId in ('other', None, [], ''):
            with self.subTest(reviewId=reviewId), self.assertRaises(ValueError):
                core.ValidateRequest(dict(request, reviewId=reviewId), 'doc', {})

    def test_saved_table_requests_are_document_bound_and_validated_before_queueing(self):
        message = {'action': 'saveTable', 'requestId': 'table-1', 'tableId': '', 'revision': '',
                   'documentKey': 'doc', 'name': '  Review  ', 'layout': {}}
        self.assertEqual('Review', core.ValidateRequest(message, 'doc', {})['name'])
        for changes in ({'documentKey': 'other'}, {'name': ''}, {'name': None}, {'layout': []},
                        {'requestId': None}, {'tableId': 123}, {'revision': []}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                core.ValidateRequest(dict(message, **changes), 'doc', {})

    def test_parameter_selection_accepts_only_known_string_ids_and_preserves_refresh_selection(self):
        snapshot = {'parameters': [{'id': '1'}, {'id': '2'}], 'selectedParameterIds': ['2']}
        request = core.ValidateRequest({'action': 'refresh', 'documentKey': 'doc'}, 'doc', snapshot)
        self.assertEqual(['2'], request['parameterIds'])
        request = core.ValidateRequest({'action': 'refresh', 'documentKey': 'doc', 'parameterIds': ['1', '1']}, 'doc', snapshot)
        self.assertEqual(['1'], request['parameterIds'])
        for ids in ('1', [1], ['unknown'], None):
            with self.subTest(ids=ids), self.assertRaises(ValueError):
                core.ValidateRequest({'action': 'refresh', 'documentKey': 'doc', 'parameterIds': ids}, 'doc', snapshot)

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
           'PublishActions': core.PublishActions,
           'Visibility': NS(Collapsed=0, Visible=1), 'DB': object(), '_DATA_TABLE_WINDOWS': []}
    api.update({name: getattr(saved, name) for name in ('CaptureLayout', 'ResolveLayout', 'TableMetadata', 'ValidateLayout')})
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
        self.window.DocumentRevision = 0
        self.window.IsClosed = False
        self.window.Handler = self.handler
        self.window.ExternalEvent = NS(Raise=lambda: 'Accepted')
        self.window.SendResult = mock.Mock()
        self.handler.Window = self.window

    def send(self, message, source='file:///index.html'):
        self.window.OnWebMessage(None, NS(Source=source, WebMessageAsJson=json.dumps(message)))

    def test_preview_collects_fresh_values_in_api_context_without_cloud_or_storage_writes(self):
        import test_cloud_publish as fixtures
        self.api.update(CreateCloudClient=mock.Mock(), PrepareSync=fixtures.cloud.PrepareSync)
        client = self.api['CreateCloudClient'].return_value
        self.window.Document = NS(IsValidObject=True, Equals=lambda document: True)
        self.window.Application = NS(Application=NS(Username='Autodesk user'))
        self.window.CloudBusy = False
        self.window.PublishPlan = {'projects': [{'id': fixtures.PROJECT, 'name': 'Mechanical'}]}
        self.window.CloudStore = NS(Ensure=mock.Mock())
        self.window.StartCloudWork = mock.Mock()
        self.window.Send = mock.Mock()
        self.window.SendSnapshot = mock.Mock()
        self.window.SendContext = mock.Mock()
        self.api['BuildSnapshot'] = mock.Mock(return_value=fixtures.snapshot())
        self.window.Snapshot['selectedParameterIds'] = ['1']
        self.send({'action': 'publishReview', 'documentKey': 'doc', 'requestId': 'publish-1',
                   'projectId': fixtures.PROJECT})
        self.api['BuildSnapshot'].assert_not_called()
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=self.window.Document)))
        self.assertEqual(['1'], self.api['BuildSnapshot'].call_args.args[3])
        self.window.CloudStore.Ensure.assert_not_called()
        self.window.StartCloudWork.assert_not_called()
        self.assertEqual([], client.mock_calls)
        self.assertEqual(fixtures.PROJECT, self.window.PublishReview['request']['projectId'])
        self.assertEqual('publishResult', self.window.Send.call_args.args[0]['type'])

    def test_sync_rejects_stale_reviews_and_binds_project_before_deferred_http(self):
        import test_cloud_publish as fixtures
        self.api['SyncReview'] = fixtures.cloud.SyncReview
        self.window.CloudBusy = False
        review = fixtures.cloud.PrepareSync(fixtures.snapshot(), fixtures.PROJECT)
        review['documentRevision'] = 0
        self.window.PublishReview = review
        self.window.PublishClient = mock.Mock()
        self.window.CloudStore = NS(Ensure=mock.Mock())
        self.window.StartCloudWork = mock.Mock()
        request = {'action': 'publishCommit', 'reviewId': review['reviewId']}
        self.window.DocumentRevision = 1
        with self.assertRaisesRegex(ValueError, 'Revit model changed'):
            self.window.HandlePublishRequest(request)
        self.window.StartCloudWork.assert_not_called()
        self.window.CloudStore.Ensure.assert_not_called()
        self.window.DocumentRevision = 0
        self.window.HandlePublishRequest(request)
        self.window.CloudStore.Ensure.assert_called_once_with(fixtures.PROJECT)
        self.assertFalse(self.window.WritingTables)
        self.window.PublishClient.sync_revit_spaces.assert_not_called()
        self.window.StartCloudWork.call_args.args[1]()
        self.window.PublishClient.sync_revit_spaces.assert_called_once_with(review['request'])

    def test_cloud_worker_returns_on_dispatcher_and_ignores_closed_windows(self):
        # Run the worker body deterministically; its completion must still be marshalled.
        callbacks = []
        self.api.update(Action=lambda callback: callback,
                        threading=NS(Thread=lambda target: NS(start=target)))
        self.window.Dispatcher = NS(BeginInvoke=callbacks.append)
        self.window.Send = mock.Mock()
        request = {'action': 'publishCatalog', 'requestId': 'publish-1'}
        data = {'projects': []}
        self.window.StartCloudWork(request, lambda: data)
        self.assertTrue(self.window.CloudBusy)
        self.window.Send.assert_not_called()
        callbacks.pop()()
        self.assertFalse(self.window.CloudBusy)
        self.assertIs(data, self.window.PublishPlan)
        self.assertEqual('publishResult', self.window.Send.call_args.args[0]['type'])
        self.window.Send.reset_mock()
        self.window.StartCloudWork(request, lambda: data)
        self.window.IsClosed = True
        callbacks.pop()()
        self.window.Send.assert_not_called()

    def test_publish_failure_is_correlated_and_never_uses_saved_table_results(self):
        self.window.Send = mock.Mock()
        self.window.SendTableFailure({'action': 'publishReview', 'requestId': 'publish-1'}, ValueError('Reload projects'))
        response = self.window.Send.call_args.args[0]
        self.assertEqual(('publishResult', 'publish-1', False),
                         (response['type'], response['requestId'], response['success']))

    def test_uncertain_sync_keeps_review_for_retry_without_claiming_it_was_not_written(self):
        import test_cloud_publish as fixtures
        for error in (fixtures.rpc.RpcError('statement timeout', '57014', 500), OSError('Network lost'),
                      fixtures.rpc.RpcError('Unreadable response', 'INVALID_RESPONSE', 200)):
            self.api.update(Action=lambda callback: callback,
                            threading=NS(Thread=lambda target: NS(start=target)))
            self.window.Dispatcher = NS(BeginInvoke=lambda callback: callback())
            self.window.Send = mock.Mock()
            review = {'reviewId': 'review'}
            self.window.PublishReview = review
            self.window.StartCloudWork({'action': 'publishCommit', 'requestId': 'publish-2'},
                                       mock.Mock(side_effect=error))
            response = self.window.Send.call_args.args[0]
            self.assertTrue(response['retry'])
            self.assertIs(review, self.window.PublishReview)
            self.assertNotIn('No changes', response['message'])

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

    def test_saved_table_failure_is_correlated_and_preserves_the_working_snapshot(self):
        self.window.Document = NS(IsValidObject=True, Equals=lambda document: True)
        self.window.Send = mock.Mock()
        self.window.TableStore = NS(Write=mock.Mock(side_effect=ValueError('owned by another user')))
        self.api['CaptureLayout'] = mock.Mock(return_value={'portable': 'layout'})
        before = self.window.Snapshot
        self.send({'action': 'saveTable', 'requestId': 'table-1', 'documentKey': 'doc',
                   'name': 'Review', 'layout': {}})
        self.window.TableStore.Write.assert_not_called()
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=self.window.Document)))
        self.assertIs(before, self.window.Snapshot)
        self.assertFalse(self.window.WritingTables)
        response = self.window.Send.call_args.args[0]
        self.assertEqual(('tableResult', 'table-1', False),
                         (response['type'], response['requestId'], response['success']))
        self.assertIn('owned', response['message'])

    def test_saved_table_bridge_transacts_metadata_and_loads_fresh_space_values(self):
        import test_saved_tables as fixtures
        host = fixtures.StoreHost()
        host.document.IsValidObject = True
        host.document.Equals = lambda document: document is host.document
        self.window.Document = host.document
        self.window.TableStore = host.store
        self.window.Send = mock.Mock()
        self.window.Snapshot = dict(fixtures.snapshot(), documentKey='doc', rows=[])
        self.api['DB'] = host.db
        self.send({'action': 'saveTable', 'documentKey': 'doc', 'requestId': 'save-1',
                   'name': 'Review', 'layout': fixtures.runtime()})
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=host.document)))
        response = self.window.Send.call_args.args[0]
        self.assertTrue(response['success'])
        self.assertEqual(1, len(host.elements))
        tableId = response['table']['id']
        self.assertEqual(tableId, response['catalog']['savedTables'][0]['id'])
        self.assertTrue(host.transactions[0].disposed)
        host.definitions[fixtures.SHARED_GUID] = NS(Id=NS(Value=201))
        current = dict(fixtures.snapshot(), documentKey='doc', rows=[{'fresh': 'Space values'}])
        self.api['BuildSnapshot'] = mock.Mock(return_value=current)
        self.send({'action': 'loadTable', 'documentKey': 'doc', 'requestId': 'load-2', 'tableId': tableId})
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=host.document)))
        response = self.window.Send.call_args.args[0]
        self.assertTrue(response['success'])
        self.assertEqual([{'fresh': 'Space values'}], response['snapshot']['rows'])
        self.assertEqual(fixtures.runtime(), response['layout'])
        self.assertEqual(['201'], self.api['BuildSnapshot'].call_args.args[3])
        self.assertEqual(1, len(host.transactions), 'opening a saved view does not create transactions')

    def test_native_close_waits_for_browser_choice_and_closed_documents_bypass_prompt(self):
        self.window.Ready = True
        self.window.ClosePermit = False
        self.window.Document = NS(IsValidObject=True)
        self.window.Send = mock.Mock()
        args = NS(Cancel=False)
        self.window.HandleClosing(None, args)
        self.assertTrue(args.Cancel)
        self.window.Send.assert_called_once_with({'type': 'closeRequest'})
        self.window.Document.IsValidObject = False
        args = NS(Cancel=False)
        self.window.HandleClosing(None, args)
        self.assertFalse(args.Cancel)

    def test_confirmed_close_needs_no_document_action_even_when_another_model_is_active(self):
        self.window.DocumentActive = False
        self.window.Close = mock.Mock()
        self.send({'action': 'closeWindow', 'documentKey': 'doc'})
        self.window.Close.assert_called_once()
        self.assertTrue(self.window.ClosePermit)
        self.assertEqual(0, len(self.handler.Requests))

    def test_parameter_selection_is_read_inside_external_event_and_published_in_snapshot(self):
        self.window.Document = NS(IsValidObject=True, Equals=lambda document: True)
        self.window.SendSnapshot = mock.Mock()
        self.window.SendContext = mock.Mock()
        snapshot = {'rows': [], 'skippedCount': 0, 'selectedParameterIds': ['1']}
        buildSnapshot = mock.Mock(return_value=snapshot)
        self.api['BuildSnapshot'] = buildSnapshot
        self.handler.Requests.append({'action': 'refresh', 'parameterIds': ['1']})
        self.handler.Execute(NS(ActiveUIDocument=NS(Document=self.window.Document)))
        self.assertEqual(['1'], buildSnapshot.call_args.args[3])
        self.assertIs(snapshot, self.window.Snapshot)
        self.assertFalse(self.window.Outdated)
        self.window.SendSnapshot.assert_called_once()

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
        self.window.Closing = Event()
        self.window.Closing += self.window.HandleClosing
        self.window.Closed = Event()
        self.window.Closed += self.window.HandleClosed
        self.window.ExternalEvent = NS(Dispose=mock.Mock())
        self.api['_DATA_TABLE_WINDOWS'].append(self.window)
        self.handler.Requests.append({'action': 'refresh'})
        self.window.HandleClosed(None, None)
        self.window.HandleClosed(None, None)  # Cleanup is idempotent.
        for event in [documentChanged, documentClosed, viewActivated, *coreEvents,
                      self.window.Loaded, self.window.Closing, self.window.Closed,
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
