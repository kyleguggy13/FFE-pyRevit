# -*- coding: utf-8 -*-
"""Purpose: verify all-parameter extraction independently of visible table columns.
Behaviors: display units, duplicate names, types, unset values and exact references.
Revit API: 2025/2026 calls are exercised through the existing fake API fixtures.
Design: keep real collector logic, substituting only native Revit boundaries.
"""
import json
import pathlib
import sys
import unittest
from types import SimpleNamespace as NS
from unittest.mock import Mock

import test_revit_data_table as fixture

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / 'lib'))
from revit_parameters import CollectSyncParameters


class FullParameterTests(unittest.TestCase):
    def setUp(self):
        self.host = fixture.SpaceSnapshotTests()
        self.host.setUp()
        self.document, self.db, self.space = self.host.document, self.host.db, self.host.space
        self.space.GetTypeId = lambda: fixture.Id(-1)
        self.unit = NS(TypeId='display-unit')
        self.document.GetUnits = lambda: NS(GetFormatOptions=lambda spec: NS(
            GetUnitTypeId=lambda: self.unit, GetSymbolTypeId=lambda: NS(TypeId='cfm')))
        self.db.UnitUtils = NS(IsMeasurableSpec=lambda spec: spec.TypeId == 'spec:flow',
                               ConvertFromInternalUnits=Mock(side_effect=lambda value, unit: value * 60))
        self.db.LabelUtils.GetLabelForSymbol = lambda symbol: 'CFM'
        self.db.LabelUtils.GetLabelForUnit = lambda unit: 'Cubic feet per minute'
        self.rows = [{'id': str(self.space.Id.Value), 'uniqueId': self.space.UniqueId}]

    def test_hidden_instance_and_type_parameters_are_collected_without_table_selection(self):
        flow = fixture.Parameter(-30, 'Specified Supply Airflow', 'Double', 2, 'spec:flow', '120 CFM')
        text = fixture.Parameter(31, 'Hidden notes', 'String', ' leading and trailing ')
        yes = fixture.Parameter(32, 'Occupiable', 'Integer', 1, 'spec:yesno')
        empty = fixture.Parameter(33, 'Unset', 'String', None)
        reference = fixture.Parameter(34, 'Reference', 'ElementId', 9007199254740995, text='Terminal')
        typeParameter = fixture.Parameter(35, 'Type description', 'String', 'Type data')
        self.space.Parameters = [flow, text, yes, empty, reference]
        self.space.GetTypeId = lambda: fixture.Id(101)
        self.host.elements[101] = NS(Parameters=[typeParameter])
        data = CollectSyncParameters(self.document, self.rows, self.db)
        values = data['spaces'][self.space.UniqueId]
        self.assertEqual({'v': 120.0, 's': '120 CFM'}, values['Specified Supply Airflow'])
        self.assertEqual(' leading and trailing ', values['Hidden notes']['v'])
        self.assertTrue(values['Occupiable']['v'])
        self.assertEqual({'v': None, 's': ''}, values['Unset'])
        self.assertEqual('9007199254740995', values['Reference']['v'])
        self.assertEqual('Type data', values['Type description']['v'])
        self.assertEqual('CFM', next(item for item in data['catalog'] if item['name'] == 'Specified Supply Airflow')['unit'])
        self.assertTrue(next(item for item in data['catalog'] if item['name'] == 'Type description')['isType'])
        self.assertEqual([], data['issues'])
        json.dumps(data, allow_nan=False)

    def test_duplicate_names_keep_every_value_and_stay_stable_for_group_subsets(self):
        first = fixture.Parameter(-30, 'Comments', 'String', 'built in')
        other = fixture.Parameter(31, 'Comments', 'String', 'project')
        otherSpace = fixture.Space(2)
        otherSpace.Parameters = [other]
        otherSpace.GetTypeId = lambda: fixture.Id(-1)
        self.space.Parameters = [first]
        self.document.spaces.append(otherSpace)
        rows = self.rows + [{'id': '2', 'uniqueId': otherSpace.UniqueId}]
        full = CollectSyncParameters(self.document, rows, self.db)
        subset = CollectSyncParameters(self.document, rows[1:], self.db, rows)
        self.assertEqual({'Comments'}, set(full['spaces'][self.space.UniqueId]))
        self.assertEqual(full['spaces'][otherSpace.UniqueId], subset['spaces'][otherSpace.UniqueId])
        self.assertTrue(next(iter(subset['spaces'][otherSpace.UniqueId])).startswith('Comments [instance '))
        # Reading definitions for other Spaces must not read their parameter values.
        first.AsString.reset_mock()
        CollectSyncParameters(self.document, rows[1:], self.db, rows)
        first.AsString.assert_not_called()

    def test_unreadable_values_are_explicit_and_whole_unreadable_spaces_are_skipped(self):
        broken = fixture.Parameter(31, 'Unreadable', 'String', 'before')
        broken.AsString.side_effect = ValueError('API failure')
        nan = fixture.Parameter(32, 'Invalid double', 'Double', float('nan'))
        self.space.Parameters = [broken, nan]
        data = CollectSyncParameters(self.document, self.rows, self.db)
        self.assertEqual(2, data['unavailableCount'])
        self.assertEqual({'v': None, 's': 'Unavailable'}, data['spaces'][self.space.UniqueId]['Unreadable'])
        missing = self.rows + [{'id': '999', 'uniqueId': 'deleted'}]
        data = CollectSyncParameters(self.document, missing, self.db)
        self.assertEqual(1, data['skippedCount'])
        self.assertNotIn('deleted', data['spaces'])
        self.assertTrue(any('Space skipped' in issue for issue in data['issues']))

    def test_long_text_is_bounded_with_a_warning_and_empty_text_is_preserved(self):
        self.space.Parameters = [fixture.Parameter(31, 'Notes', 'String', 'x' * 2001),
                                 fixture.Parameter(32, 'Empty', 'String', '')]
        data = CollectSyncParameters(self.document, self.rows, self.db)
        self.assertEqual(2000, len(data['spaces'][self.space.UniqueId]['Notes']['v']))
        self.assertEqual('', data['spaces'][self.space.UniqueId]['Empty']['v'])
        self.assertTrue(any('truncated' in issue for issue in data['issues']))


if __name__ == '__main__':
    unittest.main()
