"""View/sheet analytics behavior with Revit objects replaced at the API boundary.

Run: python -m unittest discover -s tests -p 'test_keynote_analytics.py'
"""
import json
import types
import unittest

from test_keynote_annotation_library import load_functions


class Parameter:
    def __init__(self, value):
        self.value = value

    def AsString(self):
        return self.value


class Element:
    def __init__(self, identity, name='', **parameters):
        self.Id = types.SimpleNamespace(Value=identity)
        self.Name = name
        self.parameters = parameters

    def get_Parameter(self, parameter_id):
        if parameter_id in self.parameters:
            return Parameter(self.parameters[parameter_id])
        return None


class Collector(list):
    def OfClass(self, element_class):
        return self

    def WhereElementIsNotElementType(self):
        return self


class KeynoteAnalyticsTests(unittest.TestCase):
    def setUp(self):
        self.api = load_functions()
        self.elements = {}
        self.viewports = []
        self.doc = types.SimpleNamespace(GetElement=lambda identity: self.elements.get(identity.Value))
        self.api.update({
            'BuiltInParameter': types.SimpleNamespace(
                SYMBOL_NAME_PARAM='typeName', VIEW_DESCRIPTION='title', VIEW_TYPE='shortType',
                VIEW_FAMILY_AND_TYPE_SCHEDULES='familyType',
                VIEWPORT_DETAIL_NUMBER='detail'),
            'Viewport': object,
            'FilteredElementCollector': lambda doc: Collector(self.viewports),
        })
        self.sheet = self.add_sheet(1933263, 'AF721', 'WALL FINISH DETAILS')
        self.next_element_id = 9876543210

    def add_sheet(self, identity, number, name):
        sheet = Element(identity, name, familyType='Sheet: Drawing')
        sheet.SheetNumber = number
        self.elements[identity] = sheet
        return sheet

    def add_view(self, identity, name, title='', family_type='Drafting View: Detail'):
        view = Element(identity, name, title=title, familyType=family_type, shortType='Detail')
        view.Document = self.doc
        self.elements[identity] = view
        return view

    def add_viewport(self, view, sheet, detail):
        viewport = Element(len(self.viewports) + 1, detail=detail)
        viewport.ViewId = view.Id
        viewport.SheetId = sheet.Id
        self.viewports.append(viewport)

    def collect(self, placements):
        lookup = self.api['build_view_sheet_lookup'](self.doc)
        rows = {}
        for view, source_type in placements:
            annotation = Element(self.next_element_id)
            self.next_element_id += 1
            annotation.OwnerViewId = view.Id
            self.api['record_keynote_analytics_placement'](
                rows, {'A': {'text': 'Finish transition'}}, 'A', source_type,
                annotation, self.doc, lookup)
        return self.api['finalize_keynote_analytics_rows'](rows)[0]

    def test_requested_structure_counts_each_view_and_preserves_unicode(self):
        views = []
        expected_views = {}
        for index, identity in enumerate((1933996, 1932163, 1932540, 1932547,
                                          1932554, 1933282, 1933289, 1933296), 1):
            title = 'TRANSITION STRIP \u2013 DETAIL {}'.format(index)
            view = self.add_view(identity, 'Arch_' + title, title)
            self.add_viewport(view, self.sheet, str(index))
            views.append(view)
            expected_views[str(identity)] = {
                'View Name': 'Arch_' + title, 'Title on Sheet': title,
                'Family and Type': 'Drafting View: Detail',
                'Detail Number': str(index), 'count': 2,
            }
        result = self.collect([(view, 'genericAnnotation') for view in views for _ in range(2)])
        # Round-trip the payload sent to Supabase, including view IDs as JSON object keys.
        sheet = json.loads(json.dumps(result['sheets']))[0]
        self.assertEqual({
            'id': '1933263', 'name': 'WALL FINISH DETAILS', 'number': 'AF721',
            'count': 16, 'viewIds': expected_views, 'userKeynoteCount': 0,
            'genericAnnotationCount': 16,
        }, sheet)
        self.assertEqual(16, result['placedCount'])
        self.assertEqual(1, result['sheetCount'])
        self.assertEqual(0, result['unsheetedCount'])

    def test_same_view_on_two_sheets_has_independent_detail_numbers_and_counts(self):
        view = self.add_view(10, 'Common legend')
        other_sheet = self.add_sheet(20, 'AF722', 'OTHER DETAILS')
        self.add_viewport(view, self.sheet, '8')
        self.add_viewport(view, other_sheet, '3')
        result = self.collect([(view, 'genericAnnotation'), (view, 'userKeynote')])
        by_number = {sheet['number']: sheet for sheet in result['sheets']}
        self.assertEqual('8', by_number['AF721']['viewIds']['10']['Detail Number'])
        self.assertEqual('3', by_number['AF722']['viewIds']['10']['Detail Number'])
        for sheet in by_number.values():
            self.assertEqual(2, sheet['viewIds']['10']['count'])
            self.assertEqual(2, sheet['count'])
            self.assertEqual(1, sheet['userKeynoteCount'])
            self.assertEqual(1, sheet['genericAnnotationCount'])
        self.assertEqual(2, result['placedCount'])
        self.assertEqual(2, result['sheetCount'])

    def test_views_with_same_name_stay_separate_and_titles_fall_back_to_view_name(self):
        first = self.add_view(10, 'Same name')
        second = self.add_view(11, 'Same name', 'Alternate title')
        self.add_viewport(first, self.sheet, '1')
        self.add_viewport(second, self.sheet, '2')
        result = self.collect([(first, 'userKeynote'), (second, 'userKeynote'), (second, 'userKeynote')])
        views = result['sheets'][0]['viewIds']
        self.assertEqual({'10', '11'}, set(views))
        self.assertEqual(1, views['10']['count'])
        self.assertEqual(2, views['11']['count'])
        self.assertEqual('Same name', views['10']['Title on Sheet'])
        self.assertEqual('Alternate title', views['11']['Title on Sheet'])

    def test_missing_family_parameter_uses_view_type_and_missing_detail_stays_blank(self):
        view = self.add_view(10, 'Finish detail', family_type='')
        view_type = Element(30, 'Detail')
        view_type.FamilyName = 'Drafting View'
        self.elements[30] = view_type
        view.GetTypeId = lambda: view_type.Id
        self.add_viewport(view, self.sheet, '')
        detail = self.collect([(view, 'genericAnnotation')])['sheets'][0]['viewIds']['10']
        self.assertEqual('Drafting View: Detail', detail['Family and Type'])
        self.assertEqual('', detail['Detail Number'])
        view.GetTypeId = lambda: types.SimpleNamespace(Value=999)
        detail = self.collect([(view, 'genericAnnotation')])['sheets'][0]['viewIds']['10']
        self.assertEqual('', detail['Family and Type'])
        self.assertEqual(1, detail['count'])

    def test_unsheeted_view_keeps_placement_metadata_without_inventing_a_sheet(self):
        view = self.add_view(10, 'Unsheeted detail', 'Finish detail')
        result = self.collect([(view, 'genericAnnotation')])
        self.assertEqual([], result['sheets'])
        self.assertEqual(1, result['unsheetedCount'])
        self.assertEqual('Finish detail', result['placements'][0]['view']['titleOnSheet'])
        self.assertEqual('9876543210', result['placements'][0]['elementId'])

    def test_annotations_directly_on_sheet_use_sheet_as_owner_and_no_detail_number(self):
        result = self.collect([(self.sheet, 'genericAnnotation')])
        sheet = result['sheets'][0]
        self.assertEqual(1, sheet['count'])
        self.assertEqual({
            'View Name': 'WALL FINISH DETAILS', 'Title on Sheet': 'WALL FINISH DETAILS',
            'Family and Type': 'Sheet: Drawing', 'Detail Number': '', 'count': 1,
        }, sheet['viewIds']['1933263'])


if __name__ == '__main__':
    unittest.main()
