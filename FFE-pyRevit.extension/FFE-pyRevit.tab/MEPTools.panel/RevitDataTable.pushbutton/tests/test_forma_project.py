import os
import sys
import unittest
import uuid
from types import SimpleNamespace as NS
from unittest.mock import Mock

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'lib'))
from forma_project import NormalizeProjectId, ProjectNameFromPath, ReadFormaContext, CheckFormaProject, CreateFormaProject

PROJECT = '12345678-1234-4234-8234-123456789abc'
MODEL = '22345678-1234-4234-8234-123456789abc'
REPOSITORY = '32345678-1234-4234-8234-123456789abc'


class FormaTests(unittest.TestCase):
    def context(self):
        return {'available': True, 'formaProjectId': PROJECT, 'formaProjectName': '12345 - Building',
                'formaHubId': 'b.hub-id', 'cloudModelGuid': MODEL}

    def test_cloud_id_uses_data_management_project_id_not_an_unrelated_cloud_guid(self):
        doc = NS(IsModelInCloud=True, PathName='Autodesk Docs://12345 - Building/Project Files/Mechanical.rvt',
                 GetProjectId=lambda: 'b.' + PROJECT.upper(), GetHubId=lambda: 'b.hub-id',
                 GetCloudModelPath=lambda: NS(GetModelGUID=lambda: MODEL, GetProjectGUID=lambda: str(uuid.uuid4())))
        result = ReadFormaContext(doc)
        self.assertEqual(result['formaProjectId'], PROJECT)
        self.assertEqual(result['formaProjectName'], '12345 - Building')
        self.assertEqual(result['cloudModelGuid'], MODEL)

    def test_named_paths_and_missing_names_do_not_guess_from_local_names_or_guids(self):
        for prefix in ('BIM 360', 'Autodesk Docs', 'Autodesk Forma', 'Forma'):
            self.assertEqual(ProjectNameFromPath(prefix + '://12345%20-%20Building/Mechanical.rvt'), '12345 - Building')
        self.assertIsNone(ProjectNameFromPath('Autodesk Docs://' + PROJECT + '/Mechanical.rvt'))
        self.assertIsNone(ProjectNameFromPath('C:\\Projects\\Building\\Mechanical.rvt'))
        self.assertEqual(ProjectNameFromPath('Autodesk Docs://Caf%C3%A9/Mechanical.rvt'), 'Caf\u00e9')
        self.assertIsNone(ReadFormaContext(NS(IsModelInCloud=False)))
        result = ReadFormaContext(NS(IsModelInCloud=True, GetProjectId=lambda: ''))
        self.assertFalse(result['available'])

    def test_existing_project_auto_connects_without_mutations_and_lookup_errors_propagate(self):
        client = Mock()
        client.find_forma_project.return_value = {'project': {'projectId': REPOSITORY, 'name': 'Building'}}
        result = CheckFormaProject(client, self.context())
        self.assertEqual(result['project']['projectId'], REPOSITORY)
        client.ensure_forma_project.assert_not_called()
        client.find_forma_project.side_effect = RuntimeError('network unavailable')
        with self.assertRaises(RuntimeError): CheckFormaProject(client, self.context())
        client.ensure_forma_project.assert_not_called()

    def test_confirmed_creation_uses_detected_id_and_is_repeatable(self):
        client = Mock()
        client.find_forma_project.return_value = {'project': None}
        plan = CheckFormaProject(client, self.context())
        client.ensure_forma_project.assert_not_called()
        CreateFormaProject(client, plan, ' 12345 - Building ')
        CreateFormaProject(client, plan, '12345 - Building')
        expected = {'formaProjectId': PROJECT, 'formaHubId': 'b.hub-id', 'name': '12345 - Building'}
        self.assertEqual(client.ensure_forma_project.call_args_list[0].args[0], expected)
        self.assertEqual(client.ensure_forma_project.call_args_list[1].args[0], expected)
        with self.assertRaises(ValueError): CreateFormaProject(client, plan, '   ')

    def test_existing_project_link_is_bound_explicitly_and_conflicting_links_stop(self):
        client = Mock()
        client.find_forma_project.return_value = {'project': None}
        client.get_project.return_value = {'project': {'projectId': REPOSITORY, 'name': 'Existing'}}
        plan = CheckFormaProject(client, self.context(), REPOSITORY)
        client.get_project.assert_called_with({'projectId': REPOSITORY})
        CreateFormaProject(client, plan, 'Forma name')
        self.assertEqual(client.ensure_forma_project.call_args.args[0]['projectId'], REPOSITORY)
        client.find_forma_project.return_value = {'project': {'projectId': str(uuid.uuid4())}}
        with self.assertRaises(ValueError): CheckFormaProject(client, self.context(), REPOSITORY)
        client.find_forma_project.return_value = {'project': {'projectId': REPOSITORY}}
        self.assertIsNone(CheckFormaProject(client, self.context(), REPOSITORY)['existingProject'])

    def test_project_ids_normalize(self):
        self.assertEqual(NormalizeProjectId('b.' + PROJECT.upper()), PROJECT)
        with self.assertRaises(ValueError): NormalizeProjectId('b.00000000-0000-0000-0000-000000000000')


if __name__ == '__main__':
    unittest.main()
