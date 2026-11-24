"""Exercise the real startup methods without constructing WPF or touching Revit."""
import ast
import collections
import importlib.util
import json
import pathlib
import sys
import unittest
import uuid
from types import SimpleNamespace as NS
from unittest.mock import Mock

BUTTON = pathlib.Path(__file__).resolve().parents[1]
LIBRARY = BUTTON / 'lib'
sys.path.insert(0, str(LIBRARY))
import forma_project


def host():
    tree = ast.parse((BUTTON / 'script.py').read_text(encoding='utf-8'))
    api = {'Window': object, 'IExternalEventHandler': object, 'collections': collections,
           'json': json, 'uuid': uuid, 'MessageBox': NS(Show=Mock(return_value='No')),
           'MessageBoxButton': NS(YesNo='YesNo'), 'MessageBoxImage': NS(Question='Question', Warning='Warning'),
           'MessageBoxResult': NS(Yes='Yes', No='No'), 'forms': NS(ask_for_string=Mock()),
           'CheckFormaProject': forma_project.CheckFormaProject, 'CreateFormaProject': forma_project.CreateFormaProject}
    classes = [node for node in tree.body if isinstance(node, ast.ClassDef)]
    exec(compile(ast.Module(body=classes, type_ignores=[]), str(BUTTON / 'script.py'), 'exec'), api)
    window = object.__new__(api['DataTableWindow'])
    window.IsClosed = False
    window.DocumentActive = True
    window.PendingFormaResponse = None
    window.FormaProject = None
    window.StartupClient = Mock()
    window.StartupLinkedProjectId = None
    window.FormaContext = {'available': True, 'formaProjectId': str(uuid.uuid4()),
                           'formaProjectName': '12345 - Building', 'formaHubId': 'b.hub', 'cloudModelGuid': str(uuid.uuid4())}
    window.SendResult = Mock()
    window.SendContext = Mock()
    window.StartCloudWork = Mock()
    return window, api


class StartupTests(unittest.TestCase):
    def response(self, window, project=None):
        return {'success': True, 'data': {'project': project, 'existingProject': None, 'forma': window.FormaContext}}

    def test_missing_project_prompts_and_no_does_not_write(self):
        window, api = host()
        window.OnFormaLookup(self.response(window))
        api['MessageBox'].Show.assert_called_once()
        self.assertIn(window.FormaContext['formaProjectName'], api['MessageBox'].Show.call_args.args[1])
        self.assertIn(window.FormaContext['formaProjectId'], api['MessageBox'].Show.call_args.args[1])
        window.StartupClient.ensure_forma_project.assert_not_called()
        window.StartCloudWork.assert_not_called()

    def test_yes_defers_creation_to_worker_and_records_shared_destination(self):
        window, api = host(); api['MessageBox'].Show.return_value = 'Yes'
        window.OnFormaLookup(self.response(window))
        window.StartupClient.ensure_forma_project.assert_not_called()
        request, work, complete = window.StartCloudWork.call_args.args
        self.assertEqual(request['action'], 'formaCreate')
        project = {'projectId': str(uuid.uuid4()), 'name': '12345 - Building'}
        window.StartupClient.ensure_forma_project.return_value = {'ok': True, 'created': True, 'project': project}
        result = work()
        payload = window.StartupClient.ensure_forma_project.call_args.args[0]
        self.assertEqual(payload['formaProjectId'], window.FormaContext['formaProjectId'])
        self.assertNotIn('requestId', payload)
        self.assertEqual(payload['name'], window.FormaContext['formaProjectName'])
        complete({'success': True, 'data': result})
        self.assertEqual(window.FormaProject, project)

    def test_existing_project_connects_without_prompt_or_creation(self):
        window, api = host(); project = {'projectId': str(uuid.uuid4()), 'name': 'Shared'}
        window.OnFormaLookup(self.response(window, project))
        self.assertEqual(window.FormaProject, project)
        api['MessageBox'].Show.assert_not_called()
        window.StartCloudWork.assert_not_called()

    def test_missing_endpoint_and_network_errors_never_become_create_prompts(self):
        for code in ('PGRST202', 'NETWORK_ERROR'):
            window, api = host()
            window.OnFormaLookup({'success': False, 'code': code, 'message': 'Connection failed'})
            api['MessageBox'].Show.assert_not_called()
            window.StartCloudWork.assert_not_called()
            self.assertEqual(window.SendResult.call_args.args[0], 'warning')

    def test_unavailable_name_is_requested_after_yes_without_using_revit_project_information(self):
        window, api = host(); window.FormaContext['formaProjectName'] = None
        api['MessageBox'].Show.return_value = 'Yes'
        api['forms'].ask_for_string.return_value = 'Name from Forma'
        window.OnFormaLookup(self.response(window))
        window.StartCloudWork.call_args.args[1]()
        self.assertEqual(window.StartupClient.ensure_forma_project.call_args.args[0]['name'], 'Name from Forma')

    def test_inactive_and_closed_documents_do_not_prompt(self):
        window, api = host(); window.DocumentActive = False
        response = self.response(window)
        window.OnFormaLookup(response)
        self.assertIs(window.PendingFormaResponse, response)
        api['MessageBox'].Show.assert_not_called()
        window.DocumentActive = True; window.IsClosed = True
        window.OnFormaLookup(response)
        api['MessageBox'].Show.assert_not_called()

    def test_startup_lookup_is_once_per_window_and_has_no_revit_access_on_worker(self):
        window, api = host()
        window.BeginProjectLookup(); window.BeginProjectLookup()
        self.assertEqual(window.StartCloudWork.call_count, 1)
        window.StartupClient.find_forma_project.return_value = {'project': None}
        result = window.StartCloudWork.call_args.args[1]()
        self.assertIsNone(result['project'])
        window.StartupClient.ensure_forma_project.assert_not_called()

    def test_transient_creation_retry_reuses_original_request_id(self):
        window, api = host(); api['MessageBox'].Show.return_value = 'Yes'
        window.OnFormaLookup(self.response(window))
        original = window.StartCloudWork.call_args.args
        original[2]({'success': False, 'retry': True, 'message': 'Network lost'})
        retry = window.StartCloudWork.call_args.args
        self.assertIs(retry[0], original[0]); self.assertIs(retry[1], original[1])

    def test_sync_catalog_uses_forma_destination_and_project_binding(self):
        window, api = host()
        api['LoadCatalog'] = Mock(return_value={})
        api['CreateCloudClient'] = Mock()
        window.Application = NS(Application=object())
        window.CloudBusy = False
        project_id = str(uuid.uuid4())
        window.CloudStore = NS(Read=lambda: project_id)
        window.FormaProject = {'projectId': project_id, 'name': 'Building'}
        window.HandlePublishRequest({'action': 'publishCatalog', 'requestId': 'catalog-1'})
        window.StartCloudWork.call_args.args[1]()
        self.assertEqual(api['LoadCatalog'].call_args.args[1:], (project_id, project_id, window.FormaContext))


if __name__ == '__main__':
    unittest.main()
