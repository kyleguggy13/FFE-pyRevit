import json
import os
import sys
import unittest
import uuid

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'lib'))
from ffe_supabase import Client, RpcError, OPERATIONS, _request
from revit_spaces import SpacesPayload, SyncSpaces, SpacesSnapshot, PushSpaces


class ClientTests(unittest.TestCase):
    def test_v3_envelope_and_attribution(self):
        calls = []
        def transport(url, headers, body):
            calls.append((url, headers, json.loads(body)))
            return 200, json.dumps({'ok': True})
        client = Client(actor_name='  Tester ', transport=transport)
        payload = {'projectId': str(uuid.uuid4()), 'name': 'Project'}
        client.create_project(payload)
        self.assertTrue(calls[0][0].endswith('/rest/v1/rpc/ffe_v3_create_project'))
        self.assertNotIn('Authorization', calls[0][1])
        body = calls[0][2]['p_request']
        self.assertEqual(body['schemaVersion'], 3)
        self.assertEqual(body['actor'], {'name': 'Tester'})
        self.assertNotIn('actor', payload)
        self.assertEqual(len(OPERATIONS), 26)

    def test_reads_without_actor_and_invalid_configuration(self):
        seen = []
        client = Client(transport=lambda url, headers, body: (seen.append(json.loads(body)) or (200, '{"ok":true,"items":[]}')))
        self.assertEqual(client.list_projects()['items'], [])
        self.assertNotIn('actor', seen[0]['p_request'])
        with self.assertRaises(ValueError): client.call('commit_draft', {})
        with self.assertRaises(ValueError): Client(url='http://example.com')
        with self.assertRaises(ValueError): Client(publishable_key='sb_secret_bad')
        with self.assertRaises(ValueError): Client(actor_name='x' * 201)
        odd = Client(transport=lambda url, headers, body: (200, '{"items":[]}'))
        with self.assertRaises(RpcError): odd.list_projects()

    def test_conflict_error_keeps_postgres_details(self):
        client = Client(transport=lambda url, headers, body: (400, '{"message":"Room changed","code":"40001","details":"reload"}'))
        with self.assertRaises(RpcError) as error: client.save_room({'room': {}})
        self.assertEqual(error.exception.code, '40001')
        self.assertEqual(error.exception.status, 400)
        self.assertEqual(error.exception.details, 'reload')

    def test_actual_cpython_http_transport_does_not_follow_redirects(self):
        from http.server import BaseHTTPRequestHandler, HTTPServer
        import threading
        requests = []
        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args): pass
            def do_POST(self):
                requests.append((self.path, self.headers.get('apikey'), self.headers.get('Authorization')))
                self.rfile.read(int(self.headers.get('Content-Length', '0')))
                self.send_response(302)
                self.send_header('Location', '/must-not-follow')
                self.end_headers()
                self.wfile.write(b'{"message":"redirect"}')
        server = HTTPServer(('127.0.0.1', 0), Handler)
        thread = threading.Thread(target=server.serve_forever)
        thread.daemon = True
        thread.start()
        try:
            status, body = _request('http://127.0.0.1:%d/rpc' % server.server_address[1], {'apikey': 'test-key'}, '{}')
            self.assertEqual(status, 302)
            self.assertEqual(json.loads(body)['message'], 'redirect')
            self.assertEqual(requests, [('/rpc', 'test-key', None)])
        finally:
            server.shutdown()
            server.server_close()
            thread.join()


class RevitSpacesTests(unittest.TestCase):
    def rows(self):
        return [{'id': '9007199254740993', 'uniqueId': 'space-unique', 'number': ' 002 ', 'name': 'Office',
                 'level': 'Level 2', 'areaValue': 120.000004, 'heightValue': 9.0, 'parameterValues': {'p1': 'x'}},
                {'uniqueId': 'space-2', 'number': '003', 'name': 'Unplaced', 'areaValue': None, 'heightValue': None}]

    def test_payload_sends_only_revit_owned_fields_with_string_ids(self):
        payload = SpacesPayload(self.rows(), 'office')
        self.assertEqual(payload['spaceTypeId'], 'office')
        first, second = payload['spaces']
        self.assertEqual(first, {'uniqueId': 'space-unique', 'elementId': '9007199254740993', 'number': '002',
                                 'name': 'Office', 'level': 'Level 2', 'area': 120.0, 'ceilingHeight': 9.0})
        self.assertNotIn('area', second)
        self.assertNotIn('ceilingHeight', second)

    def test_invalid_rows_are_rejected_before_any_request(self):
        for change in ({'uniqueId': ''}, {'id': 9007199254740993}, {'areaValue': -1}, {'areaValue': float('nan')},
                       {'areaValue': True}, {'number': 'x' * 51}):
            rows = self.rows()
            rows[0].update(change)
            with self.assertRaises(ValueError): SpacesPayload(rows)
        with self.assertRaises(ValueError): SpacesPayload([self.rows()[0], self.rows()[0]])

    def test_sync_sends_the_project(self):
        sent = []
        class Fake(object):
            def sync_revit_spaces(self, request):
                sent.append(request)
                return {'ok': True}
        project_id = str(uuid.uuid4())
        SyncSpaces(Fake(), project_id, self.rows())
        self.assertEqual(sent[0]['projectId'], project_id)
        self.assertNotIn('spaceTypeId', sent[0])


class PushSpacesTests(unittest.TestCase):
    def rows(self, count=3):
        return [{'uniqueId': 'uid-{0}'.format(i), 'id': str(1939673 + i), 'number': str(115 + i), 'name': 'ADL',
                 'level': 'Level 1', 'parameters': {
                     'Specified Supply Airflow': {'v': 2697.5308641975303, 's': '2698 CFM'},
                     'FFE_Space_Air Handling Unit_Mark': {'v': '34-AHU-1', 's': '34-AHU-1'},
                     'Occupiable': {'v': True}, 'Comments': {'v': None}}} for i in range(count)]

    def test_snapshot_keeps_every_parameter_and_rounds_floats(self):
        space = SpacesSnapshot(self.rows(1))[0]
        self.assertEqual(space['elementId'], '1939673')
        self.assertEqual(space['parameters']['Specified Supply Airflow'], {'v': 2697.530864, 's': '2698 CFM'})
        self.assertEqual(space['parameters']['Occupiable'], {'v': True})
        self.assertEqual(space['parameters']['Comments'], {'v': None})

    def test_invalid_snapshots_are_rejected_before_any_request(self):
        for change in ({'parameters': {'Area': 180.5}}, {'parameters': {'Area': {'v': float('nan')}}},
                       {'parameters': {'Area': {'v': [1]}}}, {'parameters': {'Area': {'v': 1, 'unit': 'SF'}}},
                       {'parameters': {'': {'v': 1}}}, {'parameters': ['Area']}, {'id': 1939673}, {'uniqueId': None}):
            rows = self.rows(1)
            rows[0].update(change)
            with self.assertRaises(ValueError): SpacesSnapshot(rows)
        with self.assertRaises(ValueError): SpacesSnapshot(self.rows(1) + self.rows(1))

    def test_push_sends_chunks_with_one_sync_id_and_only_the_last_is_final(self):
        sent = []
        class Fake(object):
            def push_revit_spaces(self, request):
                sent.append(request)
                return {'ok': True, 'created': len(request['spaces']), 'updated': 0, 'unchanged': 0,
                        'newlyMissing': ['old'] if request['final'] else [], 'sync': {'syncId': request['syncId']}}
        project_id = str(uuid.uuid4())
        catalog = [{'name': 'Specified Supply Airflow', 'storageType': 'Double', 'unit': 'CFM', 'group': 'Mechanical - Flow'}]
        result = PushSpaces(Fake(), project_id, self.rows(5), catalog, chunk_size=2)
        self.assertEqual([len(r['spaces']) for r in sent], [2, 2, 1])
        self.assertEqual([r['final'] for r in sent], [False, False, True])
        self.assertEqual(len(set(r['syncId'] for r in sent)), 1)
        self.assertEqual(sent[0]['parameters'][0], {'name': 'Specified Supply Airflow', 'storageType': 'Double', 'unit': 'CFM',
                                                    'group': 'Mechanical - Flow', 'isShared': False, 'isType': False})
        self.assertNotIn('parameters', sent[1])
        self.assertEqual(sent[0]['projectId'], project_id)
        self.assertEqual((result['created'], result['newlyMissing'], result['chunks']), (5, ['old'], 3))

    def test_empty_or_oversized_pushes_are_refused(self):
        class Fake(object):
            def push_revit_spaces(self, request): raise AssertionError('no request expected')
        with self.assertRaises(ValueError): PushSpaces(Fake(), 'p', [])
        with self.assertRaises(ValueError): PushSpaces(Fake(), 'p', self.rows(1), chunk_size=1001)
        with self.assertRaises(ValueError): PushSpaces(Fake(), 'p', self.rows(1), [{'name': 'A'}, {'name': 'A'}])


if __name__ == '__main__':
    unittest.main()
