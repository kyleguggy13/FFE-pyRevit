# -*- coding: utf-8 -*-
"""FFE shared RPC client for IronPython 2.7 / CPython, with no Revit dependency.

SQL is owned by FFE-Mechanical/supabase/migrations. Uses the shared publishable
key without individual sign-in. Actor metadata is client-reported attribution.
Perform network requests away from the Revit UI/ExternalEvent thread.
"""
import json
import sys
import uuid

PROJECT_URL = "https://ncxylphudovrmvdgqsol.supabase.co"
PUBLISHABLE_KEY = "sb_publishable_VFbJW7ZQQh0Ro254lfXXvw_xesXNxTK"


OPERATIONS = ('create_project', 'register_model', 'list_projects', 'list_models', 'get_catalog', 'submit_collection', 'compare_collection',
              'find_forma_project', 'ensure_forma_project',
              'create_draft', 'update_draft', 'get_draft', 'list_drafts', 'discard_draft', 'resolve_draft', 'commit_draft',
              'get_state', 'get_history', 'get_commit', 'compare_commits', 'prepare_pull', 'acknowledge_pull', 'get_pull',
              'get_sync', 'save_layout', 'rename_layout', 'delete_layout', 'list_layouts', 'get_layout', 'layout_history')
WRITES = frozenset(('create_project', 'ensure_forma_project', 'register_model', 'submit_collection', 'create_draft', 'update_draft', 'resolve_draft',
                   'discard_draft', 'commit_draft', 'prepare_pull', 'acknowledge_pull', 'save_layout', 'rename_layout', 'delete_layout'))
try:
    STRING_TYPES = (basestring,)
except NameError:
    STRING_TYPES = (str,)


class RpcError(Exception):
    def __init__(self, message, code=None, status=None, details=None):
        Exception.__init__(self, message)
        self.code = code
        self.status = status
        self.details = details


def _request(url, headers, payload):
    """Return (HTTP status, decoded JSON text); HTTP errors remain structured."""
    if sys.platform == "cli":
        from System.IO import StreamReader
        from System.Net import WebRequest, WebException
        from System.Text import Encoding
        request = WebRequest.Create(url)
        request.Method = "POST"
        request.ContentType = "application/json"
        request.Timeout = 30000
        request.ReadWriteTimeout = 30000
        request.AllowAutoRedirect = False
        for name, value in headers.items():
            request.Headers[name] = value
        data = Encoding.UTF8.GetBytes(payload)
        request.ContentLength = data.Length
        stream = request.GetRequestStream()
        try:
            stream.Write(data, 0, data.Length)
        finally:
            stream.Close()
        try:
            response = request.GetResponse()
        except WebException as error:
            response = error.Response
            if response is None:
                raise RpcError("Could not connect to Supabase.", "NETWORK_ERROR")
        try:
            reader = StreamReader(response.GetResponseStream())
            try:
                return int(response.StatusCode), reader.ReadToEnd()
            finally:
                reader.Close()
        finally:
            response.Close()
    from urllib.request import Request, build_opener, HTTPRedirectHandler
    from urllib.error import HTTPError

    class NoRedirect(HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, hdrs, newurl):
            return None

    request = Request(url, data=payload.encode("utf-8"), headers=dict(headers, **{"Content-Type": "application/json"}))
    try:
        response = build_opener(NoRedirect).open(request, timeout=30)
    except HTTPError as error:
        response = error
    try:
        return response.code, response.read().decode("utf-8")
    finally:
        response.close()


class Client(object):
    def __init__(self, url=PROJECT_URL, publishable_key=PUBLISHABLE_KEY,
                 client_id=None, actor_name="", user_id="", source="revit", transport=None):
        try:
            from urllib.parse import urlparse
        except ImportError:
            from urlparse import urlparse
        parsed = urlparse(url)
        if (parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password
                or parsed.path not in ("", "/") or parsed.query or parsed.fragment):
            raise ValueError("Use the HTTPS Supabase project URL.")
        if not publishable_key.startswith("sb_publishable_"):
            raise ValueError("Use a Supabase publishable key.")
        if (not isinstance(actor_name, STRING_TYPES) or len(actor_name) > 200
                or not isinstance(user_id, STRING_TYPES) or len(user_id) > 200 or source not in ("web", "revit")):
            raise ValueError("Invalid client attribution.")
        if client_id is not None:
            client_id = str(uuid.UUID(client_id))
        self.url = url.rstrip("/")
        self.key = publishable_key
        self.actor = {"clientId": client_id, "name": actor_name, "userId": user_id, "source": source}
        self.transport = transport or _request

    def rpc(self, name, args=None):
        import re
        if not re.match(r"^ffe_v2_[a-z_]+\Z", name):
            raise ValueError("Use a version 2 repository RPC.")
        status, raw = self.transport(self.url + "/rest/v1/rpc/" + name,
                                     {"apikey": self.key},
                                     json.dumps(args or {}, allow_nan=False))
        try:
            result = json.loads(raw)
        except (ValueError, TypeError):
            raise RpcError("Supabase returned an unreadable response.", "INVALID_RESPONSE", status)
        if not 200 <= status < 300:
            error = result if isinstance(result, dict) else {}
            raise RpcError(error.get("message", "The shared data request failed."), error.get("code"), status, error.get("details"))
        return result

    def call(self, action, data=None):
        if action not in OPERATIONS or (data is not None and not isinstance(data, dict)):
            raise ValueError("Invalid repository request.")
        request = dict(data or {})
        request["schemaVersion"] = 2
        if action in WRITES:
            if not self.actor["clientId"] or not self.actor["name"].strip():
                raise ValueError("Writes require a persistent client_id and actor_name.")
            request["actor"] = dict(self.actor)
            request["requestId"] = str(uuid.UUID(request["requestId"])) if request.get("requestId") else str(uuid.uuid4())
        try:
            return self.rpc("ffe_v2_" + action, {"p_request": request})
        except Exception as error:
            error.request_id = request.get("requestId")
            raise

    def publish_snapshot(self, *args, **kwargs):
        raise ValueError("Snapshot overwrites were retired. Submit a collection, stage a draft, and commit it.")

    def get_snapshot(self, *args, **kwargs):
        raise ValueError("Use get_state for versioned repository values.")


def _operation(action):
    def invoke(self, request=None):
        return self.call(action, request)
    invoke.__name__ = action
    invoke.__doc__ = "Call ffe_v2_" + action + " with a version 2 request dictionary."
    return invoke


for _action in OPERATIONS:
    setattr(Client, _action, _operation(_action))
