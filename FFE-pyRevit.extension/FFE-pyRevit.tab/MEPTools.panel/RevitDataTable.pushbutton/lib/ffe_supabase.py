# -*- coding: utf-8 -*-
"""Loadwise (FFE Mechanical) RPC client for IronPython 2.7 / CPython, with no Revit dependency.

SQL is owned by FFE-Mechanical/supabase/migrations (API v3). Uses the shared
publishable key without individual sign-in. The actor name is client-reported
attribution. Perform network requests away from the Revit UI/ExternalEvent thread.
"""
import json
import sys

PROJECT_URL = "https://ncxylphudovrmvdgqsol.supabase.co"
PUBLISHABLE_KEY = "sb_publishable_VFbJW7ZQQh0Ro254lfXXvw_xesXNxTK"


OPERATIONS = ('list_projects', 'get_project', 'create_project', 'update_project', 'save_air_system',
              'delete_air_system', 'find_forma_project', 'ensure_forma_project', 'list_space_types', 'list_rooms',
              'save_room', 'delete_room', 'sync_revit_spaces', 'import_hap', 'list_hap_imports', 'get_hap_import',
              'delete_hap_import', 'import_excel', 'list_excel_imports', 'get_excel_import', 'delete_excel_import',
              'push_revit_spaces', 'list_revit_spaces', 'list_revit_parameters', 'list_space_links', 'save_space_link')
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
    def __init__(self, url=PROJECT_URL, publishable_key=PUBLISHABLE_KEY, actor_name="", transport=None):
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
        if not isinstance(actor_name, STRING_TYPES) or len(actor_name.strip()) > 200:
            raise ValueError("Author names are limited to 200 characters.")
        self.url = url.rstrip("/")
        self.key = publishable_key
        self.actor_name = actor_name.strip()
        self.transport = transport or _request

    def rpc(self, name, args=None):
        import re
        if not re.match(r"^ffe_v3_[a-z_]+\Z", name):
            raise ValueError("Use a version 3 Loadwise RPC.")
        payload = json.dumps(args or {}, allow_nan=False)
        try:
            status, raw = self.transport(self.url + "/rest/v1/rpc/" + name, {"apikey": self.key}, payload)
        except RpcError:
            raise
        except Exception as error:
            # CLR WebException can occur before a response exists, including
            # while obtaining a request stream. Preserve a retryable result.
            raise RpcError("Could not confirm the Loadwise request: {0}".format(error), "NETWORK_ERROR")
        try:
            result = json.loads(raw)
        except (ValueError, TypeError):
            raise RpcError("Supabase returned an unreadable response.", "INVALID_RESPONSE", status)
        if not 200 <= status < 300:
            error = result if isinstance(result, dict) else {}
            raise RpcError(error.get("message", "The shared data request failed."), error.get("code"), status, error.get("details"))
        if not isinstance(result, dict) or result.get("ok") is not True:
            raise RpcError("Supabase returned an unexpected response.", "INVALID_RESPONSE", status)
        return result

    def call(self, action, data=None):
        if action not in OPERATIONS or (data is not None and not isinstance(data, dict)):
            raise ValueError("Invalid Loadwise request.")
        request = dict(data or {})
        request["schemaVersion"] = 3
        if self.actor_name:
            request["actor"] = {"name": self.actor_name}
        return self.rpc("ffe_v3_" + action, {"p_request": request})


def _operation(action):
    def invoke(self, request=None):
        return self.call(action, request)
    invoke.__name__ = action
    invoke.__doc__ = "Call ffe_v3_" + action + " with a version 3 request dictionary."
    return invoke


for _action in OPERATIONS:
    setattr(Client, _action, _operation(_action))
