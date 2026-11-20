# -*- coding: utf-8 -*-
"""Purpose: call FFE Mechanical's shared RPCs from IronPython 2.7 or CPython.
Behaviors: bounded HTTPS requests, structured errors, revision-checked publishing.
Revit API: 2025/2026 compatible; this transport never accesses Revit objects.
Design: packaged from FFE-Mechanical/integrations/pyrevit/ffe_supabase.py so the
command can be installed independently. Keep RPC signatures aligned with that client.

SQL is owned by FFE-Mechanical/supabase/migrations. Uses the shared publishable
key without individual sign-in. modified_by is optional, client-reported attribution.
Perform network requests away from the Revit UI/ExternalEvent thread.
"""
import json
import sys

PROJECT_URL = "https://ncxylphudovrmvdgqsol.supabase.co"
PUBLISHABLE_KEY = "sb_publishable_VFbJW7ZQQh0Ro254lfXXvw_xesXNxTK"


class RpcError(Exception):
    def __init__(self, message, code=None, status=None):
        Exception.__init__(self, message)
        self.code = code
        self.status = status


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
    def __init__(self, url=PROJECT_URL, publishable_key=PUBLISHABLE_KEY, modified_by="", transport=None):
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
        try:
            string_types = (basestring,)
        except NameError:
            string_types = (str,)
        if not isinstance(modified_by, string_types) or len(modified_by) > 200:
            raise ValueError("Uploader identity must be at most 200 characters.")
        self.url = url.rstrip("/")
        self.key = publishable_key
        self.modified_by = modified_by
        self.transport = transport or _request

    def rpc(self, name, args=None):
        import re
        if not re.match(r"^ffe_[a-z_]+\Z", name):
            raise ValueError("Invalid FFE RPC name.")
        status, raw = self.transport(self.url + "/rest/v1/rpc/" + name,
                                     {"apikey": self.key},
                                     json.dumps(args or {}, allow_nan=False))
        try:
            result = json.loads(raw)
        except (ValueError, TypeError):
            raise RpcError("Supabase returned an unreadable response.", "INVALID_RESPONSE", status)
        if not 200 <= status < 300:
            error = result if isinstance(result, dict) else {}
            raise RpcError(error.get("message", "The shared data request failed."), error.get("code"), status)
        return result

    def list_projects(self):
        return self.rpc("ffe_list_projects")

    def create_project(self, project_id, name):
        return self.rpc("ffe_create_project", {"p_project_id": project_id, "p_name": name})

    def list_models(self, project_id):
        return self.rpc("ffe_list_models", {"p_project_id": project_id})

    def create_model(self, project_id, model_id, name):
        return self.rpc("ffe_create_model", {"p_project_id": project_id, "p_model_id": model_id, "p_name": name, "p_modified_by": self.modified_by})

    def get_snapshot(self, model_id):
        return self.rpc("ffe_get_snapshot", {"p_model_id": model_id})

    def publish_snapshot(self, model_id, expected_revision, snapshot):
        return self.rpc("ffe_publish_snapshot", {"p_model_id": model_id, "p_expected_revision": expected_revision, "p_snapshot": snapshot, "p_modified_by": self.modified_by})

    def list_layouts(self, model_id):
        return self.rpc("ffe_list_layouts", {"p_model_id": model_id})

    def get_layout(self, model_id, layout_id):
        return self.rpc("ffe_get_layout", {"p_model_id": model_id, "p_layout_id": layout_id})

    def save_layout(self, model_id, layout_id, name, layout, expected_revision=None, version=4):
        return self.rpc("ffe_save_layout", {"p_model_id": model_id, "p_layout_id": layout_id,
                        "p_expected_revision": expected_revision, "p_name": name, "p_layout": layout, "p_version": version, "p_modified_by": self.modified_by})

    def delete_layout(self, model_id, layout_id, expected_revision):
        return self.rpc("ffe_delete_layout", {"p_model_id": model_id, "p_layout_id": layout_id, "p_expected_revision": expected_revision})
