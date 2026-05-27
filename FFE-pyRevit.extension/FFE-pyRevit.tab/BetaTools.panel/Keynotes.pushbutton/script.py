# -*- coding: utf-8 -*-
__title__ = "FFE-Keynotes"
__version__ = "Version = v0.5"
__persistentengine__ = True
__min_revit_ver__ = 2025
__doc__ = """Version = v0.5
Date    = 05.26.2026
__________________________________________________________________
Description:
Persistent WebView2 keynote manager for the active Revit document's
external keynote text file.

Key behaviors:
- Opens a modeless WebView2 window and keeps it alive with pyRevit's
  persistent engine.
- Reads and rewrites the keynote file assigned to the current Revit document.
- Edits structured tab-delimited keynote rows in a WebView.
- Saves with row-level merge checks, timestamped backups, sidecar file locks,
  Supabase mirror updates, and an immediate Revit keynote table reload.
- When a keynote key is renamed, writable placed/model keynote references using
  the old key are updated to the new key.
__________________________________________________________________
How-To:
- Click the button to open the keynote manager.
- Edit structured Key, Text, and Parent values.
- Click Save to merge edits into the assigned keynote file and reload Revit keynotes.
__________________________________________________________________
Last update:
- [05.19.2026] - v0.1 WebView2 keynote manager
- [05.20.2026] - v0.2 Refactor to support future features and simplify code maintenance.
- [05.20.2026] - v0.3 Made window stay on top of Revit and show in taskbar to prevent it from getting lost behind the main UI.
- [05.21.2026] - v0.4 Updated UI styles and layout, and added more robust file encoding detection and handling.
- [05.26.2026] - v0.5 Added in-place shared-file concurrent editing with Supabase/Postgres mirroring.

__________________________________________________________________
Author: Kyle Guggenheim"""


"""
Revit API notes:
- Targets Revit 2026.
- Modeless refresh/save requests are routed through ExternalEvent so
  document API calls run in a valid Revit API context.

Design decisions:
- V1 only manages the keynote file already assigned to the model. It
  does not repoint the project to a different keynote file.
- The shared text file is canonical. Supabase/Postgres is a coordination and
  realtime mirror layer, not the source of truth.
- Malformed source lines are shown and block save because the structured
  editor cannot safely preserve or repair arbitrary tab layouts.
"""


# ____________________________________________________________________ IMPORTS
import codecs
import hashlib
import json
import os
import shutil
import time
import traceback
import uuid

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit.DB import (
    BuiltInParameter,
    FilteredElementCollector,
    KeyBasedTreeEntriesLoadResults,
    KeynoteTable,
    Material,
    ModelPathUtils,
    Transaction,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from pyrevit import forms, revit, script


# ____________________________________________________________________ PYTHON COMPATIBILITY
try:
    unicode
except NameError:
    unicode = str


# ____________________________________________________________________ VARIABLES
revit_app = __revit__.Application
doc = revit.doc

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

APP_NAME = "FFE Keynote Manager"
APP_VERSION = "v0.5"
LOCAL_APP_NAME = "KeynoteManager"

LOGGER = script.get_logger()

try:
    WINDOW_REFS
except NameError:
    WINDOW_REFS = []


# ____________________________________________________________________ BASIC HELPERS
def safe_str(value):
    if value is None:
        return ""
    try:
        return str(value)
    except:
        try:
            return value.ToString()
        except:
            return ""


def safe_unicode(value):
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value, "utf-8")
    except:
        try:
            return unicode(value)
        except:
            return u""


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True)


def json_loads(value):
    return json.loads(value)


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path


def get_local_app_dir():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return ensure_dir(os.path.join(base_folder, "FFE-pyRevit", LOCAL_APP_NAME))


def get_webview_user_data_folder():
    return ensure_dir(os.path.join(get_local_app_dir(), "WebView2"))


def get_settings_path():
    return os.path.join(get_local_app_dir(), "window-state.json")


def get_supabase_settings_path():
    return os.path.join(get_local_app_dir(), "supabase-settings.json")


def get_generated_at():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def get_document_title(target_doc):
    try:
        return safe_str(target_doc.Title)
    except:
        return "Untitled Revit Document"


def get_client_name():
    try:
        username = safe_str(revit_app.Username).strip()
        if username:
            return username
    except:
        pass

    username = safe_str(os.environ.get("USERNAME")).strip()
    computer = safe_str(os.environ.get("COMPUTERNAME")).strip()
    if username and computer:
        return "{0}@{1}".format(username, computer)
    if username:
        return username
    return "Unknown User"


def normalize_path(path):
    value = safe_str(path)
    if not value:
        return ""
    try:
        return os.path.normcase(os.path.abspath(value))
    except:
        return os.path.normcase(value)


def read_json_file(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as file_obj:
            return json.loads(file_obj.read())
    except:
        return None


def write_json_file(path, payload):
    try:
        ensure_dir(os.path.dirname(path))
        with open(path, "w") as file_obj:
            file_obj.write(json.dumps(payload, ensure_ascii=True, indent=2))
    except:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass


def ask_for_supabase_value(prompt, default_value=""):
    try:
        value = forms.ask_for_string(
            prompt=prompt,
            default=default_value or "",
            title="Supabase Keynote Settings"
        )
        if value is None:
            return None
        return safe_str(value).strip()
    except:
        return None


def load_supabase_settings(prompt_if_missing=True, force_prompt=False):
    settings_path = get_supabase_settings_path()
    settings = read_json_file(settings_path) or {}

    url = safe_str(settings.get("url")).strip()
    anon_key = safe_str(settings.get("anonKey") or settings.get("publishableKey")).strip()
    client_id = safe_str(settings.get("clientId")).strip()

    if not client_id:
        client_id = safe_str(uuid.uuid4())

    if force_prompt or (prompt_if_missing and not url):
        next_url = ask_for_supabase_value(
            "Enter the Supabase project URL for the FFE Keynote Manager:",
            url
        )
        if next_url is not None:
            url = next_url

    if force_prompt or (prompt_if_missing and url and not anon_key):
        next_anon_key = ask_for_supabase_value(
            "Enter the Supabase publishable/anon key for the FFE Keynote Manager:",
            anon_key
        )
        if next_anon_key is not None:
            anon_key = next_anon_key

    payload = {
        "url": url,
        "anonKey": anon_key,
        "clientId": client_id,
        "clientName": get_client_name(),
        "settingsPath": settings_path,
        "configured": bool(url and anon_key),
    }

    if force_prompt or url or anon_key:
        settings.update({
            "url": url,
            "anonKey": anon_key,
            "clientId": client_id,
            "clientName": payload["clientName"],
        })
        write_json_file(settings_path, settings)

    return payload


# ____________________________________________________________________ WEBVIEW HELPERS
def get_revit_install_dir():
    try:
        app_path = revit_app.ApplicationPath
        if app_path and os.path.exists(app_path):
            return os.path.dirname(app_path)
    except:
        pass

    try:
        version = safe_str(revit_app.VersionNumber) or "2026"
    except:
        version = "2026"

    return os.path.join(
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        "Autodesk",
        "Revit {0}".format(version)
    )


def load_webview2_types():
    """Load Revit's bundled WebView2 assemblies and return the WPF control types."""
    revit_dir = get_revit_install_dir()
    dll_paths = [
        os.path.join(revit_dir, "Microsoft.Web.WebView2.Core.dll"),
        os.path.join(revit_dir, "Microsoft.Web.WebView2.Wpf.dll"),
    ]

    for dll_path in dll_paths:
        if not os.path.exists(dll_path):
            raise Exception("WebView2 assembly was not found: {0}".format(dll_path))
        try:
            clr.AddReferenceToFileAndPath(dll_path)
        except AttributeError:
            clr.AddReference(dll_path)

    path_parts = os.environ.get("PATH", "").split(os.pathsep)
    if revit_dir not in path_parts:
        os.environ["PATH"] = revit_dir + os.pathsep + os.environ.get("PATH", "")

    from Microsoft.Web.WebView2.Wpf import CoreWebView2CreationProperties, WebView2
    return WebView2, CoreWebView2CreationProperties


def make_file_uri(path):
    absolute_path = os.path.abspath(path).replace("\\", "/")
    return Uri("file:///" + absolute_path.replace(" ", "%20"))


# ____________________________________________________________________ KEYNOTE PATH HELPERS
def iterate_external_resource_references(refs):
    if refs is None:
        return []

    try:
        return list(dict(refs).items())
    except:
        pass

    try:
        return [(item.Key, item.Value) for item in refs]
    except:
        pass

    try:
        items = []
        enumerator = refs.GetEnumerator()
        while enumerator.MoveNext():
            current = enumerator.Current
            items.append((current.Key, current.Value))
        if items:
            return items
    except:
        pass

    try:
        return list(refs.Items)
    except:
        return []


def get_reference_path(resource_ref):
    if resource_ref is None:
        return ""

    try:
        path = safe_str(resource_ref.InSessionPath)
        if path:
            return path
    except:
        pass

    try:
        model_path = resource_ref.GetAbsolutePath()
        if model_path:
            path = safe_str(ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path))
            if path:
                return path
    except:
        pass

    try:
        if resource_ref.HasValidDisplayPath():
            path = safe_str(resource_ref.GetResourceShortDisplayName())
            if path:
                return path
    except:
        pass

    return ""


def get_keynote_reference(target_doc):
    """
    Get the keynote table and file reference for the given document, or raise an exception if it cannot be accessed.
    """
    if target_doc is None:
        raise Exception("No active Revit document is available.")

    keynote_table = KeynoteTable.GetKeynoteTable(target_doc)
    if keynote_table is None:
        raise Exception("This document does not expose a Revit keynote table.")

    refs = None
    try:
        refs = keynote_table.GetExternalResourceReferences()
    except Exception as exc:
        raise Exception("Could not read the keynote table external reference: {0}".format(exc))

    for ref_type, resource_ref in iterate_external_resource_references(refs):
        keynote_path = get_reference_path(resource_ref)
        if keynote_path:
            return keynote_table, resource_ref, keynote_path

    raise Exception("This document does not have a file-based keynote table reference.")


def looks_like_remote_resource(path):
    """
    Detect if the path looks like a remote resource that cannot be read as a local file.
    """
    value = safe_str(path).lower()
    return (
        value.startswith("http://") or
        value.startswith("https://") or
        value.startswith("rsn://") or
        value.startswith("bim 360://") or
        value.startswith("acc://")
    )


# ____________________________________________________________________ FILE ENCODING HELPERS
def read_binary_file(path):
    with open(path, "rb") as file_obj:
        return file_obj.read()


def write_binary_file(path, data):
    with open(path, "wb") as file_obj:
        file_obj.write(data)


def detect_line_ending(text):
    if "\r\n" in text:
        return "\r\n"
    if "\n" in text:
        return "\n"
    if "\r" in text:
        return "\r"
    return "\r\n"


def decode_keynote_bytes(raw_bytes):
    if raw_bytes.startswith(codecs.BOM_UTF8):
        return raw_bytes[len(codecs.BOM_UTF8):].decode("utf-8"), "utf-8-sig"

    if raw_bytes.startswith(codecs.BOM_UTF16_LE):
        return raw_bytes[len(codecs.BOM_UTF16_LE):].decode("utf-16-le"), "utf-16-le-bom"

    if raw_bytes.startswith(codecs.BOM_UTF16_BE):
        return raw_bytes[len(codecs.BOM_UTF16_BE):].decode("utf-16-be"), "utf-16-be-bom"

    if b"\x00" in raw_bytes[:200]:
        try:
            return raw_bytes.decode("utf-16"), "utf-16"
        except:
            pass

    try:
        return raw_bytes.decode("utf-8"), "utf-8"
    except:
        pass

    try:
        return raw_bytes.decode("mbcs"), "mbcs"
    except:
        return raw_bytes.decode("latin-1"), "latin-1"


def encode_keynote_text(text, encoding):
    value = safe_unicode(text)
    encoding = safe_str(encoding) or "utf-8"

    if encoding == "utf-8-sig":
        return codecs.BOM_UTF8 + value.encode("utf-8")

    if encoding == "utf-16-le-bom":
        return codecs.BOM_UTF16_LE + value.encode("utf-16-le")

    if encoding == "utf-16-be-bom":
        return codecs.BOM_UTF16_BE + value.encode("utf-16-be")

    return value.encode(encoding)


def get_file_state(path):
    raw_bytes = read_binary_file(path)
    return {
        "fileHash": hashlib.sha256(raw_bytes).hexdigest(),
        "lastWriteUtc": os.path.getmtime(path),
        "size": len(raw_bytes),
    }


def check_file_write_available(path):
    if not os.path.exists(path):
        return False, "The keynote file does not exist."

    if not os.path.isfile(path):
        return False, "The keynote path is not a file."

    try:
        if os.access(path, os.W_OK) is False:
            return False, "The keynote file is not writable for the current Windows user."
    except:
        pass

    try:
        file_obj = open(path, "r+b")
        file_obj.close()
        return True, "The keynote file is writable."
    except Exception as exc:
        return False, "The keynote file could not be opened for writing: {0}".format(exc)


# ____________________________________________________________________ KEYNOTE PARSING / VALIDATION
def make_issue(severity, message, key="", line_number=None, code=""):
    """
    Create a structured issue dictionary for reporting problems with keynote entries.
     - severity: "error" or "warning"
     - message: human-readable description of the issue
     - key: the keynote key associated with the issue, if applicable
     - line_number: the line number in the source text where the issue was found, if applicable
     - code: a short identifier for the type of issue, useful for programmatic handling   
    """
    issue = {
        "severity": severity,
        "message": message,
        "key": key or "",
        "code": code or "",
    }
    if line_number is not None:
        issue["lineNumber"] = line_number
    return issue


def normalize_keynote_lines(text):
    value = safe_unicode(text)
    return value.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def parse_keynote_text(text):
    entries = []
    issues = []

    for line_index, line in enumerate(normalize_keynote_lines(text)):
        line_number = line_index + 1
        if not line.strip():
            continue

        parts = line.split("\t")
        if len(parts) == 2:
            key = parts[0].strip()
            keynote_text = parts[1].strip()
            parent_key = ""
        elif len(parts) == 3:
            key = parts[0].strip()
            keynote_text = parts[1].strip()
            parent_key = parts[2].strip()
        else:
            issues.append(make_issue(
                "error",
                "Line {0} is not a valid keynote row. Expected 2 or 3 tab-delimited columns.".format(line_number),
                "",
                line_number,
                "malformedLine"
            ))
            continue

        entries.append({
            "id": "line-{0}".format(line_number),
            "key": key,
            "text": keynote_text,
            "parentKey": parent_key,
            "lineNumber": line_number,
        })

    return entries, issues


def validate_entries(entries, source_issues=None):
    issues = []
    entries = entries or []
    source_issues = source_issues or []

    for issue in source_issues:
        issues.append(issue)

    key_counts = {}
    key_to_entry = {}

    for entry in entries:
        key = safe_unicode(entry.get("key")).strip()
        text = safe_unicode(entry.get("text")).strip()
        parent_key = safe_unicode(entry.get("parentKey")).strip()
        line_number = entry.get("lineNumber")

        if not key:
            issues.append(make_issue("error", "Key is required.", key, line_number, "emptyKey"))
        if not text:
            issues.append(make_issue("error", "Text is required.", key, line_number, "emptyText"))

        for field_name, field_value in [("Key", key), ("Text", text), ("Parent", parent_key)]:
            if "\t" in field_value or "\r" in field_value or "\n" in field_value:
                issues.append(make_issue(
                    "error",
                    "{0} cannot contain tabs or line breaks.".format(field_name),
                    key,
                    line_number,
                    "invalidFieldCharacter"
                ))

        if key:
            key_counts[key] = key_counts.get(key, 0) + 1
            if key not in key_to_entry:
                key_to_entry[key] = entry

    for key, count in key_counts.items():
        if count > 1:
            issues.append(make_issue(
                "error",
                "Duplicate keynote key: {0}".format(key),
                key,
                None,
                "duplicateKey"
            ))

    for entry in entries:
        key = safe_unicode(entry.get("key")).strip()
        parent_key = safe_unicode(entry.get("parentKey")).strip()
        line_number = entry.get("lineNumber")

        if parent_key and parent_key not in key_to_entry:
            issues.append(make_issue(
                "error",
                "Parent key '{0}' was not found.".format(parent_key),
                key,
                line_number,
                "missingParent"
            ))

    parent_map = {}
    for entry in entries:
        key = safe_unicode(entry.get("key")).strip()
        parent_key = safe_unicode(entry.get("parentKey")).strip()
        if key:
            parent_map[key] = parent_key

    visited_cycles = set()
    for entry in entries:
        key = safe_unicode(entry.get("key")).strip()
        if not key or key in visited_cycles:
            continue

        chain = []
        seen = set()
        cursor = key
        while cursor:
            if cursor in seen:
                issues.append(make_issue(
                    "error",
                    "Parent cycle detected at key '{0}'.".format(cursor),
                    key,
                    entry.get("lineNumber"),
                    "parentCycle"
                ))
                for chain_key in chain:
                    visited_cycles.add(chain_key)
                break
            seen.add(cursor)
            chain.append(cursor)
            cursor = parent_map.get(cursor, "")

    if not entries:
        issues.append(make_issue(
            "warning",
            "The keynote file contains no readable keynote entries.",
            "",
            None,
            "emptyFile"
        ))

    return issues


def has_error_issues(issues):
    for issue in issues or []:
        if safe_str(issue.get("severity")).lower() == "error":
            return True
    return False


def canonicalize_entries(entries, line_ending):
    entries = entries or []
    line_ending = line_ending or "\r\n"

    by_key = {}
    children = {}
    ordered_entries = []

    for entry in entries:
        normalized = {
            "key": safe_unicode(entry.get("key")).strip(),
            "text": safe_unicode(entry.get("text")).strip(),
            "parentKey": safe_unicode(entry.get("parentKey")).strip(),
            "lineNumber": entry.get("lineNumber"),
        }
        ordered_entries.append(normalized)
        by_key[normalized["key"]] = normalized
        parent_key = normalized["parentKey"]
        if parent_key not in children:
            children[parent_key] = []
        children[parent_key].append(normalized)

    result = []
    visited = set()

    def append_branch(parent_key):
        for child in children.get(parent_key, []):
            child_key = child["key"]
            if child_key in visited:
                continue
            visited.add(child_key)
            result.append(child)
            append_branch(child_key)

    append_branch("")

    for entry in ordered_entries:
        key = entry["key"]
        if key not in visited:
            visited.add(key)
            result.append(entry)

    lines = []
    for entry in result:
        if entry["parentKey"]:
            lines.append(u"{0}\t{1}\t{2}".format(entry["key"], entry["text"], entry["parentKey"]))
        else:
            lines.append(u"{0}\t{1}".format(entry["key"], entry["text"]))

    if not lines:
        return u""

    return safe_unicode(line_ending).join(lines) + safe_unicode(line_ending)


# ____________________________________________________________________ PAYLOAD / SAVE HELPERS
def build_base_payload(target_doc, status, message):
    return {
        "name": APP_NAME,
        "version": APP_VERSION,
        "docTitle": get_document_title(target_doc),
        "keynotePath": "",
        "displayPath": "",
        "libraryKey": "",
        "encoding": "",
        "lineEnding": "\r\n",
        "lastWriteUtc": None,
        "fileHash": "",
        "writeAvailable": False,
        "writeMessage": "",
        "generatedAt": get_generated_at(),
        "supabase": load_supabase_settings(prompt_if_missing=True),
        "status": status,
        "message": message,
        "entries": [],
        "issues": [],
        "entryCount": 0,
    }


def build_keynote_payload(target_doc):
    payload = build_base_payload(target_doc, "error", "")

    try:
        keynote_table, resource_ref, keynote_path = get_keynote_reference(target_doc)
        payload["keynotePath"] = keynote_path
        payload["displayPath"] = keynote_path
        payload["libraryKey"] = normalize_path(keynote_path)

        if looks_like_remote_resource(keynote_path):
            payload["status"] = "unsupported"
            payload["message"] = "The current keynote table is a remote resource. This manager supports local and network file paths in v1."
            payload["issues"] = [make_issue("error", payload["message"], "", None, "unsupportedReference")]
            return payload

        if not os.path.exists(keynote_path):
            payload["status"] = "missingFile"
            payload["message"] = "The keynote file was not found: {0}".format(keynote_path)
            payload["issues"] = [make_issue("error", payload["message"], "", None, "missingFile")]
            return payload

        raw_bytes = read_binary_file(keynote_path)
        text, encoding = decode_keynote_bytes(raw_bytes)
        line_ending = detect_line_ending(text)
        entries, parse_issues = parse_keynote_text(text)
        issues = validate_entries(entries, parse_issues)
        file_state = get_file_state(keynote_path)
        write_ok, write_message = check_file_write_available(keynote_path)
        if not write_ok:
            issues.append(make_issue("warning", write_message, "", None, "writeUnavailable"))

        payload.update({
            "status": "invalidFormat" if has_error_issues(issues) else "ready",
            "message": "Loaded {0} keynote entries from the shared keynote file.".format(len(entries)),
            "encoding": encoding,
            "lineEnding": line_ending,
            "lastWriteUtc": file_state.get("lastWriteUtc"),
            "fileHash": file_state.get("fileHash"),
            "writeAvailable": write_ok,
            "writeMessage": write_message,
            "entries": entries,
            "issues": issues,
            "entryCount": len(entries),
        })

        if has_error_issues(issues):
            payload["message"] = "Loaded {0} entries, but validation errors must be fixed before saving.".format(len(entries))

    except Exception as exc:
        payload["status"] = "error"
        payload["message"] = safe_str(exc) or "Could not read the keynote file."
        payload["issues"] = [make_issue("error", payload["message"], "", None, "loadError")]
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass

    return payload


def create_backup_file(keynote_path):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    keynote_dir = os.path.dirname(keynote_path)
    base_name = os.path.basename(keynote_path)
    backup_name = "{0}.{1}.bak".format(base_name, timestamp)

    candidate_dirs = [
        os.path.join(keynote_dir, "_FFE_Keynote_Backups"),
        os.path.join(get_local_app_dir(), "Backups"),
    ]

    last_error = None
    for backup_dir in candidate_dirs:
        try:
            ensure_dir(backup_dir)
            backup_path = os.path.join(backup_dir, backup_name)
            shutil.copy2(keynote_path, backup_path)
            return backup_path
        except Exception as exc:
            last_error = exc

    raise Exception("Could not create a keynote backup: {0}".format(last_error))


def reload_revit_keynotes(target_doc):
    keynote_table, resource_ref, keynote_path = get_keynote_reference(target_doc)
    load_results = KeyBasedTreeEntriesLoadResults()
    transaction = Transaction(target_doc, "Reload Keynote Table")

    try:
        transaction.Start()
        load_result = keynote_table.LoadFrom(resource_ref, load_results)
        transaction.Commit()
    except:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    failure_messages = []
    try:
        for failure_message in load_results.GetFailureMessages():
            failure_messages.append(safe_str(failure_message.GetDescriptionText()))
    except:
        pass

    if failure_messages:
        raise Exception("Revit reported keynote reload failures: {0}".format(" | ".join(failure_messages)))

    if safe_str(load_result).lower() in ["false", "0"]:
        raise Exception("Revit did not reload the keynote table successfully.")

    return keynote_path


def get_element_id_key(element):
    try:
        return int(element.Id.IntegerValue)
    except:
        return safe_str(getattr(element, "Id", ""))


def get_parameter_text(parameter):
    if parameter is None:
        return ""

    try:
        value = parameter.AsString()
        if value:
            return safe_unicode(value).strip()
    except:
        pass

    try:
        value = parameter.AsValueString()
        if value:
            return safe_unicode(value).strip()
    except:
        pass

    return ""


def get_keynote_builtin_parameter_ids():
    parameter_ids = []
    for parameter_name in ["KEYNOTE_PARAM", "KEY_VALUE"]:
        try:
            parameter_id = getattr(BuiltInParameter, parameter_name)
            if parameter_id not in parameter_ids:
                parameter_ids.append(parameter_id)
        except:
            pass
    return parameter_ids


def get_keynote_parameters(element):
    parameters = []
    for parameter_id in get_keynote_builtin_parameter_ids():
        try:
            parameter = element.get_Parameter(parameter_id)
            if parameter is not None and parameter not in parameters:
                parameters.append(parameter)
        except:
            pass
    return parameters


def iter_keynote_reference_candidates(target_doc):
    seen = set()

    collectors = [
        FilteredElementCollector(target_doc).WhereElementIsNotElementType(),
        FilteredElementCollector(target_doc).WhereElementIsElementType(),
        FilteredElementCollector(target_doc).OfClass(Material),
    ]

    for collector in collectors:
        try:
            for element in collector:
                element_key = get_element_id_key(element)
                if not element_key or element_key in seen:
                    continue
                seen.add(element_key)
                yield element
        except:
            continue


def update_model_keynote_references(target_doc, key_renames):
    key_renames = key_renames or {}
    summary = {
        "renamedKeyCount": len(key_renames),
        "updatedCount": 0,
        "readOnlyCount": 0,
        "failedCount": 0,
        "unchangedCount": 0,
        "failures": [],
    }

    if not key_renames:
        return summary

    transaction = Transaction(target_doc, "Update Renamed Keynote References")

    try:
        transaction.Start()
        updated_element_ids = set()

        for element in iter_keynote_reference_candidates(target_doc):
            element_updated = False
            element_key = get_element_id_key(element)

            for parameter in get_keynote_parameters(element):
                current_key = get_parameter_text(parameter)
                next_key = key_renames.get(current_key)
                if not next_key:
                    continue

                try:
                    if parameter.IsReadOnly:
                        summary["readOnlyCount"] += 1
                        continue
                except:
                    pass

                try:
                    set_result = parameter.Set(next_key)
                    if safe_str(set_result).lower() in ["false", "0"]:
                        summary["failedCount"] += 1
                        if len(summary["failures"]) < 10:
                            summary["failures"].append("Element {0}: Revit rejected keynote value '{1}'.".format(element_key, next_key))
                        continue
                    element_updated = True
                except Exception as exc:
                    summary["failedCount"] += 1
                    if len(summary["failures"]) < 10:
                        summary["failures"].append("Element {0}: {1}".format(element_key, exc))

            if element_updated and element_key not in updated_element_ids:
                updated_element_ids.add(element_key)
                summary["updatedCount"] += 1

        transaction.Commit()
    except:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    return summary


def restore_backup_file(backup_path, keynote_path):
    if backup_path and os.path.exists(backup_path):
        shutil.copy2(backup_path, keynote_path)


def get_sync_sidecar_path(keynote_path):
    return "{0}.ffe-sync.json".format(keynote_path)


def get_sync_lock_path(keynote_path):
    return "{0}.ffe-sync.lock".format(keynote_path)


def write_sync_sidecar(keynote_path, payload):
    write_json_file(get_sync_sidecar_path(keynote_path), payload)


def acquire_sync_lock(lock_path, timeout_seconds=10, stale_seconds=120):
    start_time = time.time()

    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                lock_payload = json_dumps({
                    "clientName": get_client_name(),
                    "createdAt": get_generated_at(),
                })
                try:
                    os.write(fd, lock_payload)
                except TypeError:
                    os.write(fd, lock_payload.encode("utf-8"))
            finally:
                os.close(fd)
            return lock_path
        except OSError:
            try:
                if os.path.exists(lock_path):
                    lock_age = time.time() - os.path.getmtime(lock_path)
                    if lock_age > stale_seconds:
                        os.remove(lock_path)
                        continue
            except:
                pass

            if time.time() - start_time >= timeout_seconds:
                raise Exception("Timed out waiting for the keynote sync lock: {0}".format(lock_path))

            time.sleep(0.25)


def release_sync_lock(lock_path):
    try:
        if lock_path and os.path.exists(lock_path):
            os.remove(lock_path)
    except:
        pass


def normalize_merge_entry(entry):
    entry = entry or {}
    return {
        "id": safe_str(entry.get("id")),
        "key": safe_unicode(entry.get("key")).strip(),
        "text": safe_unicode(entry.get("text")).strip(),
        "parentKey": safe_unicode(entry.get("parentKey")).strip(),
        "lineNumber": entry.get("lineNumber"),
    }


def merge_entry_fields_equal(first_entry, second_entry):
    first_entry = normalize_merge_entry(first_entry)
    second_entry = normalize_merge_entry(second_entry)
    return (
        first_entry["key"] == second_entry["key"] and
        first_entry["text"] == second_entry["text"] and
        first_entry["parentKey"] == second_entry["parentKey"]
    )


def keyed_entries(entries):
    result = {}
    for entry in entries or []:
        normalized = normalize_merge_entry(entry)
        if normalized["key"]:
            result[normalized["key"]] = normalized
    return result


def indexed_entries(entries):
    result = {}
    for entry in entries or []:
        normalized = normalize_merge_entry(entry)
        if normalized["id"]:
            result[normalized["id"]] = normalized
    return result


def make_save_changes(baseline_entries, desired_entries):
    baseline_by_id = indexed_entries(baseline_entries)
    desired_by_id = indexed_entries(desired_entries)
    changes = []
    seen_ids = set()

    for baseline in baseline_entries or []:
        base_entry = normalize_merge_entry(baseline)
        if not base_entry["id"]:
            continue
        desired_entry = desired_by_id.get(base_entry["id"])
        seen_ids.add(base_entry["id"])

        if desired_entry is None:
            changes.append({
                "type": "delete",
                "base": base_entry,
                "next": None,
            })
        elif not merge_entry_fields_equal(base_entry, desired_entry):
            changes.append({
                "type": "update",
                "base": base_entry,
                "next": normalize_merge_entry(desired_entry),
            })

    for desired in desired_entries or []:
        desired_entry = normalize_merge_entry(desired)
        if not desired_entry["id"] or desired_entry["id"] in seen_ids:
            continue
        changes.append({
            "type": "insert",
            "base": None,
            "next": desired_entry,
        })

    return changes


def make_key_rename_map(baseline_entries, desired_entries):
    desired_by_id = indexed_entries(desired_entries)
    key_renames = {}

    for baseline in baseline_entries or []:
        base_entry = normalize_merge_entry(baseline)
        desired_entry = desired_by_id.get(base_entry["id"])
        if not desired_entry:
            continue

        old_key = base_entry["key"]
        new_key = normalize_merge_entry(desired_entry)["key"]
        if old_key and new_key and old_key != new_key:
            key_renames[old_key] = new_key

    return key_renames


def current_entry_changed(current_entry, baseline_entry):
    if current_entry is None:
        return True
    return not merge_entry_fields_equal(current_entry, baseline_entry)


def find_entry_index(entries, key):
    for index, entry in enumerate(entries or []):
        if normalize_merge_entry(entry)["key"] == key:
            return index
    return -1


def make_conflict_issue(message, key):
    return make_issue("error", message, key or "", None, "rowConflict")


def merge_keynote_entries(current_entries, baseline_entries, desired_entries):
    """
    Merge the desired keynote entries with the current entries, using the baseline entries to detect conflicts.
        - current_entries: the list of entries currently in the shared keynote file
    """
    current_list = [normalize_merge_entry(entry) for entry in (current_entries or [])]
    current_by_key = keyed_entries(current_list)
    conflicts = []
    changes = make_save_changes(baseline_entries, desired_entries)

    for change in changes:
        change_type = change.get("type")
        base_entry = change.get("base")
        next_entry = change.get("next")

        if change_type == "insert":
            if current_by_key.get(next_entry["key"]):
                conflicts.append(make_conflict_issue(
                    "A keynote with key '{0}' already exists in the shared file.".format(next_entry["key"]),
                    next_entry["key"]
                ))
            continue

        current_entry = current_by_key.get(base_entry["key"])
        if current_entry_changed(current_entry, base_entry):
            conflicts.append(make_conflict_issue(
                "Key '{0}' changed in the shared file before your save.".format(base_entry["key"]),
                base_entry["key"]
            ))
            continue

        if change_type == "update" and next_entry["key"] != base_entry["key"]:
            existing_target = current_by_key.get(next_entry["key"])
            if existing_target and existing_target["key"] != base_entry["key"]:
                conflicts.append(make_conflict_issue(
                    "A keynote with key '{0}' already exists in the shared file.".format(next_entry["key"]),
                    next_entry["key"]
                ))

    if conflicts:
        return None, conflicts

    for change in changes:
        change_type = change.get("type")
        base_entry = change.get("base")
        next_entry = change.get("next")

        if change_type == "insert":
            current_list.append(next_entry)
            current_by_key[next_entry["key"]] = next_entry
            continue

        entry_index = find_entry_index(current_list, base_entry["key"])
        if entry_index < 0:
            continue

        if change_type == "delete":
            removed = current_list.pop(entry_index)
            current_by_key.pop(removed["key"], None)
            continue

        if change_type == "update":
            old_key = base_entry["key"]
            new_key = next_entry["key"]
            current_list[entry_index] = next_entry
            current_by_key.pop(old_key, None)
            current_by_key[new_key] = next_entry
            if old_key != new_key:
                for entry in current_list:
                    if entry["parentKey"] == old_key:
                        entry["parentKey"] = new_key

    validation_issues = validate_entries(current_list, [])
    if has_error_issues(validation_issues):
        return None, validation_issues

    return current_list, []


def save_keynote_payload(target_doc, save_payload):
    save_payload = save_payload or {}

    try:
        source_has_malformed = bool(save_payload.get("sourceHasMalformed"))
        if source_has_malformed:
            return {
                "status": "error",
                "message": "The source file contains malformed rows. Refresh after fixing those lines outside the manager before saving.",
                "issues": [make_issue("error", "Malformed source rows block structured save.", "", None, "malformedLine")],
            }

        keynote_table, resource_ref, current_path = get_keynote_reference(target_doc)
        requested_path = safe_str(save_payload.get("keynotePath"))
        if normalize_path(current_path) != normalize_path(requested_path):
            return {
                "status": "error",
                "message": "The document keynote path changed. Refresh before saving.",
                "issues": [make_issue("error", "The document keynote path changed.", "", None, "pathChanged")],
            }

        if not os.path.exists(current_path):
            return {
                "status": "error",
                "message": "The keynote file no longer exists: {0}".format(current_path),
                "issues": [make_issue("error", "The keynote file no longer exists.", "", None, "missingFile")],
            }

        entries = save_payload.get("entries") or []
        baseline_entries = save_payload.get("baselineEntries") or []
        issues = validate_entries(entries, [])
        if has_error_issues(issues):
            return {
                "status": "error",
                "message": "Validation errors must be fixed before saving.",
                "issues": issues,
            }

        encoding = safe_str(save_payload.get("encoding")) or "utf-8"
        line_ending = save_payload.get("lineEnding") or "\r\n"
        key_renames = make_key_rename_map(baseline_entries, entries)
        reference_update = {}
        lock_path = get_sync_lock_path(current_path)
        backup_path = ""

        acquire_sync_lock(lock_path)

        try:
            write_ok, write_message = check_file_write_available(current_path)
            if not write_ok:
                return {
                    "status": "error",
                    "message": write_message,
                    "issues": [make_issue("error", write_message, "", None, "writeUnavailable")],
                    "payload": build_keynote_payload(target_doc),
                }

            raw_bytes = read_binary_file(current_path)
            current_text, current_encoding = decode_keynote_bytes(raw_bytes)
            current_line_ending = detect_line_ending(current_text)
            current_entries, current_parse_issues = parse_keynote_text(current_text)
            current_issues = validate_entries(current_entries, current_parse_issues)
            if has_error_issues(current_issues):
                return {
                    "status": "error",
                    "message": "The shared keynote file changed into an invalid format. Refresh after fixing the text file.",
                    "issues": current_issues,
                    "payload": build_keynote_payload(target_doc),
                }

            merged_entries, merge_issues = merge_keynote_entries(current_entries, baseline_entries, entries)
            if merge_issues:
                return {
                    "status": "conflict",
                    "message": "Some keynote rows changed in the shared file before your save. Refresh to review the latest file.",
                    "issues": merge_issues,
                    "payload": build_keynote_payload(target_doc),
                }

            if current_encoding:
                encoding = current_encoding
            if current_line_ending:
                line_ending = current_line_ending

            keynote_text = canonicalize_entries(merged_entries, line_ending)
            encoded = encode_keynote_text(keynote_text, encoding)
            backup_path = create_backup_file(current_path)
            write_binary_file(current_path, encoded)
            reload_revit_keynotes(target_doc)
            reference_update = update_model_keynote_references(target_doc, key_renames)
        except Exception as exc:
            restore_backup_file(backup_path, current_path)
            try:
                reload_revit_keynotes(target_doc)
            except:
                pass
            return {
                "status": "error",
                "message": "The keynote save failed. The previous file was restored when a backup was available: {0}".format(exc),
                "backupPath": backup_path,
                "issues": [make_issue("error", safe_str(exc), "", None, "reloadFailed")],
                "payload": build_keynote_payload(target_doc),
            }
        finally:
            release_sync_lock(lock_path)

        payload = build_keynote_payload(target_doc)
        result_issues = payload.get("issues") or []
        if reference_update.get("readOnlyCount"):
            result_issues.append(make_issue(
                "warning",
                "{0} keynote reference parameter(s) matched renamed keys but were read-only.".format(reference_update.get("readOnlyCount")),
                "",
                None,
                "keynoteReferenceReadOnly"
            ))
        if reference_update.get("failedCount"):
            result_issues.append(make_issue(
                "warning",
                "{0} keynote reference parameter(s) could not be updated after key renaming.".format(reference_update.get("failedCount")),
                "",
                None,
                "keynoteReferenceUpdateFailed"
            ))
        try:
            file_state = get_file_state(current_path)
            write_sync_sidecar(current_path, {
                "libraryKey": normalize_path(current_path),
                "encoding": payload.get("encoding") or encoding,
                "lineEnding": payload.get("lineEnding") or line_ending,
                "fileHash": file_state.get("fileHash"),
                "lastWriteUtc": file_state.get("lastWriteUtc"),
                "size": file_state.get("size"),
                "savedAt": get_generated_at(),
                "clientName": get_client_name(),
                "keyRenames": key_renames,
                "keynoteReferenceUpdate": reference_update,
                "source": "sharedFileCanonical",
            })
        except Exception as exc:
            result_issues.append(make_issue(
                "warning",
                "Saved the shared keynote file, but could not write sync metadata: {0}".format(exc),
                "",
                None,
                "sidecarWriteFailed"
            ))
        message = "Merged keynote edits into the shared file and reloaded Revit keynotes."
        if key_renames:
            message += " Updated {0} model keynote reference(s) for renamed keys.".format(
                reference_update.get("updatedCount", 0)
            )

        return {
            "status": "ready",
            "message": message,
            "backupPath": backup_path,
            "keyRenames": key_renames,
            "keynoteReferenceUpdate": reference_update,
            "issues": result_issues,
            "payload": payload,
        }

    except Exception as exc:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass
        return {
            "status": "error",
            "message": safe_str(exc) or "Could not save the keynote file.",
            "issues": [make_issue("error", safe_str(exc), "", None, "saveError")],
        }


# ____________________________________________________________________ EXTERNAL EVENT
class KeynoteManagerEventHandler(IExternalEventHandler):
    def __init__(self):
        self.window = None
        self.pending_action = None
        self.pending_payload = None

    def GetName(self):
        return "FFE Keynote Manager Bridge"

    def queue_refresh(self):
        self.pending_action = "refresh"
        self.pending_payload = None

    def queue_save(self, payload):
        self.pending_action = "save"
        self.pending_payload = payload

    def clear_pending(self):
        self.pending_action = None
        self.pending_payload = None

    def Execute(self, uiapp):
        window = self.window
        action = self.pending_action
        payload = self.pending_payload
        self.clear_pending()

        if window is None:
            return

        if action == "refresh":
            keynote_payload = build_keynote_payload(window.document)
            window.set_payload(keynote_payload)
            window.send_keynote_payload(force=True)
            window.send_status(keynote_payload.get("status"), keynote_payload.get("message"))
            return

        if action == "save":
            result = save_keynote_payload(window.document, payload)
            if result.get("payload"):
                window.set_payload(result.get("payload"))
            window.send_save_result(result)
            return

        window.send_status("warning", "No keynote manager action was queued.")


# ____________________________________________________________________ WEBVIEW WINDOW
class KeynoteManagerWindow(Window):
    """
    Modeless WPF window hosting a WebView2 control for managing Revit keynote files, with communication routed through an ExternalEvent handler.
        - Initializes WebView2 and loads the local index.html file as the UI.
    """
    def __init__(self, webview_type, creation_properties_type, keynote_payload, event_handler, external_event):
        """
        Initialize the keynote manager window with the given WebView2 types, initial payload, and ExternalEvent handler.
         - webview_type: the WPF WebView2 control type
         - creation_properties_type: the type for configuring WebView2 creation properties
         - keynote_payload: the initial data payload to send to the web app after loading
         - event_handler: the instance of KeynoteManagerEventHandler that will handle refresh/save requests from the web app
         - external_event: the ExternalEvent instance associated with the event handler, used for invoking actions in Revit's API context
        """
        Window.__init__(self)

        try:
            WindowInteropHelper(self).Owner = __revit__.MainWindowHandle
            self.ShowInTaskbar = True
        except:
            pass

        self.document = doc
        self.keynote_payload = keynote_payload
        self.event_handler = event_handler
        self.external_event = external_event
        self.has_sent_payload = False
        self.has_dirty_edits = False
        self.close_discard_confirmed = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = self.get_window_title()
        self.Width = 1320
        self.Height = 860
        self.MinWidth = 420
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize
        self.apply_saved_window_state()

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing FFE Keynote Manager WebView2..."

        try:
            creation_properties = creation_properties_type()
            creation_properties.UserDataFolder = get_webview_user_data_folder()
            self.browser.CreationProperties = creation_properties
        except Exception as exc:
            self.status_text.Text = "Could not configure WebView2 user data folder:\n{0}".format(exc)

        content_grid = Grid()
        content_grid.Children.Add(self.browser)
        content_grid.Children.Add(self.status_text)
        self.Content = content_grid

        self.Loaded += self.on_loaded
        self.Closing += self.on_closing
        self.Closed += self.on_closed
        self.browser.CoreWebView2InitializationCompleted += self.on_core_webview2_initialized
        self.browser.NavigationCompleted += self.on_navigation_completed

    def get_window_title(self):
        doc_title = self.keynote_payload.get("docTitle") or get_document_title(self.document)
        return "{0} - {1}".format(APP_NAME, doc_title)

    def apply_saved_window_state(self):
        state = read_json_file(get_settings_path()) or {}
        try:
            width = float(state.get("width") or 0)
            height = float(state.get("height") or 0)
            left = state.get("left")
            top = state.get("top")

            if width >= self.MinWidth:
                self.Width = width
            if height >= self.MinHeight:
                self.Height = height
            if left is not None and top is not None:
                self.Left = float(left)
                self.Top = float(top)
                self.WindowStartupLocation = WindowStartupLocation.Manual
        except:
            pass

    def save_window_state(self):
        try:
            write_json_file(get_settings_path(), {
                "width": float(self.Width),
                "height": float(self.Height),
                "left": float(self.Left),
                "top": float(self.Top),
            })
        except:
            pass

    def set_payload(self, keynote_payload):
        self.keynote_payload = keynote_payload or build_base_payload(self.document, "error", "No keynote payload.")
        self.Title = self.get_window_title()

    def on_loaded(self, sender, args):
        self.status_text.Visibility = Visibility.Visible
        self.status_text.Text = "Loading FFE Keynote Manager from:\n{0}".format(self.index_uri.AbsoluteUri)
        try:
            self.browser.EnsureCoreWebView2Async()
        except Exception as exc:
            self.status_text.Text = "Could not initialize WebView2:\n{0}".format(exc)

    def on_core_webview2_initialized(self, sender, args):
        try:
            if args.IsSuccess:
                self.browser.CoreWebView2.WebMessageReceived += self.on_web_message_received
                self.browser.CoreWebView2.Navigate(self.index_uri.AbsoluteUri)
            else:
                message = "WebView2 initialization failed."
                try:
                    message = "{0}\n{1}".format(message, args.InitializationException.Message)
                except:
                    pass
                self.status_text.Text = message
        except:
            self.status_text.Text = "WebView2 initialized but navigation failed:\n{0}".format(traceback.format_exc())

    def on_navigation_completed(self, sender, args):
        try:
            if not args.IsSuccess:
                self.status_text.Visibility = Visibility.Visible
                self.status_text.Text = "FFE Keynote Manager navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus)
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_keynote_payload()

    def confirm_discard_changes(self):
        if not self.has_dirty_edits:
            return True

        return bool(forms.alert(
            "Close the manager and discard unsaved keynote edits?",
            title="Discard Unsaved Changes",
            ok=False,
            yes=True,
            no=True,
            warn_icon=True
        ))

    def on_closing(self, sender, args):
        if self.close_discard_confirmed:
            return

        if not self.confirm_discard_changes():
            try:
                args.Cancel = True
            except:
                pass
            return

        self.close_discard_confirmed = True

    def on_closed(self, sender, args):
        self.save_window_state()
        try:
            self.external_event.Dispose()
        except:
            pass
        try:
            self.browser.Dispose()
        except:
            pass
        try:
            WINDOW_REFS.remove(self)
        except:
            pass

    def execute_script(self, script_text):
        try:
            self.browser.ExecuteScriptAsync(script_text)
        except:
            self.status_text.Visibility = Visibility.Visible
            self.status_text.Text = "Could not send data to the Keynote Manager web app:\n{0}".format(
                traceback.format_exc()
            )

    def call_keynote_app(self, method_name, payload):
        self.execute_script(
            "window.ffeKeynotes && window.ffeKeynotes.{0}({1});".format(
                method_name,
                json_dumps(payload)
            )
        )

    def send_keynote_payload(self, force=False):
        if self.has_sent_payload and not force:
            return
        self.has_sent_payload = True
        self.call_keynote_app("loadData", self.keynote_payload)

    def send_status(self, status, message):
        self.call_keynote_app("setStatus", {
            "status": status or "idle",
            "message": message or "",
        })

    def send_save_result(self, result):
        self.call_keynote_app("handleSaveResult", result or {})

    def request_refresh_from_app(self):
        self.call_keynote_app("requestRefresh", {})

    def raise_external_event(self, action_name):
        try:
            self.external_event.Raise()
        except Exception as exc:
            self.event_handler.clear_pending()
            result = {
                "status": "error",
                "message": "Could not raise the Revit {0} event: {1}".format(action_name, exc),
            }
            self.send_save_result(result)

    def on_web_message_received(self, sender, args):
        raw_message = ""
        try:
            raw_message = args.TryGetWebMessageAsString()
        except:
            try:
                raw_message = args.WebMessageAsJson
            except:
                raw_message = ""

        if not raw_message:
            return

        try:
            message = json_loads(raw_message)
        except:
            return

        message_type = message.get("type")

        if message_type == "appReady":
            self.send_keynote_payload()
            return

        if message_type == "dirtyStateChanged":
            self.has_dirty_edits = bool(message.get("dirty"))
            return

        if message_type == "refreshData":
            self.event_handler.queue_refresh()
            self.send_status("warning", "Refreshing keynote file...")
            self.raise_external_event("refresh")
            return

        if message_type == "configureSupabase":
            settings = load_supabase_settings(prompt_if_missing=True, force_prompt=True)
            keynote_payload = build_keynote_payload(self.document)
            keynote_payload["supabase"] = settings
            self.set_payload(keynote_payload)
            self.send_keynote_payload(force=True)
            self.send_status(keynote_payload.get("status"), keynote_payload.get("message"))
            return

        if message_type == "saveKeynotes":
            self.event_handler.queue_save(message.get("payload") or {})
            self.send_status("warning", "Saving keynote file and reloading Revit...")
            self.raise_external_event("save")
            return

        if message_type == "closeWindow":
            self.close_discard_confirmed = True
            if message.get("discardConfirmed"):
                self.has_dirty_edits = False
            self.Close()


# ____________________________________________________________________ WINDOW LIFETIME
def focus_existing_window():
    for window in list(WINDOW_REFS):
        try:
            if window.IsVisible:
                window.Activate()
                window.request_refresh_from_app()
                return True
        except:
            try:
                WINDOW_REFS.remove(window)
            except:
                pass
    return False


# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The FFE Keynote Manager web app was not found:\n{0}".format(PATH_INDEX),
        title=APP_NAME,
        exitscript=True
    )

if not focus_existing_window():
    try:
        keynote_payload = build_keynote_payload(doc)
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as startup_error:
        forms.alert(
            "Could not start the FFE Keynote Manager.\n\n{0}".format(startup_error),
            title=APP_NAME,
            exitscript=True
        )

    handler = KeynoteManagerEventHandler()
    external_event = ExternalEvent.Create(handler)
    window = KeynoteManagerWindow(
        WebView2,
        CoreWebView2CreationProperties,
        keynote_payload,
        handler,
        external_event
    )
    handler.window = window
    WINDOW_REFS.append(window)
    window.Show()
