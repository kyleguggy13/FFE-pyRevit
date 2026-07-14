# -*- coding: utf-8 -*-
__title__ = "FFE-Keynotes"
__version__ = "v0.16"
__persistentengine__ = True
__min_revit_ver__ = 2026
__doc__ = """Version = v0.16
Date    = 07.14.2026
__________________________________________________________________
Description:
Persistent WebView2 keynote manager for the active Revit document's
external keynote text file.
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
- [05.27.2026] - v0.6 Added row-level locking with Supabase to prevent concurrent edit conflicts.
- [05.29.2026] - v0.7 Improved error handling and improved UI
- [06.02.2026] - v0.8 Added Generic Annotation keynote placement and type synchronization.
- [06.02.2026] - v0.9 Applied the standard leader arrowhead to Generic Annotation keynote types.
- [06.03.2026] - v0.10 Added Analytics tracking for keynote manager usage and errors.
- [06.09.2026] - v0.11 Made Place As persist across sessions.
- [06.10.2026] - v0.12 Added placement filter and collapsible division panel.
- [06.12.2026] - v0.13 Added automatic model health scan and Safe Mode.
- [07.10.2026] - v0.14 Added per-keynote family type/text file conflict resolution in Safe Mode.
- [07.14.2026] - v0.15 Added per-row ellipsis actions for copying, deleting, moving, sequencing, and uppercasing notes.
- [07.14.2026] - v0.16 Added note promotion and parent menus with safe demotion and subnote-aware deletion.
__________________________________________________________________
Author: Kyle Guggenheim"""


"""
TODO:
- Select placed keynote icon and display list of sheets/views where it is placed.
- Add realtime updates when a keynote is added or removed.
- Figure out how to work around Keynote's workset.


Key behaviors:
- Opens a modeless WebView2 window and keeps it alive with pyRevit's
  persistent engine.
- Reads and rewrites the keynote file assigned to the current Revit document.
- Edits structured tab-delimited keynote rows in a WebView.
- Saves with row-level merge checks, timestamped backups, sidecar file locks,
  Supabase mirror updates, and an immediate Revit keynote table reload.
- When a keynote key is renamed, writable placed/model keynote references using
  the old key are updated to the new key.
- Places saved rows as either Revit User Keynotes or FFE_Symbol_Keynote Generic
  Annotations.
- Provides row-specific utility actions without changing the selected keynote first.
- Supports promoting notes to parents and demoting parents while preserving their descendant hierarchy.

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
from System.Windows import Clipboard, ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Interop import WindowInteropHelper

try:
    from pyrevit.runtime.types import DocumentEventUtils
except:
    DocumentEventUtils = None

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementType,
    Family,
    FilteredElementCollector,
    KeyBasedTreeEntriesLoadResults,
    KeynoteTable,
    Material,
    ModelPathUtils,
    Transaction,
    Viewport,
    ViewSheet,
    WorksetKind,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, PostableCommand, RevitCommandId

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
APP_VERSION = "v0.16"
LOCAL_APP_NAME = "KeynoteManager"
GENERIC_KEYNOTE_FAMILY_NAME = "FFE_Symbol_Keynote (Type)"
GENERIC_KEYNOTE_NUMBER_PARAMETER = "Number"
GENERIC_KEYNOTE_TEXT_PARAMETER = "Text"
GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME = "Arrow Filled 20 Degree"
MODEL_HEALTH_SAFE_MODE_MISSING_KEY_COUNT = 5
MODEL_HEALTH_SAFE_MODE_RATIO = 0.20
MODEL_HEALTH_SAFE_MODE_RATIO_MIN_MISSING_KEYS = 2
MODEL_HEALTH_SAFE_MODE_RATIO_MIN_PLACED_KEYS = 5

SHARED_SUPABASE_SETTINGS_PARTS = (
    "FFE Inc",
    "FFE Revit Users - Documents",
    "00-General",
    "Revit_Add-Ins",
    "FFE-pyRevit",
    "Documentation",
    "supabase.json",
)
PROVIDED_SHARED_SUPABASE_SETTINGS_PATH = "C:\\Users\\kyleg\\FFE Inc\\FFE Revit Users - Documents\\00-General\\Revit_Add-Ins\\FFE-pyRevit\\Documentation\\supabase.json"
SUPABASE_URL_KEYS = ("url", "projectUrl", "projectURL", "supabaseUrl", "Supabase URL", "Project URL")
SUPABASE_ANON_KEY_KEYS = ("anonKey", "anon_key", "publishableKey", "publishable_key", "Publishable key", "Anon key", "Public anon key")
SUPABASE_CLIENT_ID_KEYS = ("clientId", "client_id", "Client ID")

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


def looks_like_uri_path(path):
    value = safe_str(path).strip()
    marker_index = value.find("://")
    if marker_index <= 0:
        return False

    prefix = value[:marker_index]
    return "/" not in prefix and "\\" not in prefix


def normalize_path(path):
    value = safe_str(path).strip()
    if not value:
        return ""

    # Revit returns cloud resources as URI-like display paths such as
    # "Autodesk Docs://Project/Model.rvt". os.path.abspath treats those as
    # relative filesystem paths and incorrectly prefixes the process cwd.
    if looks_like_uri_path(value):
        return value

    try:
        return os.path.normcase(os.path.abspath(value))
    except:
        return os.path.normcase(value)


def get_shared_supabase_settings_paths():
    paths = []
    user_folder = os.path.expanduser("~")

    if user_folder and user_folder != "~":
        paths.append(os.path.join(user_folder, *SHARED_SUPABASE_SETTINGS_PARTS))

    paths.append(PROVIDED_SHARED_SUPABASE_SETTINGS_PATH)

    unique_paths = []
    seen_paths = {}
    for path in paths:
        normalized_path = normalize_path(path)
        if not normalized_path or normalized_path in seen_paths:
            continue
        seen_paths[normalized_path] = True
        unique_paths.append(path)

    return unique_paths


def read_json_file(path):
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as file_obj:
            raw_value = file_obj.read()

        if raw_value.startswith(codecs.BOM_UTF8):
            raw_value = raw_value[len(codecs.BOM_UTF8):]

        try:
            return json.loads(raw_value.decode("utf-8"))
        except:
            return json.loads(raw_value)
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


def read_user_settings():
    settings = read_json_file(get_settings_path()) or {}
    if isinstance(settings, dict):
        return settings
    return {}


def normalize_placement_mode(value):
    value = safe_str(value).strip()
    if value == "genericAnnotation":
        return "genericAnnotation"
    return "userKeynote"


def get_saved_placement_mode():
    return normalize_placement_mode(read_user_settings().get("placementMode"))


def save_placement_mode_setting(value):
    placement_mode = normalize_placement_mode(value)
    settings = read_user_settings()
    settings["placementMode"] = placement_mode
    write_json_file(get_settings_path(), settings)
    return placement_mode


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


def normalize_supabase_config_key(key):
    value = safe_str(key).strip().lower()
    value = value.replace(" ", "")
    value = value.replace("_", "")
    value = value.replace("-", "")
    return value


def get_supabase_config_value(settings, allowed_keys):
    if not isinstance(settings, dict):
        return ""

    allowed_lookup = {}
    for key in allowed_keys:
        allowed_lookup[normalize_supabase_config_key(key)] = True

    for key, value in settings.items():
        if normalize_supabase_config_key(key) in allowed_lookup:
            result = safe_str(value).strip()
            if result:
                return result

    return ""


def iter_supabase_settings_candidates(value):
    if isinstance(value, dict):
        yield value
        for nested_value in value.values():
            if isinstance(nested_value, (dict, list, tuple)):
                for candidate in iter_supabase_settings_candidates(nested_value):
                    yield candidate
    elif isinstance(value, (list, tuple)):
        for item in value:
            for candidate in iter_supabase_settings_candidates(item):
                yield candidate


def extract_supabase_settings(value):
    partial_settings = {}

    for candidate in iter_supabase_settings_candidates(value):
        url = get_supabase_config_value(candidate, SUPABASE_URL_KEYS)
        anon_key = get_supabase_config_value(candidate, SUPABASE_ANON_KEY_KEYS)
        client_id = get_supabase_config_value(candidate, SUPABASE_CLIENT_ID_KEYS)

        if not url and not anon_key:
            continue

        settings = {
            "url": url,
            "anonKey": anon_key,
            "clientId": client_id,
        }

        if url and anon_key:
            return settings

        if not partial_settings:
            partial_settings = settings

    return partial_settings


def load_shared_supabase_settings():
    for path in get_shared_supabase_settings_paths():
        settings = extract_supabase_settings(read_json_file(path))
        if settings.get("url") or settings.get("anonKey"):
            settings["settingsSourcePath"] = path
            return settings

    return {}


def load_supabase_settings(prompt_if_missing=True, force_prompt=False):
    settings_path = get_supabase_settings_path()
    settings = read_json_file(settings_path) or {}

    url = safe_str(settings.get("url")).strip()
    anon_key = safe_str(settings.get("anonKey") or settings.get("publishableKey")).strip()
    client_id = safe_str(settings.get("clientId")).strip()
    settings_source_path = settings_path

    if not client_id:
        client_id = safe_str(uuid.uuid4())

    if not (url and anon_key):
        shared_settings = load_shared_supabase_settings()
        if shared_settings:
            if not url:
                url = safe_str(shared_settings.get("url")).strip()
            if not anon_key:
                anon_key = safe_str(shared_settings.get("anonKey")).strip()
            settings_source_path = safe_str(shared_settings.get("settingsSourcePath")).strip() or settings_source_path

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
        "settingsSourcePath": settings_source_path,
        "configured": bool(url and anon_key),
    }

    if force_prompt or url or anon_key:
        settings.update({
            "url": url,
            "anonKey": anon_key,
            "clientId": client_id,
            "clientName": payload["clientName"],
            "settingsSourcePath": settings_source_path,
        })
        write_json_file(settings_path, settings)

    return payload


# ____________________________________________________________________ WEBVIEW HELPERS
def get_revit_install_dir():
    """
    Get the Revit installation directory, which is needed to load the bundled WebView2 assemblies. 
    This function tries multiple methods to find the path in case of different Revit versions or configurations.
    """
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


def is_keynote_comment_line(line):
    return safe_unicode(line).lstrip().startswith(u"#")


def natural_sort_parts(value):
    value = safe_unicode(value).strip().lower()
    parts = []
    token = u""
    token_is_digit = None

    def append_token():
        if not token:
            return
        if token_is_digit:
            parts.append((0, int(token), len(token)))
        else:
            parts.append((1, token))

    for character in value:
        character_is_digit = character.isdigit()
        if token and character_is_digit != token_is_digit:
            append_token()
            token = character
        else:
            token += character
        token_is_digit = character_is_digit

    append_token()
    return tuple(parts)


def natural_entry_sort_key(entry):
    key = safe_unicode((entry or {}).get("key")).strip()
    text = safe_unicode((entry or {}).get("text")).strip()
    try:
        line_number = int((entry or {}).get("lineNumber") or 0)
    except:
        line_number = 0
    return (natural_sort_parts(key), key.lower(), text.lower(), line_number)


def escape_keynote_metadata_value(value):
    return safe_unicode(value).replace(u"\"", u"\\\"")


def format_keynote_encoding_label(encoding):
    value = safe_str(encoding).strip().lower()
    if value.startswith("utf-16"):
        return "UTF-16"
    if value == "utf-8-sig":
        return "UTF-8 BOM"
    if value == "utf-8":
        return "UTF-8"
    if value == "mbcs":
        return "Windows ANSI"
    if value == "latin-1":
        return "Latin-1"
    return safe_str(encoding).strip() or "UTF-8"


def make_keynote_metadata_header(source_path, encoding):
    return u'# @datastore("txt") @source("{0}") @encoding("{1}")'.format(
        escape_keynote_metadata_value(source_path),
        escape_keynote_metadata_value(format_keynote_encoding_label(encoding))
    )


def parse_keynote_text(text):
    entries = []
    issues = []

    for line_index, line in enumerate(normalize_keynote_lines(text)):
        line_number = line_index + 1
        if not line.strip():
            continue
        if is_keynote_comment_line(line):
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


def canonicalize_entries(entries, line_ending, source_path="", encoding=""):
    entries = entries or []
    line_ending = line_ending or "\r\n"

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
        parent_key = normalized["parentKey"]
        if parent_key not in children:
            children[parent_key] = []
        children[parent_key].append(normalized)

    for parent_key in children:
        children[parent_key].sort(key=natural_entry_sort_key)

    lines = [
        make_keynote_metadata_header(source_path, encoding),
        u'# --------------------- @table(categories:"Root Keynotes Table")',
    ]
    visited = set()

    def append_category_row(entry):
        lines.append(u"{0}\t{1}".format(entry["key"], entry["text"]))

    def append_keynote_row(entry):
        lines.append(u"{0}\t{1}\t{2}".format(entry["key"], entry["text"], entry["parentKey"]))

    def append_keynote_branch(entry):
        entry_key = entry["key"]
        if entry_key in visited:
            return
        visited.add(entry_key)
        append_keynote_row(entry)
        for child in children.get(entry_key, []):
            append_keynote_branch(child)

    for root_entry in children.get("", []):
        append_category_row(root_entry)

    lines.append(u'# --------------------- @table(keynotes:"Keynotes Table")')

    for root_entry in children.get("", []):
        for child_entry in children.get(root_entry["key"], []):
            append_keynote_branch(child_entry)

    for entry in sorted(ordered_entries, key=natural_entry_sort_key):
        if entry["parentKey"] and entry["key"] not in visited:
            append_keynote_branch(entry)

    if len(lines) <= 3:
        return u""

    return safe_unicode(line_ending).join(lines) + safe_unicode(line_ending)


# ____________________________________________________________________ PAYLOAD / SAVE HELPERS
def make_empty_model_health(status="notScanned", message="Model health has not been scanned."):
    return {
        "status": status,
        "message": message,
        "scannedAt": None,
        "safeModeRecommended": False,
        "signature": "",
        "placedKeyCount": 0,
        "placedCount": 0,
        "missingKeyCount": 0,
        "missingPlacedCount": 0,
        "missingRatio": 0,
        "userKeynoteCount": 0,
        "genericAnnotationCount": 0,
        "sheetCount": 0,
        "unsheetedCount": 0,
        "skippedCount": 0,
        "placedKeyMap": {},
        "issues": [],
    }


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
        "preferences": {
            "placementMode": get_saved_placement_mode(),
        },
        "status": status,
        "message": message,
        "entries": [],
        "issues": [],
        "entryCount": 0,
        "sheetVisibleKeynotes": {},
        "modelHealth": make_empty_model_health(),
    }


def build_keynote_payload(target_doc, include_model_health=True):
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
        model_health = make_empty_model_health()
        if include_model_health:
            model_health = build_model_health(target_doc, {
                "libraryKey": normalize_path(keynote_path),
                "displayPath": keynote_path,
                "keynotePath": keynote_path,
                "encoding": encoding,
                "lineEnding": line_ending,
                "fileHash": file_state.get("fileHash"),
                "lastWriteUtc": file_state.get("lastWriteUtc"),
                "entries": entries,
                "entryCount": len(entries),
            })

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
            "sheetVisibleKeynotes": model_health.get("placedKeyMap") or {},
            "modelHealth": model_health,
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
        return int(element.Id.Value)
    except:
        pass

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


def get_lookup_parameter_text(element, parameter_names):
    for parameter_name in parameter_names or []:
        try:
            value = get_parameter_text(element.LookupParameter(parameter_name))
            if value:
                return value
        except:
            pass
    return ""


def get_keynote_tag_key(tag):
    key = get_lookup_parameter_text(tag, ["Key Value", "Keynote Key"])
    if key:
        return key

    for parameter in get_keynote_parameters(tag):
        key = get_parameter_text(parameter)
        if key:
            return key
    return ""


def get_element_name(element):
    if element is None:
        return ""

    try:
        parameter = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        value = get_parameter_text(parameter)
        if value:
            return value
    except:
        pass

    try:
        return safe_unicode(element.Name).strip()
    except:
        return ""


def get_generic_annotation_keynote_family(target_doc):
    if target_doc is None:
        return None

    try:
        category = target_doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericAnnotation)
        category_id = category.Id
    except:
        return None

    try:
        families = FilteredElementCollector(target_doc).OfClass(Family)
    except:
        return None
    for family in families:
        try:
            if family.Name == GENERIC_KEYNOTE_FAMILY_NAME and family.FamilyCategory.Id == category_id:
                return family
        except:
            continue

    return None


def get_family_symbols(target_doc, family):
    symbols = []
    if target_doc is None or family is None:
        return symbols

    try:
        symbol_ids = family.GetFamilySymbolIds()
    except:
        return symbols

    for symbol_id in symbol_ids:
        try:
            symbol = target_doc.GetElement(symbol_id)
            if symbol is not None:
                symbols.append(symbol)
        except:
            continue

    return symbols


def get_generic_annotation_symbol_key(symbol):
    key = get_lookup_parameter_text(symbol, [GENERIC_KEYNOTE_NUMBER_PARAMETER])
    if key:
        return key
    return get_element_name(symbol)


def generic_annotation_symbol_matches_key(symbol, key):
    key = safe_unicode(key).strip()
    if not key:
        return False
    return (
        get_lookup_parameter_text(symbol, [GENERIC_KEYNOTE_NUMBER_PARAMETER]) == key or
        get_element_name(symbol) == key
    )


def get_generic_annotation_instance_symbol(target_doc, instance):
    try:
        symbol = instance.Symbol
        if symbol is not None:
            return symbol
    except:
        pass

    try:
        return target_doc.GetElement(instance.GetTypeId())
    except:
        return None


def iter_generic_annotation_keynote_instances(target_doc):
    if target_doc is None:
        return

    instances = (
        FilteredElementCollector(target_doc)
        .OfCategory(BuiltInCategory.OST_GenericAnnotation)
        .WhereElementIsNotElementType()
    )

    for instance in instances:
        symbol = get_generic_annotation_instance_symbol(target_doc, instance)
        if symbol is None:
            continue
        try:
            family = symbol.Family
            if family is not None and family.Name == GENERIC_KEYNOTE_FAMILY_NAME:
                yield instance, symbol
        except:
            continue


def collect_generic_annotation_instances_by_symbol(target_doc):
    result = {}
    for instance, symbol in iter_generic_annotation_keynote_instances(target_doc):
        symbol_key = get_element_id_key(symbol)
        if symbol_key not in result:
            result[symbol_key] = []
        result[symbol_key].append(instance)
    return result


def make_generic_annotation_sync_summary():
    return {
        "familyFound": False,
        "createdCount": 0,
        "updatedCount": 0,
        "migratedCount": 0,
        "retypedCount": 0,
        "deletedCount": 0,
        "preservedCount": 0,
        "failedCount": 0,
        "failures": [],
    }


def record_generic_annotation_sync_failure(summary, message):
    summary["failedCount"] += 1
    if len(summary["failures"]) < 10:
        summary["failures"].append(safe_str(message))


def get_writable_symbol_parameter(symbol, parameter_name):
    try:
        parameter = symbol.LookupParameter(parameter_name)
    except:
        parameter = None

    if parameter is None:
        raise Exception("Type '{0}' is missing the '{1}' parameter.".format(get_element_name(symbol), parameter_name))

    try:
        if parameter.IsReadOnly:
            raise Exception("Type '{0}' has a read-only '{1}' parameter.".format(get_element_name(symbol), parameter_name))
    except AttributeError:
        pass

    return parameter


def get_writable_builtin_symbol_parameter(symbol, parameter_id, parameter_label):
    try:
        parameter = symbol.get_Parameter(parameter_id)
    except:
        parameter = None

    if parameter is None:
        raise Exception("Type '{0}' is missing the '{1}' parameter.".format(get_element_name(symbol), parameter_label))

    try:
        if parameter.IsReadOnly:
            raise Exception("Type '{0}' has a read-only '{1}' parameter.".format(get_element_name(symbol), parameter_label))
    except AttributeError:
        pass

    return parameter


def get_element_id_value(element_id):
    try:
        return int(element_id.Value)
    except:
        pass

    try:
        return int(element_id.IntegerValue)
    except:
        return safe_str(element_id)


def get_leader_arrowhead_type(target_doc):
    if target_doc is None:
        return None

    try:
        element_types = (
            FilteredElementCollector(target_doc)
            .OfClass(ElementType)
            .WhereElementIsElementType()
        )
    except:
        return None

    for element_type in element_types:
        try:
            if (
                safe_str(element_type.FamilyName).strip() == "Arrowhead" and
                get_element_name(element_type) == GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME
            ):
                return element_type
        except:
            continue

    return None


def set_parameter_text(parameter, value):
    value = safe_unicode(value)
    if get_parameter_text(parameter) == value:
        return False

    set_result = parameter.Set(value)
    if safe_str(set_result).lower() in ["false", "0"]:
        raise Exception("Revit rejected parameter value '{0}'.".format(value))
    return True


def set_symbol_type_name(symbol, value):
    value = safe_unicode(value).strip()
    if get_element_name(symbol) == value:
        return False

    direct_name_error = ""
    try:
        symbol.Name = value
        return True
    except Exception as exc:
        direct_name_error = safe_str(exc)

    try:
        parameter = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    except:
        parameter = None

    if parameter is not None:
        try:
            if parameter.IsReadOnly:
                raise Exception("Type name is read-only.")
        except AttributeError:
            pass
        set_result = parameter.Set(value)
        if safe_str(set_result).lower() in ["false", "0"]:
            raise Exception("Revit rejected type name '{0}'.".format(value))
        return True

    raise Exception("Could not rename Generic Annotation type to '{0}': {1}".format(value, direct_name_error))


def set_generic_annotation_symbol_leader_arrowhead(symbol, leader_arrowhead_type):
    if leader_arrowhead_type is None:
        raise Exception("Arrowhead type '{0}' was not found.".format(GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME))

    parameter = get_writable_builtin_symbol_parameter(
        symbol,
        BuiltInParameter.LEADER_ARROWHEAD,
        "Leader Arrowhead"
    )

    try:
        if get_element_id_value(parameter.AsElementId()) == get_element_id_value(leader_arrowhead_type.Id):
            return False
    except:
        pass

    set_result = parameter.Set(leader_arrowhead_type.Id)
    if safe_str(set_result).lower() in ["false", "0"]:
        raise Exception("Revit rejected Leader Arrowhead '{0}'.".format(GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME))
    return True


def set_generic_annotation_symbol_values(symbol, key, text, leader_arrowhead_type):
    old_name = get_element_name(symbol)
    changed = False
    changed = set_symbol_type_name(symbol, key) or changed
    changed = set_parameter_text(
        get_writable_symbol_parameter(symbol, GENERIC_KEYNOTE_NUMBER_PARAMETER),
        key
    ) or changed
    changed = set_parameter_text(
        get_writable_symbol_parameter(symbol, GENERIC_KEYNOTE_TEXT_PARAMETER),
        text
    ) or changed
    changed = set_generic_annotation_symbol_leader_arrowhead(symbol, leader_arrowhead_type) or changed
    return changed, bool(old_name and old_name != safe_unicode(key).strip())


def choose_generic_annotation_symbol(symbols, key):
    key = safe_unicode(key).strip()
    matching = [symbol for symbol in symbols if generic_annotation_symbol_matches_key(symbol, key)]
    for symbol in matching:
        if get_element_name(symbol) == key:
            return symbol
    return matching and matching[0] or None


def duplicate_family_symbol(target_doc, base_symbol, new_type_name):
    duplicate_result = base_symbol.Duplicate(safe_unicode(new_type_name).strip())
    try:
        symbol = target_doc.GetElement(duplicate_result)
        if symbol is not None:
            return symbol
    except:
        pass
    return duplicate_result


def change_generic_annotation_instances_type(instances_by_symbol, old_symbol, new_symbol, summary):
    old_symbol_key = get_element_id_key(old_symbol)
    new_symbol_key = get_element_id_key(new_symbol)
    remaining_instances = []

    for instance in list(instances_by_symbol.get(old_symbol_key) or []):
        try:
            instance.ChangeTypeId(new_symbol.Id)
            summary["retypedCount"] += 1
            if new_symbol_key not in instances_by_symbol:
                instances_by_symbol[new_symbol_key] = []
            instances_by_symbol[new_symbol_key].append(instance)
        except Exception as exc:
            remaining_instances.append(instance)
            record_generic_annotation_sync_failure(
                summary,
                "Could not retype Generic Annotation instance {0}: {1}".format(
                    get_element_id_key(instance),
                    exc
                )
            )

    instances_by_symbol[old_symbol_key] = remaining_instances
    return not remaining_instances


def delete_generic_annotation_symbol_if_unused(target_doc, symbol, instances_by_symbol, summary, reason):
    symbol_key = get_element_id_key(symbol)
    if instances_by_symbol.get(symbol_key):
        summary["preservedCount"] += 1
        return False

    try:
        target_doc.Delete(symbol.Id)
        summary["deletedCount"] += 1
        return True
    except Exception as exc:
        summary["preservedCount"] += 1
        record_generic_annotation_sync_failure(
            summary,
            "Could not delete {0} Generic Annotation type '{1}': {2}".format(
                reason,
                get_element_name(symbol),
                exc
            )
        )
        return False


def sync_generic_annotation_types(target_doc, entries, key_renames, deleted_keys):
    summary = make_generic_annotation_sync_summary()
    family = get_generic_annotation_keynote_family(target_doc)
    if family is None:
        return summary

    summary["familyFound"] = True
    symbols = get_family_symbols(target_doc, family)
    if not symbols:
        record_generic_annotation_sync_failure(
            summary,
            "Family '{0}' does not contain a type to synchronize.".format(GENERIC_KEYNOTE_FAMILY_NAME)
        )
        return summary

    leader_arrowhead_type = get_leader_arrowhead_type(target_doc)
    if leader_arrowhead_type is None:
        record_generic_annotation_sync_failure(
            summary,
            "Arrowhead type '{0}' was not found. Load or create it before synchronizing Generic Annotation keynotes.".format(
                GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME
            )
        )
        return summary

    instances_by_symbol = collect_generic_annotation_instances_by_symbol(target_doc)
    deleted_symbol_keys = set()
    active_keys = set([
        safe_unicode(entry.get("key")).strip()
        for entry in entries or []
        if safe_unicode(entry.get("key")).strip()
    ])
    transaction = Transaction(target_doc, "Sync Generic Annotation Keynotes")

    try:
        transaction.Start()

        for entry in entries or []:
            key = safe_unicode(entry.get("key")).strip()
            text = safe_unicode(entry.get("text")).strip()
            if not key:
                continue

            target_symbol = choose_generic_annotation_symbol(symbols, key)
            previous_keys = [
                old_key for old_key, new_key in (key_renames or {}).items()
                if safe_unicode(new_key).strip() == key
            ]
            previous_symbols = []
            for previous_key in previous_keys:
                previous_symbols.extend([
                    symbol for symbol in symbols
                    if get_element_id_key(symbol) not in deleted_symbol_keys and
                    generic_annotation_symbol_matches_key(symbol, previous_key)
                ])

            if target_symbol is None and previous_symbols:
                target_symbol = previous_symbols[0]
            if target_symbol is None:
                continue

            try:
                changed, migrated = set_generic_annotation_symbol_values(
                    target_symbol,
                    key,
                    text,
                    leader_arrowhead_type
                )
                if changed:
                    summary["updatedCount"] += 1
                if migrated:
                    summary["migratedCount"] += 1
            except Exception as exc:
                record_generic_annotation_sync_failure(
                    summary,
                    "Could not synchronize Generic Annotation type '{0}' for keynote '{1}': {2}".format(
                        get_element_name(target_symbol),
                        key,
                        exc
                    )
                )
                continue

            redundant_symbols = []
            for symbol in symbols:
                symbol_key = get_element_id_key(symbol)
                if symbol_key == get_element_id_key(target_symbol) or symbol_key in deleted_symbol_keys:
                    continue
                if generic_annotation_symbol_matches_key(symbol, key) or symbol in previous_symbols:
                    redundant_symbols.append(symbol)

            for redundant_symbol in redundant_symbols:
                if not change_generic_annotation_instances_type(
                    instances_by_symbol,
                    redundant_symbol,
                    target_symbol,
                    summary
                ):
                    summary["preservedCount"] += 1
                    continue
                if delete_generic_annotation_symbol_if_unused(
                    target_doc,
                    redundant_symbol,
                    instances_by_symbol,
                    summary,
                    "redundant"
                ):
                    deleted_symbol_keys.add(get_element_id_key(redundant_symbol))

        for deleted_key in deleted_keys or []:
            if deleted_key in active_keys:
                continue
            for symbol in symbols:
                symbol_key = get_element_id_key(symbol)
                if symbol_key in deleted_symbol_keys:
                    continue
                if not generic_annotation_symbol_matches_key(symbol, deleted_key):
                    continue
                if delete_generic_annotation_symbol_if_unused(
                    target_doc,
                    symbol,
                    instances_by_symbol,
                    summary,
                    "deleted-keynote"
                ):
                    deleted_symbol_keys.add(symbol_key)

        transaction.Commit()
    except Exception as exc:
        try:
            transaction.RollBack()
        except:
            pass
        summary["updatedCount"] = 0
        summary["migratedCount"] = 0
        summary["retypedCount"] = 0
        summary["deletedCount"] = 0
        record_generic_annotation_sync_failure(
            summary,
            "Generic Annotation synchronization was rolled back: {0}".format(exc)
        )

    return summary


def element_is_on_non_user_workset(target_doc, element):
    try:
        workset = target_doc.GetWorksetTable().GetWorkset(element.WorksetId)
    except:
        return False

    try:
        return workset.Kind != WorksetKind.UserWorkset
    except:
        return safe_str(getattr(workset, "Kind", "")) != "UserWorkset"


def collect_sheet_visible_keynotes(target_doc):
    result = {}
    if target_doc is None:
        return result

    tags = (
        FilteredElementCollector(target_doc)
        .OfCategory(BuiltInCategory.OST_KeynoteTags)
        .WhereElementIsNotElementType()
    )

    for tag in tags:
        if not element_is_on_non_user_workset(target_doc, tag):
            continue

        key = safe_unicode(get_keynote_tag_key(tag)).strip()
        if key:
            result[key] = True

    for instance, symbol in iter_generic_annotation_keynote_instances(target_doc):
        key = safe_unicode(get_generic_annotation_symbol_key(symbol)).strip()
        if key:
            result[key] = True

    return result


def is_valid_element_id_value(value):
    try:
        return int(value) > 0
    except:
        return bool(value and safe_str(value) not in ["", "-1"])


def get_document_central_path(target_doc):
    if target_doc is None:
        return ""

    try:
        model_path = target_doc.GetWorksharingCentralModelPath()
        if model_path:
            path = safe_str(ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)).strip()
            if path:
                return path
    except:
        pass

    return ""


def get_document_path(target_doc):
    if target_doc is None:
        return ""

    try:
        return safe_str(target_doc.PathName).strip()
    except:
        return ""


def get_document_analytics_identity(target_doc):
    title = get_document_title(target_doc)
    central_path = get_document_central_path(target_doc)
    document_path = get_document_path(target_doc)
    key_source = "title"
    key_value = title

    if central_path:
        key_source = "centralPath"
        key_value = central_path
    elif document_path:
        key_source = "path"
        key_value = document_path

    if key_source == "title":
        document_key = "title:{0}".format(title)
    else:
        document_key = normalize_path(key_value)

    return {
        "documentKey": document_key,
        "documentKeySource": key_source,
        "documentTitle": title,
        "documentPath": document_path,
        "centralPath": central_path,
    }


def element_is_sheet(element):
    if element is None:
        return False

    try:
        if isinstance(element, ViewSheet):
            return True
    except:
        pass

    try:
        safe_unicode(element.SheetNumber)
        return True
    except:
        return False


def make_sheet_analytics_info(sheet):
    if sheet is None:
        return None

    try:
        sheet_id = safe_str(get_element_id_key(sheet))
    except:
        sheet_id = ""

    try:
        sheet_number = safe_unicode(sheet.SheetNumber).strip()
    except:
        sheet_number = ""

    sheet_name = get_element_name(sheet)

    if not (sheet_id or sheet_number or sheet_name):
        return None

    return {
        "id": sheet_id,
        "number": sheet_number,
        "name": sheet_name,
    }


def make_view_analytics_info(view):
    if view is None:
        return None

    try:
        view_id = safe_str(get_element_id_key(view))
    except:
        view_id = ""

    view_name = get_element_name(view)
    if not (view_id or view_name):
        return None

    return {
        "id": view_id,
        "name": view_name,
    }


def sheet_analytics_key(sheet_info):
    if not sheet_info:
        return ""
    return safe_str(sheet_info.get("id")).strip() or safe_unicode(sheet_info.get("number")).strip()


def add_sheet_to_view_lookup(result, view_id_value, sheet_info):
    if not is_valid_element_id_value(view_id_value):
        return

    sheet_key = sheet_analytics_key(sheet_info)
    if not sheet_key:
        return

    if view_id_value not in result:
        result[view_id_value] = []

    for existing in result[view_id_value]:
        if sheet_analytics_key(existing) == sheet_key:
            return

    result[view_id_value].append(sheet_info)


def build_view_sheet_lookup(target_doc):
    result = {}
    if target_doc is None:
        return result

    try:
        viewports = (
            FilteredElementCollector(target_doc)
            .OfClass(Viewport)
            .WhereElementIsNotElementType()
        )
    except:
        return result

    for viewport in viewports:
        try:
            view_id = viewport.ViewId
            view_id_value = get_element_id_value(view_id)
            sheet = target_doc.GetElement(viewport.SheetId)
            sheet_info = make_sheet_analytics_info(sheet)
            add_sheet_to_view_lookup(result, view_id_value, sheet_info)
        except:
            continue

    return result


def get_element_owner_view(target_doc, element):
    try:
        owner_view_id = element.OwnerViewId
    except:
        return None, None

    owner_view_id_value = get_element_id_value(owner_view_id)
    if not is_valid_element_id_value(owner_view_id_value):
        return None, None

    try:
        return target_doc.GetElement(owner_view_id), owner_view_id_value
    except:
        return None, owner_view_id_value


def get_element_sheet_infos(target_doc, element, view_sheet_lookup):
    owner_view, owner_view_id_value = get_element_owner_view(target_doc, element)
    if owner_view is None and not is_valid_element_id_value(owner_view_id_value):
        return [], None

    view_info = make_view_analytics_info(owner_view)
    if element_is_sheet(owner_view):
        sheet_info = make_sheet_analytics_info(owner_view)
        return sheet_info and [sheet_info] or [], view_info

    return list(view_sheet_lookup.get(owner_view_id_value) or []), view_info


def make_keynote_analytics_row(key, entry):
    return {
        "keynoteKey": safe_unicode(key).strip(),
        "keynoteText": safe_unicode((entry or {}).get("text")).strip(),
        "parentKey": safe_unicode((entry or {}).get("parentKey")).strip(),
        "inLibrary": bool(entry),
        "placed": False,
        "placedCount": 0,
        "userKeynoteCount": 0,
        "genericAnnotationCount": 0,
        "sheetCount": 0,
        "unsheetedCount": 0,
        "sheets": [],
        "_sheetMap": {},
    }


def analytics_sheet_sort_key(sheet_info):
    if not sheet_info:
        return ""
    return "{0}|{1}|{2}".format(
        safe_unicode(sheet_info.get("number")).strip(),
        safe_unicode(sheet_info.get("name")).strip(),
        safe_unicode(sheet_info.get("id")).strip()
    ).lower()


def sort_analytics_rows(rows):
    return sorted(rows, key=lambda row: safe_unicode(row.get("keynoteKey")).lower())


def record_sheet_analytics(row, sheet_info, source_type, view_info):
    sheet_key = sheet_analytics_key(sheet_info)
    if not sheet_key:
        row["unsheetedCount"] += 1
        return

    sheet_map = row["_sheetMap"]
    if sheet_key not in sheet_map:
        sheet_map[sheet_key] = {
            "id": safe_str(sheet_info.get("id")).strip(),
            "number": safe_unicode(sheet_info.get("number")).strip(),
            "name": safe_unicode(sheet_info.get("name")).strip(),
            "count": 0,
            "userKeynoteCount": 0,
            "genericAnnotationCount": 0,
            "viewIds": [],
            "viewNames": [],
            "_viewIdMap": {},
            "_viewNameMap": {},
        }

    sheet_row = sheet_map[sheet_key]
    sheet_row["count"] += 1
    if source_type == "userKeynote":
        sheet_row["userKeynoteCount"] += 1
    elif source_type == "genericAnnotation":
        sheet_row["genericAnnotationCount"] += 1

    if view_info:
        view_id = safe_str(view_info.get("id")).strip()
        view_name = safe_unicode(view_info.get("name")).strip()
        if view_id and view_id not in sheet_row["_viewIdMap"]:
            sheet_row["_viewIdMap"][view_id] = True
            sheet_row["viewIds"].append(view_id)
        if view_name and view_name not in sheet_row["_viewNameMap"]:
            sheet_row["_viewNameMap"][view_name] = True
            sheet_row["viewNames"].append(view_name)


def record_keynote_analytics_placement(rows_by_key, entry_by_key, key, source_type, element, target_doc, view_sheet_lookup):
    key = safe_unicode(key).strip()
    if not key:
        return False

    if key not in rows_by_key:
        rows_by_key[key] = make_keynote_analytics_row(key, entry_by_key.get(key))

    row = rows_by_key[key]
    row["placed"] = True
    row["placedCount"] += 1
    if source_type == "userKeynote":
        row["userKeynoteCount"] += 1
    elif source_type == "genericAnnotation":
        row["genericAnnotationCount"] += 1

    sheet_infos, view_info = get_element_sheet_infos(target_doc, element, view_sheet_lookup)
    if sheet_infos:
        for sheet_info in sheet_infos:
            record_sheet_analytics(row, sheet_info, source_type, view_info)
    else:
        row["unsheetedCount"] += 1

    return True


def finalize_keynote_analytics_rows(rows_by_key):
    rows = []
    for row in rows_by_key.values():
        sheet_values = []
        for sheet_row in row.get("_sheetMap", {}).values():
            sheet_row.pop("_viewIdMap", None)
            sheet_row.pop("_viewNameMap", None)
            sheet_row["viewIds"] = sorted(sheet_row.get("viewIds") or [])
            sheet_row["viewNames"] = sorted(sheet_row.get("viewNames") or [])
            sheet_values.append(sheet_row)
        sheet_values = sorted(sheet_values, key=analytics_sheet_sort_key)
        row["sheetCount"] = len(sheet_values)
        row["sheets"] = sheet_values
        row["placed"] = bool(row.get("placedCount"))
        row.pop("_sheetMap", None)
        rows.append(row)

    return sort_analytics_rows(rows)


def summarize_keynote_analytics_rows(rows):
    sheet_keys = set()
    placed_key_count = 0
    placed_count = 0
    user_keynote_count = 0
    generic_annotation_count = 0
    unsheeted_count = 0
    orphan_key_count = 0
    placed_key_map = {}

    for row in rows or []:
        key = safe_unicode(row.get("keynoteKey")).strip()
        if row.get("placed"):
            placed_key_count += 1
            if key:
                placed_key_map[key] = True
        if row.get("placed") and not row.get("inLibrary"):
            orphan_key_count += 1
        placed_count += int(row.get("placedCount") or 0)
        user_keynote_count += int(row.get("userKeynoteCount") or 0)
        generic_annotation_count += int(row.get("genericAnnotationCount") or 0)
        unsheeted_count += int(row.get("unsheetedCount") or 0)
        for sheet in row.get("sheets") or []:
            sheet_key = sheet_analytics_key(sheet)
            if sheet_key:
                sheet_keys.add(sheet_key)

    return {
        "placedKeyCount": placed_key_count,
        "placedCount": placed_count,
        "userKeynoteCount": user_keynote_count,
        "genericAnnotationCount": generic_annotation_count,
        "sheetCount": len(sheet_keys),
        "unsheetedCount": unsheeted_count,
        "orphanKeyCount": orphan_key_count,
        "placedKeyMap": placed_key_map,
    }


def collect_keynote_analytics(target_doc, keynote_payload):
    keynote_payload = keynote_payload or {}
    entry_by_key = {}
    rows_by_key = {}
    view_sheet_lookup = build_view_sheet_lookup(target_doc)

    for entry in keynote_payload.get("entries") or []:
        key = safe_unicode(entry.get("key")).strip()
        if not key:
            continue
        entry_by_key[key] = entry
        if key not in rows_by_key:
            rows_by_key[key] = make_keynote_analytics_row(key, entry)

    user_keynote_scanned_count = 0
    generic_annotation_scanned_count = 0
    skipped_count = 0

    tags = (
        FilteredElementCollector(target_doc)
        .OfCategory(BuiltInCategory.OST_KeynoteTags)
        .WhereElementIsNotElementType()
    )
    for tag in tags:
        key = safe_unicode(get_keynote_tag_key(tag)).strip()
        if record_keynote_analytics_placement(
            rows_by_key,
            entry_by_key,
            key,
            "userKeynote",
            tag,
            target_doc,
            view_sheet_lookup
        ):
            user_keynote_scanned_count += 1
        else:
            skipped_count += 1

    for instance, symbol in iter_generic_annotation_keynote_instances(target_doc):
        key = safe_unicode(get_generic_annotation_symbol_key(symbol)).strip()
        if record_keynote_analytics_placement(
            rows_by_key,
            entry_by_key,
            key,
            "genericAnnotation",
            instance,
            target_doc,
            view_sheet_lookup
        ):
            generic_annotation_scanned_count += 1
        else:
            skipped_count += 1

    rows = finalize_keynote_analytics_rows(rows_by_key)
    summary = summarize_keynote_analytics_rows(rows)
    identity = get_document_analytics_identity(target_doc)

    analytics = {
        "libraryKey": keynote_payload.get("libraryKey") or "",
        "displayPath": keynote_payload.get("displayPath") or keynote_payload.get("keynotePath") or "",
        "keynotePath": keynote_payload.get("keynotePath") or "",
        "encoding": keynote_payload.get("encoding") or "utf-8",
        "lineEnding": keynote_payload.get("lineEnding") or "\r\n",
        "fileHash": keynote_payload.get("fileHash") or "",
        "lastWriteUtc": keynote_payload.get("lastWriteUtc"),
        "entries": keynote_payload.get("entries") or [],
        "entryCount": len(keynote_payload.get("entries") or []),
        "analyticsRows": rows,
        "analyticsRowCount": len(rows),
        "collectedAt": get_generated_at(),
        "userKeynoteScannedCount": user_keynote_scanned_count,
        "genericAnnotationScannedCount": generic_annotation_scanned_count,
        "skippedCount": skipped_count,
    }
    analytics.update(identity)
    analytics.update(summary)
    return analytics


def make_model_health_issue(severity, code, key, message, row=None, details="", type_names=None, resolution=None):
    row = row or {}
    issue = {
        "severity": severity or "warning",
        "code": code or "modelHealthIssue",
        "key": safe_unicode(key).strip(),
        "message": safe_unicode(message).strip(),
        "details": safe_unicode(details).strip(),
        "placedCount": int(row.get("placedCount") or 0),
        "userKeynoteCount": int(row.get("userKeynoteCount") or 0),
        "genericAnnotationCount": int(row.get("genericAnnotationCount") or 0),
        "sheetCount": int(row.get("sheetCount") or 0),
        "unsheetedCount": int(row.get("unsheetedCount") or 0),
        "sheets": row.get("sheets") or [],
    }
    if type_names:
        issue["typeNames"] = [safe_unicode(name).strip() for name in type_names if safe_unicode(name).strip()]
    if resolution:
        issue["resolution"] = {
            "resolutionType": safe_unicode(resolution.get("resolutionType")).strip(),
            "familyTypeName": safe_unicode(resolution.get("familyTypeName")).strip(),
            "familyTypeText": safe_unicode(resolution.get("familyTypeText")).strip(),
            "fileText": safe_unicode(resolution.get("fileText")).strip(),
        }
    return issue


def get_symbol_parameter_health_state(symbol, parameter_name):
    result = {
        "present": False,
        "readOnly": False,
        "value": "",
    }

    try:
        parameter = symbol.LookupParameter(parameter_name)
    except:
        parameter = None

    if parameter is None:
        return result

    result["present"] = True
    result["value"] = get_parameter_text(parameter)
    try:
        result["readOnly"] = bool(parameter.IsReadOnly)
    except:
        result["readOnly"] = False

    return result


def append_generic_annotation_model_health_issues(target_doc, entry_by_key, issues):
    family = get_generic_annotation_keynote_family(target_doc)
    if family is None:
        return

    symbols = get_family_symbols(target_doc, family)
    if not symbols:
        issues.append(make_model_health_issue(
            "warning",
            "genericAnnotationFamilyHasNoTypes",
            "",
            "Generic Annotation family '{0}' does not contain any types.".format(GENERIC_KEYNOTE_FAMILY_NAME)
        ))
        return

    if get_leader_arrowhead_type(target_doc) is None:
        issues.append(make_model_health_issue(
            "warning",
            "genericAnnotationLeaderArrowheadMissing",
            "",
            "Arrowhead type '{0}' was not found for Generic Annotation keynote types.".format(
                GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME
            )
        ))

    instances_by_symbol = collect_generic_annotation_instances_by_symbol(target_doc)
    symbols_by_key = {}

    for symbol in symbols:
        type_name = get_element_name(symbol)
        number_state = get_symbol_parameter_health_state(symbol, GENERIC_KEYNOTE_NUMBER_PARAMETER)
        text_state = get_symbol_parameter_health_state(symbol, GENERIC_KEYNOTE_TEXT_PARAMETER)
        symbol_key = safe_unicode(number_state.get("value")).strip() or type_name
        instance_count = len(instances_by_symbol.get(get_element_id_key(symbol)) or [])
        row = {
            "placedCount": instance_count,
            "genericAnnotationCount": instance_count,
        }

        if symbol_key:
            if symbol_key not in symbols_by_key:
                symbols_by_key[symbol_key] = []
            symbols_by_key[symbol_key].append(symbol)

        if not number_state.get("present"):
            issues.append(make_model_health_issue(
                "warning",
                "genericAnnotationMissingNumberParameter",
                symbol_key,
                "Generic Annotation type '{0}' is missing the '{1}' parameter.".format(
                    type_name,
                    GENERIC_KEYNOTE_NUMBER_PARAMETER
                ),
                row
            ))
        elif number_state.get("readOnly"):
            issues.append(make_model_health_issue(
                "warning",
                "genericAnnotationReadOnlyNumberParameter",
                symbol_key,
                "Generic Annotation type '{0}' has a read-only '{1}' parameter.".format(
                    type_name,
                    GENERIC_KEYNOTE_NUMBER_PARAMETER
                ),
                row
            ))

        if not text_state.get("present"):
            issues.append(make_model_health_issue(
                "warning",
                "genericAnnotationMissingTextParameter",
                symbol_key,
                "Generic Annotation type '{0}' is missing the '{1}' parameter.".format(
                    type_name,
                    GENERIC_KEYNOTE_TEXT_PARAMETER
                ),
                row
            ))
        elif text_state.get("readOnly"):
            issues.append(make_model_health_issue(
                "warning",
                "genericAnnotationReadOnlyTextParameter",
                symbol_key,
                "Generic Annotation type '{0}' has a read-only '{1}' parameter.".format(
                    type_name,
                    GENERIC_KEYNOTE_TEXT_PARAMETER
                ),
                row
            ))

        if number_state.get("value") and type_name and number_state.get("value") != type_name:
            issues.append(make_model_health_issue(
                "warning",
                "genericAnnotationTypeNameMismatch",
                symbol_key,
                "Generic Annotation type '{0}' has Number '{1}'.".format(
                    type_name,
                    number_state.get("value")
                ),
                row,
                "The type name and Number parameter should match the keynote key."
            ))

        if symbol_key in entry_by_key and text_state.get("present"):
            expected_text = safe_unicode(entry_by_key[symbol_key].get("text")).strip()
            current_text = safe_unicode(text_state.get("value")).strip()
            if expected_text != current_text:
                issues.append(make_model_health_issue(
                    "warning",
                    "genericAnnotationTextMismatch",
                    symbol_key,
                    "Generic Annotation type '{0}' text does not match the keynote file.".format(type_name),
                    row,
                    "Family type text: {0} | Text file: {1}".format(current_text, expected_text)
                ))

    for symbol_key, key_symbols in symbols_by_key.items():
        if len(key_symbols) < 2:
            continue
        type_names = [get_element_name(symbol) for symbol in key_symbols]
        issues.append(make_model_health_issue(
            "warning",
            "genericAnnotationDuplicateTypes",
            symbol_key,
            "{0} Generic Annotation types match keynote key '{1}'.".format(len(key_symbols), symbol_key),
            {
                "placedCount": sum([len(instances_by_symbol.get(get_element_id_key(symbol)) or []) for symbol in key_symbols]),
                "genericAnnotationCount": sum([len(instances_by_symbol.get(get_element_id_key(symbol)) or []) for symbol in key_symbols]),
            },
            "Types: {0}".format(", ".join(type_names)),
            type_names
        ))


def model_health_issue_sort_key(issue):
    severity = safe_unicode(issue.get("severity")).lower()
    severity_rank = severity == "error" and "0" or "1"
    return "{0}|{1}|{2}".format(
        severity_rank,
        safe_unicode(issue.get("code")).lower(),
        safe_unicode(issue.get("key")).lower()
    )


def make_model_health_signature(analytics, issues, keynote_payload):
    signature_rows = []
    for issue in issues or []:
        signature_rows.append({
            "severity": issue.get("severity") or "",
            "code": issue.get("code") or "",
            "key": issue.get("key") or "",
            "placedCount": issue.get("placedCount") or 0,
            "genericAnnotationCount": issue.get("genericAnnotationCount") or 0,
            "typeNames": issue.get("typeNames") or [],
            "resolution": issue.get("resolution") or {},
        })

    source = json.dumps({
        "documentKey": analytics.get("documentKey") or "",
        "libraryKey": keynote_payload.get("libraryKey") or analytics.get("libraryKey") or "",
        "fileHash": keynote_payload.get("fileHash") or analytics.get("fileHash") or "",
        "placedKeyCount": analytics.get("placedKeyCount") or 0,
        "orphanKeyCount": analytics.get("orphanKeyCount") or 0,
        "issues": signature_rows,
    }, ensure_ascii=True, sort_keys=True)
    return hashlib.sha256(source.encode("utf-8")).hexdigest()


def collect_generic_annotation_resolution_sources(target_doc):
    """Collect the best family-type source for each Generic Annotation keynote key."""
    result = {}
    family = get_generic_annotation_keynote_family(target_doc)
    if family is None:
        return result

    instances_by_symbol = collect_generic_annotation_instances_by_symbol(target_doc)
    for symbol in get_family_symbols(target_doc, family):
        key = safe_unicode(get_generic_annotation_symbol_key(symbol)).strip()
        if not key:
            continue

        symbol_id_key = get_element_id_key(symbol)
        candidate = {
            "resolutionType": "missingGenericAnnotationKey",
            "familyTypeName": get_element_name(symbol),
            "familyTypeText": get_lookup_parameter_text(symbol, [GENERIC_KEYNOTE_TEXT_PARAMETER]),
            "instanceCount": len(instances_by_symbol.get(symbol_id_key) or []),
        }
        existing = result.get(key)
        if existing is None or candidate["instanceCount"] > existing.get("instanceCount", 0):
            result[key] = candidate

    return result


def build_model_health_from_analytics(target_doc, keynote_payload, analytics):
    keynote_payload = keynote_payload or {}
    analytics = analytics or {}
    issues = []
    entry_by_key = {}

    for entry in keynote_payload.get("entries") or []:
        key = safe_unicode(entry.get("key")).strip()
        if key:
            entry_by_key[key] = entry

    generic_annotation_sources = collect_generic_annotation_resolution_sources(target_doc)
    missing_key_count = 0
    missing_placed_count = 0
    for row in analytics.get("analyticsRows") or []:
        if not row.get("placed") or row.get("inLibrary"):
            continue
        missing_key_count += 1
        missing_placed_count += int(row.get("placedCount") or 0)
        missing_key = safe_unicode(row.get("keynoteKey")).strip()
        resolution = None
        if int(row.get("genericAnnotationCount") or 0):
            resolution = generic_annotation_sources.get(missing_key)
        issues.append(make_model_health_issue(
            "error",
            "placedKeyMissingFromLibrary",
            missing_key,
            "Placed keynote key '{0}' was not found in the active keynote file.".format(
                missing_key
            ),
            row,
            "Choose the family type to add this keynote to the text file, or choose a text-file keynote to replace the family type.",
            None,
            resolution
        ))

    append_generic_annotation_model_health_issues(target_doc, entry_by_key, issues)
    issues = sorted(issues, key=model_health_issue_sort_key)

    placed_key_count = int(analytics.get("placedKeyCount") or 0)
    missing_ratio = placed_key_count and (float(missing_key_count) / float(placed_key_count)) or 0
    safe_mode_recommended = (
        missing_key_count >= MODEL_HEALTH_SAFE_MODE_MISSING_KEY_COUNT or
        (
            missing_key_count >= MODEL_HEALTH_SAFE_MODE_RATIO_MIN_MISSING_KEYS and
            placed_key_count >= MODEL_HEALTH_SAFE_MODE_RATIO_MIN_PLACED_KEYS and
            missing_ratio >= MODEL_HEALTH_SAFE_MODE_RATIO
        )
    )
    status = safe_mode_recommended and "safeMode" or (issues and "issues" or "ready")
    message = "Model health scan found no model/keynote mismatches."
    if issues:
        message = "Model health scan found {0} issue(s).".format(len(issues))
    if safe_mode_recommended:
        message = "Safe Mode recommended: {0} placed keynote key(s) are missing from the active keynote file.".format(
            missing_key_count
        )

    health = make_empty_model_health(status, message)
    health.update({
        "scannedAt": analytics.get("collectedAt") or get_generated_at(),
        "safeModeRecommended": bool(safe_mode_recommended),
        "placedKeyCount": placed_key_count,
        "placedCount": int(analytics.get("placedCount") or 0),
        "missingKeyCount": missing_key_count,
        "missingPlacedCount": missing_placed_count,
        "missingRatio": missing_ratio,
        "userKeynoteCount": int(analytics.get("userKeynoteCount") or 0),
        "genericAnnotationCount": int(analytics.get("genericAnnotationCount") or 0),
        "sheetCount": int(analytics.get("sheetCount") or 0),
        "unsheetedCount": int(analytics.get("unsheetedCount") or 0),
        "skippedCount": int(analytics.get("skippedCount") or 0),
        "placedKeyMap": analytics.get("placedKeyMap") or {},
        "issues": issues,
    })
    health["signature"] = make_model_health_signature(analytics, issues, keynote_payload)
    return health


def build_model_health(target_doc, keynote_payload, analytics=None):
    try:
        analytics = analytics or collect_keynote_analytics(target_doc, keynote_payload)
        return build_model_health_from_analytics(target_doc, keynote_payload, analytics)
    except Exception as exc:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass
        health = make_empty_model_health(
            "error",
            "Could not scan model health: {0}".format(safe_str(exc))
        )
        health["issues"] = [make_model_health_issue(
            "warning",
            "modelHealthScanFailed",
            "",
            health["message"]
        )]
        health["signature"] = hashlib.sha256(health["message"].encode("utf-8")).hexdigest()
        return health


def collect_keynote_analytics_payload(target_doc):
    try:
        keynote_payload = build_keynote_payload(target_doc, include_model_health=False)
        if not keynote_payload.get("libraryKey"):
            return {
                "status": "error",
                "message": keynote_payload.get("message") or "No keynote library is available for analytics.",
                "issues": keynote_payload.get("issues") or [],
            }

        if has_error_issues(keynote_payload.get("issues") or []):
            return {
                "status": "error",
                "message": "Fix keynote file validation errors before collecting analytics.",
                "issues": keynote_payload.get("issues") or [],
            }

        analytics = collect_keynote_analytics(target_doc, keynote_payload)
        message = (
            "Collected analytics for {0} placed keynote key(s), {1} placed annotation(s), and {2} sheet(s)."
        ).format(
            analytics.get("placedKeyCount", 0),
            analytics.get("placedCount", 0),
            analytics.get("sheetCount", 0)
        )
        if analytics.get("orphanKeyCount"):
            message += " Found {0} placed key(s) that are not in the keynote file.".format(
                analytics.get("orphanKeyCount")
            )

        model_health = build_model_health(target_doc, keynote_payload, analytics)

        return {
            "status": "ready",
            "message": message,
            "analytics": analytics,
            "modelHealth": model_health,
        }
    except Exception as exc:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass
        return {
            "status": "error",
            "message": "Could not collect keynote analytics: {0}".format(safe_str(exc)),
            "issues": [make_issue("error", safe_str(exc), "", None, "analyticsCollectionFailed")],
        }


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


def make_model_issue_key_rename_map(model_issue_resolutions, desired_entries):
    """Build family/reference migrations selected from Safe Mode missing-key errors."""
    desired_by_key = keyed_entries(desired_entries)
    key_renames = {}

    for resolution in model_issue_resolutions or []:
        if safe_str(resolution.get("source")).strip() != "textFile":
            continue
        old_key = safe_unicode(resolution.get("issueKey")).strip()
        new_key = safe_unicode(resolution.get("replacementKey")).strip()
        if old_key and new_key and old_key != new_key and new_key in desired_by_key:
            key_renames[old_key] = new_key

    return key_renames


def make_deleted_key_set(baseline_entries, desired_entries):
    desired_by_id = indexed_entries(desired_entries)
    deleted_keys = set()

    for baseline in baseline_entries or []:
        base_entry = normalize_merge_entry(baseline)
        if base_entry["id"] and base_entry["id"] not in desired_by_id and base_entry["key"]:
            deleted_keys.add(base_entry["key"])

    return deleted_keys


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
        model_issue_key_renames = make_model_issue_key_rename_map(
            save_payload.get("modelIssueResolutions") or [],
            entries
        )
        for old_key, new_key in model_issue_key_renames.items():
            if old_key not in key_renames:
                key_renames[old_key] = new_key
        deleted_keys = make_deleted_key_set(baseline_entries, entries)
        reference_update = {}
        generic_annotation_sync = make_generic_annotation_sync_summary()
        lock_path = get_sync_lock_path(current_path)
        backup_path = ""
        merged_entries = []

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

            keynote_text = canonicalize_entries(merged_entries, line_ending, current_path, encoding)
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

        try:
            generic_annotation_sync = sync_generic_annotation_types(
                target_doc,
                merged_entries,
                key_renames,
                deleted_keys
            )
        except Exception as exc:
            generic_annotation_sync = make_generic_annotation_sync_summary()
            record_generic_annotation_sync_failure(
                generic_annotation_sync,
                "Generic Annotation synchronization could not start: {0}".format(exc)
            )
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
        if generic_annotation_sync.get("preservedCount"):
            result_issues.append(make_issue(
                "warning",
                "{0} Generic Annotation keynote type(s) were preserved because they are still in use or could not be removed.".format(
                    generic_annotation_sync.get("preservedCount")
                ),
                "",
                None,
                "genericAnnotationTypePreserved"
            ))
        if generic_annotation_sync.get("failedCount"):
            result_issues.append(make_issue(
                "warning",
                "{0} Generic Annotation keynote synchronization action(s) could not be completed.".format(
                    generic_annotation_sync.get("failedCount")
                ),
                "",
                None,
                "genericAnnotationSyncFailed"
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
                "genericAnnotationSync": generic_annotation_sync,
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
        if generic_annotation_sync.get("updatedCount") or generic_annotation_sync.get("deletedCount"):
            message += " Synchronized {0} and removed {1} Generic Annotation keynote type(s).".format(
                generic_annotation_sync.get("updatedCount", 0),
                generic_annotation_sync.get("deletedCount", 0)
            )

        return {
            "status": "ready",
            "message": message,
            "backupPath": backup_path,
            "keyRenames": key_renames,
            "keynoteReferenceUpdate": reference_update,
            "genericAnnotationSync": generic_annotation_sync,
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


def get_keynote_payload_entry(keynote_payload, entry_id, key):
    entry_id = safe_str(entry_id).strip()
    key = safe_unicode(key).strip()
    matched_by_key = None
    if not key:
        return None

    for entry in (keynote_payload or {}).get("entries") or []:
        entry_key = safe_unicode(entry.get("key")).strip()
        if entry_key == key and matched_by_key is None:
            matched_by_key = entry
        if entry_id:
            if safe_str(entry.get("id")).strip() == entry_id and entry_key == key:
                return entry
        elif entry_key == key:
            return entry
    return matched_by_key


def keynote_payload_has_entry_key(keynote_payload, entry_id, key):
    return get_keynote_payload_entry(keynote_payload, entry_id, key) is not None


def copy_text_to_clipboard(value):
    try:
        Clipboard.SetText(safe_unicode(value))
        return True, ""
    except Exception as exc:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass
        return False, safe_str(exc)


def get_keynote_tag_default_type(target_doc):
    try:
        category = target_doc.Settings.Categories.get_Item(BuiltInCategory.OST_KeynoteTags)
    except:
        category = None

    if not category:
        return None

    try:
        return target_doc.GetElement(target_doc.GetDefaultFamilyTypeId(category.Id))
    except:
        return None


def place_user_keynote(uiapp, target_doc, place_payload, keynote_payload):
    place_payload = place_payload or {}
    entry_id = safe_str(place_payload.get("id")).strip()
    key = safe_unicode(place_payload.get("key")).strip()
    clipboard_copied = False
    clipboard_error = ""
    command_id = None

    if not key:
        return {
            "status": "warning",
            "message": "Save this keynote with a key before placing it in Revit.",
        }

    if not keynote_payload_has_entry_key(keynote_payload, entry_id, key):
        return {
            "status": "warning",
            "message": "This keynote is no longer available in the shared keynote file. Refresh before placing it in Revit.",
        }

    clipboard_copied, clipboard_error = copy_text_to_clipboard(key)

    if DocumentEventUtils is None:
        fallback_message = "Could not place keynote automatically because pyRevit's document event helper was not available."
        if clipboard_copied:
            fallback_message += " Key copied to clipboard."
        elif clipboard_error:
            fallback_message += " Clipboard copy was unavailable."
        return {
            "status": "warning",
            "message": fallback_message,
        }

    if get_keynote_tag_default_type(target_doc) is None:
        return {
            "status": "warning",
            "message": "No default Keynote Tag type is loaded in this project.",
        }

    try:
        reload_revit_keynotes(target_doc)
    except Exception as exc:
        return {
            "status": "warning",
            "message": "Could not reload the current keynote table before placement: {0}".format(safe_str(exc)),
        }

    try:
        command_id = RevitCommandId.LookupPostableCommandId(PostableCommand.UserKeynote)
    except Exception as exc:
        return {
            "status": "warning",
            "message": "Could not find Revit's User Keynote command: {0}".format(safe_str(exc)),
        }

    if command_id is None:
        return {
            "status": "warning",
            "message": "Revit's User Keynote command was not available.",
        }

    try:
        if hasattr(uiapp, "CanPostCommand") and not uiapp.CanPostCommand(command_id):
            return {
                "status": "warning",
                "message": "Revit cannot start User Keynote placement in the current context.",
            }
    except:
        pass

    try:
        DocumentEventUtils.PostCommandAndUpdateNewElementProperties(
            uiapp,
            target_doc,
            PostableCommand.UserKeynote,
            "Set User Keynote Key",
            BuiltInParameter.KEY_VALUE,
            key
        )
    except Exception as exc:
        return {
            "status": "warning",
            "message": "Could not start Revit User Keynote placement: {0}".format(safe_str(exc)),
        }

    message = "Starting Revit User Keynote placement for key '{0}'.".format(key)
    message += " The placed keynote will use the selected row key."
    if clipboard_error:
        message += " Clipboard copy was unavailable, but automatic key assignment was started."

    return {
        "status": "ready",
        "message": message,
    }


def get_active_uidocument(uiapp, target_doc):
    try:
        uidoc = uiapp.ActiveUIDocument
    except:
        uidoc = None

    if uidoc is None:
        return None, "No active Revit document is available for Generic Annotation placement."

    try:
        if uidoc.Document != target_doc:
            return None, "Activate the Revit document associated with this Keynote Manager before placing a Generic Annotation."
    except:
        return None, "Could not confirm the active Revit document for Generic Annotation placement."

    return uidoc, ""


def get_valid_generic_annotation_base_symbol(symbols):
    for symbol in symbols or []:
        try:
            get_writable_symbol_parameter(symbol, GENERIC_KEYNOTE_NUMBER_PARAMETER)
            get_writable_symbol_parameter(symbol, GENERIC_KEYNOTE_TEXT_PARAMETER)
            get_writable_builtin_symbol_parameter(
                symbol,
                BuiltInParameter.LEADER_ARROWHEAD,
                "Leader Arrowhead"
            )
            return symbol
        except:
            continue
    return None


def ensure_generic_annotation_keynote_symbol(target_doc, key, text):
    summary = make_generic_annotation_sync_summary()
    family = get_generic_annotation_keynote_family(target_doc)
    if family is None:
        return None, summary, "Load the Generic Annotation family '{0}' before placing symbol keynotes.".format(
            GENERIC_KEYNOTE_FAMILY_NAME
        )

    summary["familyFound"] = True
    symbols = get_family_symbols(target_doc, family)
    if not symbols:
        return None, summary, "Family '{0}' does not contain a type to duplicate.".format(
            GENERIC_KEYNOTE_FAMILY_NAME
        )

    leader_arrowhead_type = get_leader_arrowhead_type(target_doc)
    if leader_arrowhead_type is None:
        return None, summary, "Load or create the Arrowhead type '{0}' before placing symbol keynotes.".format(
            GENERIC_KEYNOTE_LEADER_ARROWHEAD_NAME
        )

    symbol = choose_generic_annotation_symbol(symbols, key)
    transaction = Transaction(target_doc, "Prepare Generic Annotation Keynote")

    try:
        transaction.Start()

        if symbol is None:
            base_symbol = get_valid_generic_annotation_base_symbol(symbols)
            if base_symbol is None:
                raise Exception(
                    "No type in family '{0}' has writable '{1}' and '{2}' parameters.".format(
                        GENERIC_KEYNOTE_FAMILY_NAME,
                        GENERIC_KEYNOTE_NUMBER_PARAMETER,
                        GENERIC_KEYNOTE_TEXT_PARAMETER
                    )
                )
            symbol = duplicate_family_symbol(target_doc, base_symbol, key)
            if symbol is None:
                raise Exception("Revit did not return the duplicated Generic Annotation type.")
            summary["createdCount"] += 1

        changed, migrated = set_generic_annotation_symbol_values(
            symbol,
            key,
            text,
            leader_arrowhead_type
        )
        if changed:
            summary["updatedCount"] += 1
        if migrated:
            summary["migratedCount"] += 1

        try:
            if not symbol.IsActive:
                symbol.Activate()
                target_doc.Regenerate()
        except Exception as exc:
            raise Exception("Could not activate Generic Annotation type '{0}': {1}".format(key, exc))

        transaction.Commit()
    except Exception as exc:
        try:
            transaction.RollBack()
        except:
            pass
        return None, summary, safe_str(exc)

    return symbol, summary, ""


def place_generic_annotation_keynote(uiapp, target_doc, place_payload, keynote_payload):
    place_payload = place_payload or {}
    entry_id = safe_str(place_payload.get("id")).strip()
    key = safe_unicode(place_payload.get("key")).strip()

    if not key:
        return {
            "status": "warning",
            "message": "Save this keynote with a key before placing it in Revit.",
        }

    entry = get_keynote_payload_entry(keynote_payload, entry_id, key)
    if entry is None:
        return {
            "status": "warning",
            "message": "This keynote is no longer available in the shared keynote file. Refresh before placing it in Revit.",
        }

    uidoc, active_document_error = get_active_uidocument(uiapp, target_doc)
    if uidoc is None:
        return {
            "status": "warning",
            "message": active_document_error,
        }

    text = safe_unicode(entry.get("text")).strip()
    symbol, summary, prepare_error = ensure_generic_annotation_keynote_symbol(target_doc, key, text)
    if symbol is None:
        return {
            "status": "warning",
            "message": "Could not prepare Generic Annotation keynote '{0}': {1}".format(key, prepare_error),
            "genericAnnotationSync": summary,
        }

    try:
        if hasattr(uidoc, "CanPlaceElementType") and not uidoc.CanPlaceElementType(symbol):
            return {
                "status": "warning",
                "message": "Revit cannot place Generic Annotation keynote '{0}' in the active view.".format(key),
                "genericAnnotationSync": summary,
            }
    except:
        pass

    try:
        uidoc.PromptForFamilyInstancePlacement(symbol)
    except Exception as exc:
        try:
            exception_name = safe_str(exc.GetType().Name)
        except:
            exception_name = exc.__class__.__name__

        if exception_name == "OperationCanceledException":
            return {
                "status": "ready",
                "message": "Finished Generic Annotation keynote placement for key '{0}'.".format(key),
                "genericAnnotationSync": summary,
            }

        return {
            "status": "warning",
            "message": "Could not place Generic Annotation keynote '{0}': {1}".format(key, safe_str(exc)),
            "genericAnnotationSync": summary,
        }

    return {
        "status": "ready",
        "message": "Finished Generic Annotation keynote placement for key '{0}'.".format(key),
        "genericAnnotationSync": summary,
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

    def queue_collect_analytics(self):
        self.pending_action = "collectAnalytics"
        self.pending_payload = None

    def queue_place_user_keynote(self, payload):
        self.pending_action = "placeUserKeynote"
        self.pending_payload = payload

    def queue_place_generic_annotation(self, payload):
        self.pending_action = "placeGenericAnnotation"
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

        if action == "collectAnalytics":
            result = collect_keynote_analytics_payload(window.document)
            window.send_analytics_result(result)
            return

        if action == "placeUserKeynote":
            keynote_payload = build_keynote_payload(window.document)
            result = place_user_keynote(uiapp, window.document, payload, keynote_payload)
            window.send_status(result.get("status"), result.get("message"))
            return

        if action == "placeGenericAnnotation":
            keynote_payload = build_keynote_payload(window.document)
            result = place_generic_annotation_keynote(uiapp, window.document, payload, keynote_payload)
            window.send_status(result.get("status"), result.get("message"))
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
        state = read_user_settings()
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
            state = read_user_settings()
            state.update({
                "width": float(self.Width),
                "height": float(self.Height),
                "left": float(self.Left),
                "top": float(self.Top),
            })
            write_json_file(get_settings_path(), state)
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

    def send_analytics_result(self, result):
        self.call_keynote_app("handleAnalyticsResult", result or {})

    def request_refresh_from_app(self):
        self.call_keynote_app("requestRefresh", {})

    def raise_external_event(self, action_name, failure_target="save"):
        try:
            self.external_event.Raise()
        except Exception as exc:
            self.event_handler.clear_pending()
            result = {
                "status": "error",
                "message": "Could not raise the Revit {0} event: {1}".format(action_name, exc),
            }
            if failure_target == "analytics":
                self.send_analytics_result(result)
            else:
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

        if message_type == "placementModeChanged":
            placement_mode = save_placement_mode_setting(message.get("placementMode"))
            try:
                preferences = self.keynote_payload.get("preferences") or {}
                preferences["placementMode"] = placement_mode
                self.keynote_payload["preferences"] = preferences
            except:
                pass
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

        if message_type == "collectAnalytics":
            self.event_handler.queue_collect_analytics()
            self.send_status("warning", "Collecting keynote analytics from the active Revit document...")
            self.raise_external_event("collect analytics", "analytics")
            return

        if message_type == "placeUserKeynote":
            place_payload = message.get("payload") or {}
            key = safe_unicode(place_payload.get("key")).strip()
            self.event_handler.queue_place_user_keynote(place_payload)
            self.send_status(
                "warning",
                "Starting Revit User Keynote placement{0}...".format(
                    key and " for key '{0}'".format(key) or ""
                )
            )
            self.raise_external_event("place user keynote")
            return

        if message_type == "placeGenericAnnotation":
            place_payload = message.get("payload") or {}
            key = safe_unicode(place_payload.get("key")).strip()
            self.event_handler.queue_place_generic_annotation(place_payload)
            self.send_status(
                "warning",
                "Preparing Generic Annotation keynote placement{0}...".format(
                    key and " for key '{0}'".format(key) or ""
                )
            )
            self.raise_external_event("place generic annotation keynote")
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
