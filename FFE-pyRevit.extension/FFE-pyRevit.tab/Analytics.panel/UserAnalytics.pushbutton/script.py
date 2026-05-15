# -*- coding: utf-8 -*-
__title__ = "User\nAnalytics"
__version__ = "Version = v1.0"
__persistentengine__ = True
__doc__ = """Version = v1.0
Date    = 05.15.2026
__________________________________________________________________
Description:
Shows a read-only WebView2 dashboard for the current Revit user's
FFE-pyRevit usage log.
__________________________________________________________________
How-To:
- Click the button to open the current user's usage dashboard.
__________________________________________________________________
Last update:
- [10.22.2025] - v0.1 Beta Release
- [05.15.2026] - v1.0 Added WebView2 user analytics dashboard
__________________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS
import json
import os
import time
import traceback

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock

from pyrevit import forms


# ____________________________________________________________________ VARIABLES
revit_app = __revit__.Application

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

ANALYTICS_NAME = "FFE-pyRevit User Analytics"
ANALYTICS_VERSION = "v1.0"

try:
    WINDOW_REFS
except NameError:
    WINDOW_REFS = []


# ____________________________________________________________________ HELPERS
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


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True)


def get_revit_username():
    try:
        username = safe_str(revit_app.Username)
        if username:
            return username
    except:
        pass

    try:
        username = safe_str(os.environ.get("USERNAME"))
        if username:
            return username
    except:
        pass

    return "Unknown User"


def get_logs_dir():
    return os.path.join(
        os.path.expanduser("~"),
        "FFE Inc",
        "FFE Revit Users - Documents",
        "00-General",
        "Revit_Add-Ins",
        "FFE-pyRevit",
        "Logs"
    )


def get_log_file_path(username):
    return os.path.join(get_logs_dir(), "{0}_revit_log.json".format(username))


def get_generated_at():
    return time.strftime("%Y-%m-%d %H:%M:%S")


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


def get_webview_user_data_folder():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")

    user_data_folder = os.path.join(base_folder, "FFE-pyRevit", "UserAnalyticsWebView2")
    if not os.path.exists(user_data_folder):
        os.makedirs(user_data_folder)
    return user_data_folder


def build_payload(status, message, entries):
    username = get_revit_username()
    log_dir = get_logs_dir()
    log_path = get_log_file_path(username)

    return {
        "name": ANALYTICS_NAME,
        "version": ANALYTICS_VERSION,
        "username": username,
        "logDir": log_dir,
        "logPath": log_path,
        "generatedAt": get_generated_at(),
        "status": status,
        "message": message,
        "entries": entries or [],
        "entryCount": len(entries or []),
    }


def normalize_entry(entry, index, fallback_username):
    if not isinstance(entry, dict):
        return None

    action = safe_str(entry.get("action")).strip()
    if not action:
        action = "(Unknown)"

    normalized = {
        "index": index,
        "datetime": safe_str(entry.get("datetime")),
        "username": safe_str(entry.get("username") or fallback_username),
        "doc_title": safe_str(entry.get("doc_title")),
        "doc_path": safe_str(entry.get("doc_path")),
        "revit_version_number": safe_str(entry.get("revit_version_number")),
        "revit_build": safe_str(entry.get("revit_build")),
        "action": action,
        "status": safe_str(entry.get("status")),
        "family_name": safe_str(entry.get("family_name")),
        "family_origin": safe_str(entry.get("family_origin")),
        "family_path": safe_str(entry.get("family_path")),
    }

    return normalized


def load_user_analytics_payload():
    """Read and normalize the current Revit user's log without writing to it."""
    username = get_revit_username()
    log_dir = get_logs_dir()
    log_path = get_log_file_path(username)

    if not os.path.exists(log_dir):
        return build_payload(
            "missingFolder",
            "The FFE-pyRevit Logs folder was not found.",
            []
        )

    if not os.path.exists(log_path):
        return build_payload(
            "missingFile",
            "No usage log was found for the current Revit user.",
            []
        )

    try:
        with open(log_path, "r") as file_obj:
            raw_text = file_obj.read()
    except Exception as exc:
        return build_payload(
            "readError",
            "Could not read the current user's usage log: {0}".format(exc),
            []
        )

    if not raw_text or not raw_text.strip():
        return build_payload(
            "emptyLog",
            "The current user's usage log is empty.",
            []
        )

    try:
        log_data = json.loads(raw_text)
    except Exception as exc:
        return build_payload(
            "invalidJson",
            "The current user's usage log is not valid JSON: {0}".format(exc),
            []
        )

    if isinstance(log_data, dict):
        raw_entries = log_data.get("action")
    else:
        raw_entries = None

    if raw_entries is None:
        return build_payload(
            "invalidSchema",
            "The usage log does not contain the expected top-level action array.",
            []
        )

    if not isinstance(raw_entries, list):
        return build_payload(
            "invalidSchema",
            "The usage log action value is not an array.",
            []
        )

    entries = []
    for index, raw_entry in enumerate(raw_entries):
        normalized = normalize_entry(raw_entry, index, username)
        if normalized is not None:
            entries.append(normalized)

    if not entries:
        return build_payload(
            "emptyLog",
            "The current user's usage log does not contain any readable entries.",
            []
        )

    return build_payload(
        "ready",
        "Loaded {0} usage entries.".format(len(entries)),
        entries
    )


def focus_existing_window():
    for window in list(WINDOW_REFS):
        try:
            if window.IsVisible:
                window.Activate()
                window.send_analytics_payload(force=True)
                return True
        except:
            try:
                WINDOW_REFS.remove(window)
            except:
                pass
    return False


# ____________________________________________________________________ WEBVIEW WINDOW
class UserAnalyticsWindow(Window):
    def __init__(self, webview_type, creation_properties_type):
        Window.__init__(self)

        self.has_sent_analytics_payload = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = "{0} - {1}".format(ANALYTICS_NAME, get_revit_username())
        self.Width = 1280
        self.Height = 840
        self.MinWidth = 980
        self.MinHeight = 640
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing FFE-pyRevit User Analytics WebView2..."

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
        self.Closed += self.on_closed
        self.browser.CoreWebView2InitializationCompleted += self.on_core_webview2_initialized
        self.browser.NavigationCompleted += self.on_navigation_completed

    def on_loaded(self, sender, args):
        self.status_text.Visibility = Visibility.Visible
        self.status_text.Text = "Loading User Analytics from:\n{0}".format(self.index_uri.AbsoluteUri)
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
                self.status_text.Text = "User Analytics navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus)
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_analytics_payload()

    def on_closed(self, sender, args):
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
            self.status_text.Text = "Could not send data to the User Analytics web app:\n{0}".format(
                traceback.format_exc()
            )

    def send_analytics_payload(self, force=False):
        if self.has_sent_analytics_payload and not force:
            return

        self.has_sent_analytics_payload = True
        payload = load_user_analytics_payload()
        self.Title = "{0} - {1}".format(ANALYTICS_NAME, payload.get("username") or get_revit_username())
        self.execute_script(
            "window.ffeAnalytics && window.ffeAnalytics.loadData({0});".format(
                json_dumps(payload)
            )
        )

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
            message = json.loads(raw_message)
        except:
            return

        message_type = message.get("type")

        if message_type == "appReady":
            self.send_analytics_payload()
            return

        if message_type == "refreshData":
            self.send_analytics_payload(force=True)
            return

        if message_type == "closeWindow":
            self.Close()


# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The FFE-pyRevit User Analytics web app was not found:\n{0}".format(PATH_INDEX),
        title="FFE-pyRevit User Analytics",
        exitscript=True
    )

if not focus_existing_window():
    try:
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as webview_error:
        forms.alert(
            "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
            title="FFE-pyRevit User Analytics",
            exitscript=True
        )

    window = UserAnalyticsWindow(WebView2, CoreWebView2CreationProperties)
    WINDOW_REFS.append(window)
    window.Show()
