# -*- coding: utf-8 -*-
__title__   = "FFE-pyRevit \nv1.13.1"
__version__ = "Version = v1.13.1"
__persistentengine__ = True
__doc__ = """Version = v1.13.1
Date    = 05.15.2026
__________________________________________________________________
Description:
pyRevit About Form
__________________________________________________________________
How-To:
- Click the Button
__________________________________________________________________
Last update:
- [03.06.2025] - v1.0 RELEASE
- [03.25.2025] - v1.1 RELEASE
- [06.11.2025] - v1.2 RELEASE
- [07.25.2025] - v1.3 RELEASE
- [08.06.2025] - v1.4 RELEASE
- [09.14.2025] - v1.5 RELEASE
- [09.22.2025] - v1.5.1 Updated version numbering
- [10.08.2025] - v1.6.0 Added hooks
- [10.29.2025] - v1.6.2 Added Tutorials links
- [11.04.2025] - v1.7.0 Added Project Info Comparison Tool
- [11.07.2025] - v1.8.0 Added Revision Comparison Tool
- [11.14.2025] - RevisionCompare: v1.8.1 Updated to fully support "Per Sheet" revision numbering.
- [12.15.2025] - v1.9.0 Added Design Guide Model Splitter Tool
- [02.02.2026] - v1.10.0 Added Duct Network Summary to MEPTools
- [02.09.2026] - v1.10.1 Updated Duct Network Summary to work on all duct system types.
- [02.11.2026] - v1.11.0 Added Sheets_Counter tool.
- [02.17.2026] - v1.12.0 Added transferred-project-standards hook.
- [05.11.2026] - v1.13.0 Added Ductulator web app integration to MEPTools.
- [05.15.2026] - v1.13.1 CopyViewFilters v1.3: Fixed User cancelling view selection causing error
__________________________________________________________________
Author: Kyle Guggenheim from FFE Inc."""


# ____________________________________________________________________ IMPORTS
import json
import os
import traceback

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Diagnostics.Process import Start
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock

from pyrevit import forms


# ____________________________________________________________________ VARIABLES
revit_app = __revit__.Application

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

ABOUT_NAME = "FFE-pyRevit"
ABOUT_VERSION = "v1.13.1"
ABOUT_YEAR = "2026"

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

    user_data_folder = os.path.join(base_folder, "FFE-pyRevit", "AboutWebView2")
    if not os.path.exists(user_data_folder):
        os.makedirs(user_data_folder)
    return user_data_folder


def get_changelog_text():
    """Return changelog text for the About window."""
    changelog_md = os.path.join(PATH_SCRIPT, "CHANGELOG.md")
    if os.path.exists(changelog_md):
        try:
            with open(changelog_md, "r") as file_obj:
                return file_obj.read()
        except Exception as exc:
            forms.alert(
                "Could not read CHANGELOG.md:\n{0}".format(exc),
                title="FFE-pyRevit",
                warn_icon=True
            )

    if __doc__:
        return __doc__

    return "No changelog available."


def get_about_payload():
    return {
        "name": ABOUT_NAME,
        "version": ABOUT_VERSION,
        "year": ABOUT_YEAR,
        "changelog": get_changelog_text(),
    }


def is_external_url(url):
    value = safe_str(url).lower()
    return (
        value.startswith("http://") or
        value.startswith("https://") or
        value.startswith("mailto:")
    )


def open_external_url(url):
    if not is_external_url(url):
        return

    try:
        Start(url)
    except Exception as exc:
        forms.alert(
            "Could not open link:\n{0}\n\n{1}".format(url, exc),
            title="FFE-pyRevit",
            warn_icon=True
        )


def focus_existing_window():
    for window in list(WINDOW_REFS):
        try:
            if window.IsVisible:
                window.Activate()
                return True
        except:
            try:
                WINDOW_REFS.remove(window)
            except:
                pass
    return False


# ____________________________________________________________________ WEBVIEW WINDOW
class AboutWindow(Window):
    def __init__(self, webview_type, creation_properties_type, about_payload):
        Window.__init__(self)

        self.about_payload = about_payload
        self.has_sent_about_payload = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = "{0} {1}".format(ABOUT_NAME, ABOUT_VERSION)
        self.Width = 760
        self.Height = 820
        self.MinWidth = 620
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing FFE-pyRevit About WebView2..."

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
        self.status_text.Text = "Loading FFE-pyRevit About from:\n{0}".format(self.index_uri.AbsoluteUri)
        try:
            self.browser.EnsureCoreWebView2Async()
        except Exception as exc:
            self.status_text.Text = "Could not initialize WebView2:\n{0}".format(exc)

    def on_core_webview2_initialized(self, sender, args):
        try:
            if args.IsSuccess:
                self.browser.CoreWebView2.WebMessageReceived += self.on_web_message_received
                try:
                    self.browser.CoreWebView2.NavigationStarting += self.on_navigation_starting
                    self.browser.CoreWebView2.NewWindowRequested += self.on_new_window_requested
                except:
                    pass
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
                self.status_text.Text = "FFE-pyRevit About navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus)
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_about_payload()

    def on_navigation_starting(self, sender, args):
        try:
            url = safe_str(args.Uri)
            if is_external_url(url):
                args.Cancel = True
                open_external_url(url)
        except:
            pass

    def on_new_window_requested(self, sender, args):
        try:
            url = safe_str(args.Uri)
            if is_external_url(url):
                args.Handled = True
                open_external_url(url)
        except:
            pass

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
            self.status_text.Text = "Could not send data to the About web app:\n{0}".format(traceback.format_exc())

    def send_about_payload(self):
        if self.has_sent_about_payload:
            return

        self.has_sent_about_payload = True
        self.execute_script(
            "window.ffeAbout && window.ffeAbout.loadData({0});".format(
                json_dumps(self.about_payload)
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
            self.send_about_payload()
            return

        if message_type == "openExternal":
            open_external_url(message.get("url"))
            return

        if message_type == "closeWindow":
            self.Close()


# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The FFE-pyRevit About web app was not found:\n{0}".format(PATH_INDEX),
        title="FFE-pyRevit",
        exitscript=True
    )

if not focus_existing_window():
    try:
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as webview_error:
        forms.alert(
            "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
            title="FFE-pyRevit",
            exitscript=True
        )

    window = AboutWindow(
        WebView2,
        CoreWebView2CreationProperties,
        get_about_payload()
    )
    WINDOW_REFS.append(window)
    window.Show()
