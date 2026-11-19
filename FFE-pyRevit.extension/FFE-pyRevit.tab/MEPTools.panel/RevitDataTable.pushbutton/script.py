# -*- coding: utf-8 -*-
"""Purpose: browse host-model Spaces in a modeless, offline WebView2 data grid.
Behaviors: phase/level filtering, explicit refresh, selection and show in Revit.
Revit API: 2025/2026, persistent IronPython engine, queued ExternalEvent requests.
Design: browser callbacks use snapshots only; all document reads and navigation
run in API callbacks. This command opens no transactions and writes no model data.
"""
__title__ = "Revit Data\nTable"
__version__ = "0.1.1"
__min_revit_ver__ = 2025
__persistentengine__ = True

import collections
import json
import os
import sys
import uuid

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Collections.Generic import List
from System.Windows import (Window, WindowState, WindowStartupLocation,
                            Thickness, Visibility, TextWrapping)
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Interop import WindowInteropHelper
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from pyrevit import DB, forms, script

# Keep the command's helpers alongside its assets instead of requiring the
# extension-wide library directory to contain table-specific modules.
ScriptDirectory = os.path.dirname(os.path.abspath(__file__))
LibraryDirectory = os.path.join(ScriptDirectory, "lib")
if LibraryDirectory not in sys.path:
    sys.path.insert(0, LibraryDirectory)

from ffe_webview2 import load_webview2_types, format_webview2_error
from revit_data_table import BuildSnapshot, ValidateRequest, NavigateSpaces

try:
    _DATA_TABLE_WINDOWS
except NameError:
    _DATA_TABLE_WINDOWS = []

Logger = script.get_logger()
SupportDirectory = os.path.join(ScriptDirectory, "support")


class DataTableEventHandler(IExternalEventHandler):
    """Retain every accepted UI request until Revit grants a valid API context."""

    def __init__(self):
        self.Window = None
        self.Requests = collections.deque()

    def GetName(self):
        """Describe this bridge in Revit's event diagnostics."""
        return "FFE Revit Data Table"

    def Execute(self, uiApplication):
        """Recheck document identity at execution time, including queued requests."""
        while self.Requests and self.Window is not None:
            request = self.Requests.popleft()
            window = self.Window
            try:
                uiDocument = uiApplication.ActiveUIDocument
                if (not window.Document.IsValidObject or uiDocument is None or
                        not window.Document.Equals(uiDocument.Document)):
                    raise ValueError("Activate this table's model in Revit, then retry.")
                if request["action"] == "refresh":
                    snapshot = BuildSnapshot(window.Document, window.DocumentKey, DB)
                    window.Snapshot = snapshot
                    window.Outdated = False
                    window.SendSnapshot()
                    window.SendContext()
                    window.SendResult("warning" if snapshot["skippedCount"] else "success",
                                      "Refreshed {0} Space(s). Skipped {1}.".format(
                        len(snapshot["rows"]), snapshot["skippedCount"]))
                else:
                    result = NavigateSpaces(uiDocument, window.Snapshot, request, DB,
                                            lambda ids: List[DB.ElementId](ids))
                    window.SendResult(result["kind"], result["message"])
            except Exception as error:
                Logger.warning("Data table request failed: {0}".format(error))
                window.SendResult("error", str(error))


class DataTableWindow(Window):
    """Own a document-bound WebView and release all modeless resources on close."""

    def __init__(self, uiApplication, document, snapshot, handler, externalEvent,
                 webViewType, creationType):
        Window.__init__(self)
        self.Document = document
        self.DocumentKey = snapshot["documentKey"]
        self.Snapshot = snapshot
        self.Handler = handler
        self.ExternalEvent = externalEvent
        self.Application = uiApplication
        self.DocumentActive = True
        self.Outdated = False
        self.IsClosed = False
        self.Ready = False
        self.Subscribed = False
        self.Core = None
        self.IndexUri = Uri(os.path.abspath(os.path.join(SupportDirectory, "index.html"))).AbsoluteUri
        self.Title = "Revit Data Table - " + snapshot["documentTitle"]
        self.Width = 1160
        self.Height = 740
        self.MinWidth = 760
        self.MinHeight = 480
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        WindowInteropHelper(self).Owner = uiApplication.MainWindowHandle

        self.StatusText = TextBlock()
        self.StatusText.Text = "Loading Revit Data Table..."
        self.StatusText.Margin = Thickness(24)
        self.StatusText.TextWrapping = TextWrapping.Wrap
        self.Browser = webViewType()
        try:
            properties = creationType()
            userData = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                                    "FFE-pyRevit", "RevitDataTable", "WebView2")
            if not os.path.isdir(userData):
                os.makedirs(userData)
            properties.UserDataFolder = userData
            self.Browser.CreationProperties = properties
        except Exception:
            self.Browser.Dispose()
            raise
        content = Grid()
        content.Children.Add(self.Browser)
        content.Children.Add(self.StatusText)
        self.Content = content

        self.Loaded += self.OnLoaded
        self.Closed += self.HandleClosed
        self.Browser.CoreWebView2InitializationCompleted += self.OnCoreInitialized
        handler.Window = self

    def Subscribe(self):
        """Subscribe in the launching command's valid API context."""
        # Mark first so a partial subscription can still be undone during cleanup.
        self.Subscribed = True
        self.Application.Application.DocumentChanged += self.OnDocumentChanged
        self.Application.Application.DocumentClosed += self.OnDocumentClosed
        self.Application.ViewActivated += self.OnViewActivated

    def OnLoaded(self, sender, args):
        """Initialize asynchronously after WPF has created its native window."""
        try:
            self.Browser.EnsureCoreWebView2Async()
        except Exception as error:
            self.ShowFailure(format_webview2_error(error))

    def OnCoreInitialized(self, sender, args):
        """Attach the JSON bridge only after SDK/runtime initialization succeeds."""
        if self.IsClosed:
            return
        try:
            if not args.IsSuccess:
                raise Exception("WebView2 initialization failed: {0}".format(args.InitializationException))
            self.Core = self.Browser.CoreWebView2
            self.Core.WebMessageReceived += self.OnWebMessage
            self.Core.NavigationStarting += self.OnNavigationStarting
            self.Core.NavigationCompleted += self.OnNavigationCompleted
            self.Core.NewWindowRequested += self.OnNewWindowRequested
            self.Core.Navigate(self.IndexUri)
        except Exception as error:
            self.ShowFailure(format_webview2_error(error))

    def OnNavigationStarting(self, sender, args):
        """Keep the privileged bridge bound to the packaged application page."""
        if args.Uri != self.IndexUri:
            args.Cancel = True

    def OnNewWindowRequested(self, sender, args):
        """This offline table does not open browser popups."""
        args.Handled = True

    def OnNavigationCompleted(self, sender, args):
        """Report local asset/navigation failures in the native host."""
        if not args.IsSuccess:
            self.ShowFailure("Could not load the data table: {0}".format(args.WebErrorStatus))

    def OnWebMessage(self, sender, args):
        """Parse snapshot-only requests; never access the document from this callback."""
        if self.IsClosed or args.Source != self.IndexUri:
            return
        try:
            message = json.loads(args.WebMessageAsJson)
            request = ValidateRequest(message, self.DocumentKey, self.Snapshot)
            if request["action"] == "ready":
                self.Ready = True
                self.StatusText.Visibility = Visibility.Collapsed
                self.SendSnapshot()
                self.SendContext()
                return
            if not self.DocumentActive:
                raise ValueError("Activate this table's model in Revit, then retry.")
            if len(self.Handler.Requests) >= 8:
                raise ValueError("Revit is busy. Wait for the current request, then retry.")
            self.Handler.Requests.append(request)
            try:
                result = self.ExternalEvent.Raise()
                if str(result) not in ("Accepted", "Pending"):
                    raise ValueError("Revit could not accept the request ({0}). Retry when idle.".format(result))
            except Exception:
                self.Handler.Requests.pop()
                raise
        except Exception as error:
            self.SendResult("error", str(error))

    def Send(self, payload):
        """Post structured JSON instead of interpolating model text into scripts."""
        if self.Ready and not self.IsClosed and self.Core is not None:
            try:
                self.Core.PostWebMessageAsJson(json.dumps(payload, ensure_ascii=True, allow_nan=False))
            except Exception as error:
                self.ShowFailure("Could not send data to the table: {0}".format(error))

    def SendSnapshot(self):
        """Send a complete snapshot; the frontend preserves its own table state."""
        self.Send({"type": "snapshot", "snapshot": self.Snapshot})

    def SendContext(self):
        """Expose lifecycle flags without reading model data in the browser."""
        self.Send({"type": "context", "active": self.DocumentActive, "outdated": self.Outdated})

    def SendResult(self, kind, message):
        """Complete a request with a concise operation report."""
        self.Send({"type": "result", "kind": kind, "message": message})

    def ShowFailure(self, message):
        """Leave initialization/transport errors readable even if JavaScript fails."""
        self.StatusText.Text = str(message)
        self.StatusText.Visibility = Visibility.Visible

    def OnDocumentChanged(self, sender, args):
        """Invalidate after edits, undo, or redo without replacing the displayed rows."""
        if not self.IsClosed and args.GetDocument().Equals(self.Document):
            self.Outdated = True
            self.SendContext()

    def OnViewActivated(self, sender, args):
        """Document actions are enabled only for the bound model."""
        if self.IsClosed:
            return
        uiDocument = self.Application.ActiveUIDocument
        self.DocumentActive = (uiDocument is not None and self.Document.IsValidObject and
                               self.Document.Equals(uiDocument.Document))
        self.SendContext()

    def OnDocumentClosed(self, sender, args):
        """A closed document must never leave a navigable table behind."""
        if not self.IsClosed and not self.Document.IsValidObject:
            self.Close()

    def HandleClosed(self, sender, args):
        """Detach delegates before disposing their browser/event targets."""
        if self.IsClosed:
            return
        self.IsClosed = True
        self.Ready = False
        self.Handler.Window = None
        self.Handler.Requests.clear()
        if self.Subscribed:
            for detach in (self.DetachDocumentChanged, self.DetachDocumentClosed, self.DetachViewActivated):
                try:
                    detach()
                except Exception as error:
                    Logger.warning("Data table event cleanup: {0}".format(error))
            self.Subscribed = False
        if self.Core is not None:
            try:
                self.Core.WebMessageReceived -= self.OnWebMessage
                self.Core.NavigationStarting -= self.OnNavigationStarting
                self.Core.NavigationCompleted -= self.OnNavigationCompleted
                self.Core.NewWindowRequested -= self.OnNewWindowRequested
            except Exception as error:
                Logger.warning("Data table browser cleanup: {0}".format(error))
        self.Browser.CoreWebView2InitializationCompleted -= self.OnCoreInitialized
        self.Loaded -= self.OnLoaded
        self.Closed -= self.HandleClosed
        for resource in (self.Browser, self.ExternalEvent):
            try:
                resource.Dispose()
            except Exception as error:
                Logger.warning("Data table disposal: {0}".format(error))
        if self in _DATA_TABLE_WINDOWS:
            _DATA_TABLE_WINDOWS.remove(self)

    def DetachDocumentChanged(self):
        """Remove the exact bound lifecycle handler."""
        self.Application.Application.DocumentChanged -= self.OnDocumentChanged

    def DetachDocumentClosed(self):
        """Remove the exact bound lifecycle handler."""
        self.Application.Application.DocumentClosed -= self.OnDocumentClosed

    def DetachViewActivated(self):
        """Remove the exact bound lifecycle handler."""
        self.Application.ViewActivated -= self.OnViewActivated


def Main():
    """Focus this model's existing table, or collect its initial snapshot and open."""
    uiApplication = __revit__
    uiDocument = uiApplication.ActiveUIDocument
    if uiDocument is None or uiDocument.Document.IsFamilyDocument:
        forms.alert("Open a Revit project to browse Spaces.", title="Revit Data Table")
        return
    document = uiDocument.Document
    for window in list(_DATA_TABLE_WINDOWS):
        if not window.IsClosed and window.Document.Equals(document):
            if window.WindowState == WindowState.Minimized:
                window.WindowState = WindowState.Normal
            window.Activate()
            return
    externalEvent = None
    window = None
    try:
        snapshot = BuildSnapshot(document, str(uuid.uuid4()), DB)
        webViewType, creationType = load_webview2_types()
        handler = DataTableEventHandler()
        externalEvent = ExternalEvent.Create(handler)
        window = DataTableWindow(uiApplication, document, snapshot, handler, externalEvent,
                                 webViewType, creationType)
        window.Subscribe()
        _DATA_TABLE_WINDOWS.append(window)
        window.Show()
    except Exception as error:
        if window is not None:
            window.HandleClosed(None, None)
            window.Close()
        elif externalEvent is not None:
            externalEvent.Dispose()
        Logger.exception("Could not open Revit Data Table")
        forms.alert("Could not open Revit Data Table.\n\n{0}".format(error), title="Revit Data Table")


if __name__ == "__main__":
    Main()
