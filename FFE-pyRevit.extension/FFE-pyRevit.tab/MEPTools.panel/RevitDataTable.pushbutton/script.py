# -*- coding: utf-8 -*-
"""Purpose: browse host-model Spaces in a modeless WebView2 data grid.
Behaviors: shared saved views, filtering, navigation, and reviewed repository commits.
Revit API: 2025/2026, persistent IronPython engine, queued ExternalEvent requests.
Design: browser callbacks use snapshots only; all document reads and navigation
run in API callbacks. Saved-table operations transact metadata; Space values stay read-only.
"""
__title__ = "HVAC Data"
__version__ = "0.12.0"
__min_revit_ver__ = 2025
__persistentengine__ = True

import collections
import json
import os
import sys
import threading
import uuid

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri, Action
from System.Collections.Generic import List
from System.Windows import (Window, WindowState, WindowStartupLocation,
                            Thickness, Visibility, TextWrapping,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult)
from System.Windows.Controls import Grid, TextBlock
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from pyrevit import DB, forms, script

# Keep the command's helpers alongside its assets instead of requiring the
# extension-wide library directory to contain table-specific modules.
ScriptDirectory = os.path.dirname(os.path.abspath(__file__))
LibraryDirectory = os.path.join(ScriptDirectory, "lib")
if LibraryDirectory not in sys.path:
    sys.path.insert(0, LibraryDirectory)

from ffe_webview2 import load_webview2_types, format_webview2_error
from revit_data_table import BuildSnapshot, ValidateRequest, NavigateSpaces, PublishActions
from saved_tables import RevitTableStore, CaptureLayout, ResolveLayout, TableMetadata, ValidateLayout
from ffe_supabase import Client as CloudClient
from cloud_publish import CloudModelStore, LoadCatalog, PrepareCollection, ReviewCollection, CreateBatch, CommitReview
from forma_project import ReadFormaContext, ModelIdForDocument, CheckFormaProject, CreateFormaProject

try:
    _DATA_TABLE_WINDOWS
except NameError:
    _DATA_TABLE_WINDOWS = []

Logger = script.get_logger()
SupportDirectory = os.path.join(ScriptDirectory, "support")


def CreateCloudClient(application):
    """Give repository writes a persistent pyRevit installation identity and attribution."""
    config = script.get_config("HVACDataRepository")
    clientId = config.get_option("client_id", None)
    if not clientId:
        clientId = str(uuid.uuid4())
        config.client_id = clientId
        script.save_config()
    name = u"{0}".format(application.Username or os.environ.get("USERNAME") or "Revit").strip() or "Revit"
    try:
        userId = str(application.LoginUserId or "")[:200]
    except Exception:
        userId = ""
    return CloudClient(client_id=clientId, actor_name=name[:200], user_id=userId)


class DataTableEventHandler(IExternalEventHandler):
    """Retain every accepted UI request until Revit grants a valid API context."""

    def __init__(self):
        self.Window = None
        self.Requests = collections.deque()

    def GetName(self):
        """Describe this bridge in Revit's event diagnostics."""
        return "HVAC Data"

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
                    snapshot = BuildSnapshot(window.Document, window.DocumentKey, DB,
                                             request.get("parameterIds", window.Snapshot.get("selectedParameterIds", [])))
                    window.SetSnapshot(snapshot)
                    window.Outdated = False
                    window.SendSnapshot()
                    window.SendContext()
                    window.SendResult("warning" if snapshot["skippedCount"] else "success",
                                      "Refreshed {0} Space(s). Skipped {1}.".format(
                        len(snapshot["rows"]), snapshot["skippedCount"]))
                elif request["action"] in PublishActions:
                    window.HandlePublishRequest(request)
                elif request.get("requestId"):
                    window.HandleTableRequest(request)
                else:
                    result = NavigateSpaces(uiDocument, window.Snapshot, request, DB,
                                            lambda ids: List[DB.ElementId](ids))
                    window.SendResult(result["kind"], result["message"])
            except Exception as error:
                Logger.warning("Data table request failed: {0}".format(error))
                if request.get("requestId"):
                    window.SendTableFailure(request, error)
                else:
                    window.SendResult("error", str(error))


class DataTableWindow(Window):
    """Own a document-bound WebView and release all modeless resources on close."""

    def __init__(self, uiApplication, document, snapshot, handler, externalEvent,
                 webViewType, creationType, tableStore=None, formaContext=None,
                 startupClient=None, linkedModelId=None):
        Window.__init__(self)
        self.Document = document
        self.DocumentKey = snapshot["documentKey"]
        self.TableStore = tableStore
        self.SetSnapshot(snapshot)
        self.Handler = handler
        self.ExternalEvent = externalEvent
        self.Application = uiApplication
        self.DocumentActive = True
        self.Outdated = False
        self.IsClosed = False
        self.Ready = False
        self.Subscribed = False
        self.Core = None
        self.ClosePermit = False
        self.WritingTables = False
        self.CloudBusy = False
        self.FormaContext = formaContext
        self.FormaRepositoryProject = None
        self.StartupClient = startupClient
        self.StartupLinkedModelId = linkedModelId
        self.StartupAttempted = False
        self.PendingFormaResponse = None
        self.PublishPlan = None
        self.PublishReview = None
        self.PublishBatch = None
        self.PublishClient = None
        self.DocumentRevision = 0
        self.CloudStore = CloudModelStore(document, DB)
        self.IndexUri = Uri(os.path.abspath(os.path.join(SupportDirectory, "index.html"))).AbsoluteUri
        self.Title = "HVAC Data - " + snapshot["documentTitle"]
        self.Width = 1160
        self.Height = 740
        self.MinWidth = 760
        self.MinHeight = 480
        # An owned window always stays above Revit. Keep this modeless window
        # independent so normal focus changes can put Revit in front of it.
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = False
        self.ShowInTaskbar = True

        self.StatusText = TextBlock()
        self.StatusText.Text = "Loading HVAC Data..."
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
        self.Closing += self.HandleClosing
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
        """Keep the privileged table page from opening browser popups."""
        args.Handled = True

    def OnNavigationCompleted(self, sender, args):
        """Report local asset/navigation failures in the native host."""
        if not args.IsSuccess:
            self.ShowFailure("Could not load the data table: {0}".format(args.WebErrorStatus))

    def OnWebMessage(self, sender, args):
        """Parse snapshot-only requests; never access the document from this callback."""
        if self.IsClosed or args.Source != self.IndexUri:
            return
        message = None
        try:
            message = json.loads(args.WebMessageAsJson)
            request = ValidateRequest(message, self.DocumentKey, self.Snapshot)
            if request["action"] == "ready":
                self.Ready = True
                self.StatusText.Visibility = Visibility.Collapsed
                self.SendSnapshot()
                self.SendContext()
                self.BeginProjectLookup()
                return
            if request["action"] == "closeWindow":
                self.ClosePermit = True
                self.Close()
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
            if isinstance(message, dict) and message.get("requestId"):
                self.SendTableFailure(message, error)
            else:
                self.SendResult("error", str(error))

    def SetSnapshot(self, snapshot):
        """Attach the latest model catalog without creating storage on read."""
        store = getattr(self, "TableStore", None)
        if store is not None:
            try:
                records, issues = store.List()
                snapshot["savedTables"] = sorted([TableMetadata(record) for element, record in records],
                                                 key=lambda item: (item["name"].lower(), item["id"]))
                snapshot["tableIssues"] = issues
                snapshot["canSaveTables"] = not self.Document.IsReadOnly
            except Exception as error:
                snapshot["savedTables"] = []
                snapshot["tableIssues"] = ["Could not read saved tables: {0}".format(error)]
                snapshot["canSaveTables"] = False
        self.Snapshot = snapshot

    def HandleTableRequest(self, request):
        """Read or transact table definitions only from the external-event handler."""
        action = request["action"]
        if action == "loadTable":
            record = self.TableStore.Read(request["tableId"])[1] if request["tableId"] else None
            selectedIds = self.TableStore.ResolveParameterIds(record["layout"]) if record else []
            snapshot = BuildSnapshot(self.Document, self.DocumentKey, DB, selectedIds)
            layout, warnings = ResolveLayout(record["layout"], snapshot) if record else (None, [])
            self.SetSnapshot(snapshot)
            self.Outdated = False
            self.SendContext()
            self.Send({"type": "tableResult", "action": action, "requestId": request["requestId"],
                       "success": True, "table": TableMetadata(record) if record else None,
                       "snapshot": snapshot, "layout": layout, "warnings": warnings})
            return
        layout = CaptureLayout(request["layout"], self.Snapshot) if action == "saveTable" else None
        self.WritingTables = True
        try:
            record = self.TableStore.Write(action, request["tableId"], request["revision"],
                                           request["name"], layout)
        finally:
            self.WritingTables = False
        self.SetSnapshot(self.Snapshot)
        self.Send({"type": "tableResult", "action": action, "requestId": request["requestId"],
                   "success": True, "table": TableMetadata(record) if record else None,
                   "catalog": {key: self.Snapshot.get(key) for key in ("savedTables", "tableIssues", "canSaveTables")},
                   "layout": ValidateLayout(request["layout"]) if action == "saveTable" else None})

    def SendTableFailure(self, request, error):
        """Retain the browser's working view when saving/loading cannot complete."""
        publish = request.get("action") in PublishActions
        self.Send({"type": "publishResult" if publish else "tableResult", "action": request.get("action"),
                   "requestId": request.get("requestId"), "success": False, "message": str(error)})

    def BeginProjectLookup(self):
        """Use the launch-time plain-data context; this browser callback never reads Revit."""
        context = getattr(self, "FormaContext", None)
        if context is None or getattr(self, "StartupAttempted", False):
            return
        self.StartupAttempted = True
        if not context.get("available"):
            self.SendResult("warning", context["reason"])
            return
        client = self.StartupClient
        if client is None:
            self.SendResult("warning", "Could not initialize the project connection. Local Space browsing is available.")
            return
        linkedId = self.StartupLinkedModelId
        request = {"action": "formaLookup", "requestId": str(uuid.uuid4())}
        self.SendResult("info", "Checking the Autodesk Forma project connection...")
        self.StartCloudWork(request, lambda: CheckFormaProject(client, context, linkedId), self.OnFormaLookup)

    def OnFormaLookup(self, response):
        """Prompt on the UI dispatcher only after a successful missing-project lookup."""
        if self.IsClosed:
            return
        if not self.DocumentActive:
            self.PendingFormaResponse = response
            return
        self.PendingFormaResponse = None
        if not response["success"]:
            message = response["message"]
            if response.get("code") == "PGRST202":
                message = "Apply FFE Mechanical's Forma project setup migration (20261120000400), then close and reopen HVAC Data."
            self.SendResult("warning", "Could not check the Forma project. {0}".format(message))
            return
        plan = response["data"]
        if plan["project"]:
            self.FormaRepositoryProject = plan["project"]
            self.SendResult("success", u"Connected to project: {0}".format(plan["project"]["name"]))
            return
        context = plan["forma"]
        name = context.get("formaProjectName")
        title = "Connect project" if plan.get("existingProject") else "Start project"
        if plan.get("existingProject"):
            question = u'Connect the existing shared project "{0}" to this Autodesk Forma project?'.format(plan["existingProject"]["name"])
        else:
            question = u'Start a shared project for "{0}"?'.format(name or "this Autodesk Forma project")
        text = question + u"\n\nAutodesk Forma project: {0}\nProject ID: {1}\n\nSpace values will be published only after you review and commit them.".format(
            name or "Name unavailable from Revit", context["formaProjectId"])
        if MessageBox.Show(self, text, title, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes:
            self.SendResult("info", "Project setup skipped. Local Space browsing is available.")
            return
        if not name:
            name = forms.ask_for_string(prompt="Revit supplied the Forma project ID but not its name. Enter the project name shown in Autodesk Forma:",
                                        title="Autodesk Forma project name")
            if not name:
                self.SendResult("info", "Project setup canceled.")
                return
        requestId = str(uuid.uuid4())
        client = self.StartupClient
        request = {"action": "formaCreate", "requestId": requestId}

        def create():
            return CreateFormaProject(client, plan, name, requestId)

        def completed(result):
            if result["success"]:
                self.FormaRepositoryProject = result["data"]["project"]
                self.SendResult("success", u"Project ready: {0}. Use Publish to review Space changes.".format(self.FormaRepositoryProject["name"]))
            elif result.get("retry"):
                retry = MessageBox.Show(self, "The project request could not be confirmed. Retry the same request?\n\n" + result["message"],
                                        "Project connection", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No)
                if retry == MessageBoxResult.Yes:
                    self.StartCloudWork(request, create, completed)
                else:
                    self.SendResult("warning", "Project request unconfirmed. Close and reopen HVAC Data to check the existing project.")
            else:
                self.SendResult("warning", "Could not start the project: " + result["message"])

        self.StartCloudWork(request, create, completed)

    def HandlePublishRequest(self, request):
        """Collect/recheck model data in Revit's API context before starting network work."""
        if self.CloudBusy:
            raise ValueError("A publish request is already running. Wait for it to finish.")
        action = request["action"]
        if action == "publishCatalog":
            self.PublishPlan = None
            self.PublishReview = None
            self.PublishBatch = None
            self.PublishClient = None
            linkedId = self.CloudStore.Read()
            context = getattr(self, "FormaContext", None)
            modelId = ModelIdForDocument(linkedId, context)
            preferred = (getattr(self, "FormaRepositoryProject", None) or {}).get("projectId")
            client = CloudClient()
            self.StartCloudWork(request, lambda: LoadCatalog(client, modelId, linkedId is not None, preferred, context))
            return
        if action == "publishCommit":
            review = self.PublishReview
            if review is None or review["collectionId"] != request["reviewId"]:
                raise ValueError("Collect Space changes before committing.")
            batch = self.PublishBatch
            if batch is None:
                if review["documentRevision"] != self.DocumentRevision:
                    raise ValueError("The Revit model changed. Collect Space changes again.")
                batch = CreateBatch(review, request["selection"], request["message"])
            elif batch["selection"] != request["selection"] or batch["message"] != request["message"]:
                raise ValueError("Retry the same commit or start a new review.")
            client = self.PublishClient
            self.PublishBatch = batch
            self.StartCloudWork(request, lambda: CommitReview(client, review, batch))
            return
        plan = self.PublishPlan
        project = next((item for item in plan["projects"] if item["id"] == request["projectId"]), None) if plan else None
        if project is None:
            raise ValueError("Reload projects before collecting Space changes.")
        if plan["model"] is not None and plan["model"]["projectId"] != request["projectId"]:
            raise ValueError("This model is already linked to another project. Reload projects.")
        if plan.get("projectId") and plan["projectId"] != request["projectId"]:
            raise ValueError("Use this model's Autodesk Forma project.")
        self.PublishReview = None
        self.PublishBatch = None
        capabilities = {}
        snapshot = BuildSnapshot(self.Document, self.DocumentKey, DB,
                                 self.Snapshot.get("selectedParameterIds", []), capabilities)
        collection = PrepareCollection(snapshot, capabilities)
        client = CreateCloudClient(self.Application.Application)
        self.PublishClient = client
        self.WritingTables = True
        try:
            self.CloudStore.Ensure(plan["modelId"])
        finally:
            self.WritingTables = False
        self.SetSnapshot(snapshot)
        self.Outdated = False
        self.SendSnapshot()
        self.SendContext()
        labels = {row["uniqueId"]: u" ".join(part for part in (row["number"], row["name"]) if part)
                  or row["uniqueId"] for row in snapshot["rows"]}
        name = snapshot["documentTitle"].strip()[:200] or "Revit model"
        revision = self.DocumentRevision

        def review():
            result = ReviewCollection(client, plan, project, collection, name, labels)
            result["documentRevision"] = revision
            return result

        self.StartCloudWork(request, review)

    def StartCloudWork(self, request, work, onComplete=None):
        """Run JSON-only HTTPS calls off-thread and deliver results on the WPF dispatcher."""
        self.CloudBusy = True
        if onComplete is not None:
            self.SendContext()

        def run():
            response = {"type": "publishResult", "action": request["action"], "requestId": request["requestId"]}
            try:
                result = work()
                if request["action"] == "publishCommit" and not result["ok"]:
                    response.update(success=False,
                                    message="The repository changed during review. No changes from this commit were published.")
                else:
                    response.update(success=True, data=result)
            except Exception as error:
                message = str(error)
                status = getattr(error, "status", None)
                code = getattr(error, "code", None)
                databaseTimeout = code == "57014" and "statement timeout" in message.lower()
                retry = request["action"] in ("publishCommit", "formaCreate") and (
                    code == "NETWORK_ERROR" or databaseTimeout
                    or (status is not None and (status >= 500 or 200 <= status < 300))
                    or (status is None and getattr(error, "request_id", None) is not None))
                if databaseTimeout and request["action"] == "publishCommit":
                    operation = getattr(error, "publish_operation", "publishing selected changes")
                    message = "The repository database timed out while {0}. No changes from this attempt were committed.".format(operation)
                    response["retryHint"] = "Apply the repository publish-performance migration before retrying this commit."
                response.update(success=False, message=message, retry=retry, code=code)

            def complete():
                if self.IsClosed:
                    return
                self.CloudBusy = False
                if onComplete is not None:
                    self.SendContext()
                    onComplete(response)
                    return
                if request["action"] == "publishCatalog" and response["success"]:
                    self.PublishPlan = response["data"]
                elif request["action"] == "publishReview" and response["success"]:
                    self.PublishReview = response["data"]
                elif request["action"] == "publishCommit" and not response.get("retry"):
                    self.PublishReview = None
                    self.PublishBatch = None
                    self.PublishClient = None
                self.Send(response)

            try:
                self.Dispatcher.BeginInvoke(Action(complete))
            except Exception:
                # Closing Revit may dispose the dispatcher before a bounded request ends.
                pass

        worker = threading.Thread(target=run)
        worker.daemon = True
        try:
            worker.start()
        except Exception:
            self.CloudBusy = False
            raise

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
        self.Send({"type": "context", "active": self.DocumentActive, "outdated": self.Outdated,
                   "repositoryBusy": getattr(self, "CloudBusy", False)})

    def SendResult(self, kind, message):
        """Complete a request with a concise operation report."""
        self.Send({"type": "result", "kind": kind, "message": message})

    def ShowFailure(self, message):
        """Leave initialization/transport errors readable even if JavaScript fails."""
        self.StatusText.Text = str(message)
        self.StatusText.Visibility = Visibility.Visible
        self.Ready = False

    def OnDocumentChanged(self, sender, args):
        """Invalidate after edits, undo, or redo without replacing the displayed rows."""
        if not self.IsClosed and not getattr(self, "WritingTables", False) and args.GetDocument().Equals(self.Document):
            self.DocumentRevision += 1
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

        pending = getattr(self, "PendingFormaResponse", None)
        if self.DocumentActive and pending is not None:
            self.OnFormaLookup(pending)

    def OnDocumentClosed(self, sender, args):
        """A closed document must never leave a navigable table behind."""
        if not self.IsClosed and not self.Document.IsValidObject:
            self.ClosePermit = True
            self.Close()

    def HandleClosing(self, sender, args):
        """Let the browser resolve Save/Discard/Cancel before the native window closes."""
        if self.IsClosed or self.ClosePermit or not self.Ready or not self.Document.IsValidObject:
            return
        args.Cancel = True
        self.Send({"type": "closeRequest"})

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
        self.Closing -= self.HandleClosing
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
        forms.alert("Open a Revit project to browse Spaces.", title="HVAC Data")
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
        formaContext = ReadFormaContext(document, DB)
        startupClient, linkedModelId = None, None
        if formaContext and formaContext.get("available"):
            try:
                startupClient = CreateCloudClient(uiApplication.Application)
                linkedModelId = CloudModelStore(document, DB).Read()
            except Exception as error:
                Logger.warning("Could not initialize Forma project connection: {0}".format(error))
                formaContext = {"available": False, "reason": "Could not read the saved project connection. Local Space browsing is available. " + str(error)}
        webViewType, creationType = load_webview2_types()
        handler = DataTableEventHandler()
        externalEvent = ExternalEvent.Create(handler)
        window = DataTableWindow(uiApplication, document, snapshot, handler, externalEvent,
                                 webViewType, creationType, RevitTableStore(document, DB),
                                 formaContext, startupClient, linkedModelId)
        window.Subscribe()
        _DATA_TABLE_WINDOWS.append(window)
        window.Show()
    except Exception as error:
        if window is not None:
            window.HandleClosed(None, None)
            window.Close()
        elif externalEvent is not None:
            externalEvent.Dispose()
        Logger.exception("Could not open HVAC Data")
        forms.alert("Could not open HVAC Data.\n\n{0}".format(error), title="HVAC Data")


if __name__ == "__main__":
    Main()
