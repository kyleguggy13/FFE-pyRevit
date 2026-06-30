# -*- coding: utf-8 -*-
__title__ = "Lineage Viewer"
__version__ = "Version = v0.2"
__persistentengine__ = True
__doc__ = """Version = v0.2
Date    = 06.30.2026
______________________________________________________________
Description:
-> Opens a WebView2 relationship viewer for linked model lineage.
-> Shows Revit links, CAD links/imports, images/PDFs, and point clouds.
______________________________________________________________
How-to:
-> Press Button
-> Review the linked item relationship graph
______________________________________________________________
Last update:
- [03.25.2026] - v0.1 BETA RELEASE
- [06.30.2026] - v0.2 Added WebView2 viewer
______________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS (SYSTEM)
import json
import os
import re
import time
import traceback
from collections import defaultdict

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Interop import WindowInteropHelper


# ____________________________________________________________________ IMPORTS (AUTODESK)
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ExternalFileUtils,
    FilteredElementCollector,
    ImageInstance,
    ImportInstance,
    ModelPathUtils,
    PointCloudType,
    RevitLinkInstance,
    RevitLinkType,
)


# ____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms, script


# ____________________________________________________________________ VARIABLES
revit_app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

LOGGER = script.get_logger()

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


def clean_filename(name):
    if not name:
        return "Untitled"
    return re.sub(r'[\\/*?:"<>|]+', "_", safe_str(name))


def basename_or_value(path_value):
    if not path_value:
        return ""
    try:
        base_name = os.path.basename(path_value)
        return base_name if base_name else safe_str(path_value)
    except:
        return safe_str(path_value)


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path


def get_local_app_dir():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return ensure_dir(os.path.join(base_folder, "FFE-pyRevit"))


def get_webview_user_data_folder():
    return ensure_dir(os.path.join(get_local_app_dir(), "LineageViewerWebView2"))


def get_revit_install_dir():
    try:
        app_path = revit_app.ApplicationPath
        if app_path and os.path.exists(app_path):
            return os.path.dirname(app_path)
    except:
        pass

    version = "2026"
    try:
        version = safe_str(revit_app.VersionNumber) or version
    except:
        pass

    return os.path.join(
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        "Autodesk",
        "Revit {0}".format(version)
    )


def load_webview2_types():
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


def try_get_external_path(owner_doc, element_id):
    try:
        ext_ref = ExternalFileUtils.GetExternalFileReference(owner_doc, element_id)
        if ext_ref:
            model_path = ext_ref.GetAbsolutePath()
            if model_path:
                return ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
    except:
        pass
    return ""


def safe_get_link_status(link_type):
    try:
        status = RevitLinkType.GetLinkedFileStatus(link_type)
        if status:
            return safe_str(status)
    except:
        pass
    return "Unknown"


def get_import_name(import_instance, owner_doc):
    try:
        param = import_instance.Parameter[BuiltInParameter.IMPORT_SYMBOL_NAME]
        if param:
            value = safe_str(param.AsString())
            if value:
                return value
    except:
        pass

    try:
        element_type = owner_doc.GetElement(import_instance.GetTypeId())
        if element_type:
            value = safe_str(element_type.Name)
            if value:
                return value
    except:
        pass

    try:
        value = safe_str(import_instance.Name)
        if value:
            return value
    except:
        pass

    return "CAD Item"


def get_image_type_path(image_type):
    if not image_type:
        return ""

    for attr_name in ["Path", "GetPath"]:
        try:
            value = getattr(image_type, attr_name)
            if callable(value):
                path_value = value()
            else:
                path_value = value
            if path_value:
                return safe_str(path_value)
        except:
            pass

    return ""


def get_pointcloud_path(pointcloud_type):
    for attr_name in ["Path", "GetPath"]:
        try:
            value = getattr(pointcloud_type, attr_name)
            if callable(value):
                path_value = value()
            else:
                path_value = value
            if path_value:
                return safe_str(path_value)
        except:
            pass
    return ""


def doc_key(current_doc):
    try:
        if current_doc.PathName:
            return current_doc.PathName
    except:
        pass

    try:
        return "DOCID:{0}".format(current_doc.GetHashCode())
    except:
        return "DOC:{0}".format(id(current_doc))


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


# ____________________________________________________________________ GRAPH DATA
class GraphNode(object):
    def __init__(self, node_id, label, sublabel, kind, depth, parent_id=None, status=""):
        self.id = node_id
        self.label = label
        self.sublabel = sublabel
        self.kind = kind
        self.depth = depth
        self.parent_id = parent_id
        self.status = status


class LineageGraph(object):
    def __init__(self):
        self.nodes = {}
        self.edges = []
        self.node_counter = 0
        self.visited_doc_keys = set()

    def next_id(self, prefix):
        self.node_counter += 1
        return "{0}_{1}".format(prefix, self.node_counter)

    def add_node(self, label, sublabel, kind, depth, parent_id=None, status=""):
        node_id = self.next_id(kind)
        self.nodes[node_id] = GraphNode(
            node_id,
            safe_str(label),
            safe_str(sublabel),
            safe_str(kind),
            int(depth),
            parent_id,
            safe_str(status),
        )
        if parent_id:
            self.edges.append({
                "from": parent_id,
                "to": node_id,
            })
        return node_id


# ____________________________________________________________________ COLLECTION
def collect_imports(owner_doc, parent_node_id, depth, graph):
    try:
        imports = (
            FilteredElementCollector(owner_doc)
            .OfClass(ImportInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except:
        imports = []

    for import_instance in imports:
        try:
            is_link = False
            try:
                is_link = import_instance.IsLinked
            except:
                pass

            kind = "cadlink" if is_link else "cadimport"
            path_value = try_get_external_path(owner_doc, import_instance.GetTypeId())

            graph.add_node(
                label=get_import_name(import_instance, owner_doc),
                sublabel=basename_or_value(path_value) if path_value else "No path available",
                kind=kind,
                depth=depth,
                parent_id=parent_node_id,
                status="Linked" if is_link else "Imported",
            )
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass


def collect_images_and_pdfs(owner_doc, parent_node_id, depth, graph):
    try:
        image_instances = (
            FilteredElementCollector(owner_doc)
            .OfClass(ImageInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except:
        image_instances = []

    for image_instance in image_instances:
        try:
            image_type = None
            try:
                image_type = owner_doc.GetElement(image_instance.GetTypeId())
            except:
                pass

            path_value = get_image_type_path(image_type)

            source_str = "Image"
            status_str = ""
            page_number = 1

            try:
                source_str = safe_str(image_type.Source) or source_str
            except:
                pass

            try:
                status_str = safe_str(image_type.Status)
            except:
                pass

            try:
                page_number = int(image_type.PageNumber)
            except:
                page_number = 1

            owner_view_name = "Unknown View"
            try:
                owner_view = owner_doc.GetElement(image_instance.OwnerViewId)
                if owner_view:
                    owner_view_name = safe_str(owner_view.Name) or owner_view_name
            except:
                pass

            extension = os.path.splitext(path_value)[1].lower() if path_value else ""
            is_pdf = extension == ".pdf" or page_number > 1

            if is_pdf:
                label = safe_str(image_instance.Name) or "PDF"
                sublabel = "{0} | Page {1} | View: {2}".format(
                    basename_or_value(path_value) if path_value else "PDF",
                    page_number,
                    owner_view_name,
                )
                kind = "pdf"
            else:
                label = safe_str(image_instance.Name) or "Image"
                sublabel = "{0} | View: {1}".format(
                    basename_or_value(path_value) if path_value else "Image",
                    owner_view_name,
                )
                kind = "image"

            graph.add_node(
                label=label,
                sublabel=sublabel,
                kind=kind,
                depth=depth,
                parent_id=parent_node_id,
                status=status_str if status_str else source_str,
            )
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass


def collect_pointclouds(owner_doc, parent_node_id, depth, graph):
    try:
        pointcloud_types = FilteredElementCollector(owner_doc).OfClass(PointCloudType).ToElements()
    except:
        pointcloud_types = []

    for pointcloud_type in pointcloud_types:
        try:
            path_value = get_pointcloud_path(pointcloud_type)
            graph.add_node(
                label=pointcloud_type.Name,
                sublabel=basename_or_value(path_value) if path_value else "Point cloud",
                kind="pointcloud",
                depth=depth,
                parent_id=parent_node_id,
            )
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass


def collect_revit_links(owner_doc, parent_node_id, depth, graph):
    try:
        link_instances = FilteredElementCollector(owner_doc).OfClass(RevitLinkInstance).ToElements()
    except:
        link_instances = []

    for link_instance in link_instances:
        try:
            link_type = owner_doc.GetElement(link_instance.GetTypeId())
            link_doc = link_instance.GetLinkDocument()

            instance_name = ""
            try:
                instance_name = safe_str(link_instance.Name)
            except:
                pass

            if not instance_name and link_type:
                instance_name = safe_str(link_type.Name)
            if not instance_name:
                instance_name = "Revit Link"

            path_value = ""
            if link_type:
                path_value = try_get_external_path(owner_doc, link_type.Id)

            status = safe_get_link_status(link_type) if link_type else "Unknown"
            node_id = graph.add_node(
                label=instance_name,
                sublabel=basename_or_value(path_value) if path_value else "No path available",
                kind="revitlink",
                depth=depth,
                parent_id=parent_node_id,
                status=status,
            )

            if link_doc:
                key = doc_key(link_doc)
                if key not in graph.visited_doc_keys:
                    graph.visited_doc_keys.add(key)
                    collect_doc_contents(link_doc, node_id, depth + 1, graph)
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass


def collect_doc_contents(owner_doc, parent_node_id, depth, graph):
    collect_revit_links(owner_doc, parent_node_id, depth, graph)
    collect_imports(owner_doc, parent_node_id, depth, graph)
    collect_images_and_pdfs(owner_doc, parent_node_id, depth, graph)
    collect_pointclouds(owner_doc, parent_node_id, depth, graph)


def build_lineage_payload(active_doc):
    graph = LineageGraph()
    host_node_id = graph.add_node(
        label=active_doc.Title,
        sublabel=basename_or_value(active_doc.PathName) if active_doc.PathName else "Unsaved model",
        kind="host",
        depth=0,
        parent_id=None,
        status="Active Model",
    )

    graph.visited_doc_keys.add(doc_key(active_doc))
    collect_doc_contents(active_doc, host_node_id, 1, graph)

    counts = defaultdict(int)
    node_payloads = []
    for node in graph.nodes.values():
        counts[node.kind] += 1
        node_payloads.append({
            "id": node.id,
            "label": node.label,
            "sublabel": node.sublabel,
            "kind": node.kind,
            "depth": node.depth,
            "parentId": node.parent_id,
            "status": node.status,
        })

    return {
        "model": {
            "title": safe_str(active_doc.Title),
            "path": safe_str(active_doc.PathName) if active_doc.PathName else "Unsaved model",
            "generated": time.strftime("%Y-%m-%d %I:%M:%S %p"),
            "note": "Nested contents are only shown for loaded Revit links.",
        },
        "counts": dict(counts),
        "nodes": node_payloads,
        "edges": list(graph.edges),
    }


# ____________________________________________________________________ WEBVIEW WINDOW
class LineageViewerWindow(Window):
    def __init__(self, webview_type, creation_properties_type, lineage_payload):
        Window.__init__(self)

        try:
            WindowInteropHelper(self).Owner = __revit__.MainWindowHandle
            self.ShowInTaskbar = True
        except:
            pass

        self.lineage_payload = lineage_payload
        self.has_sent_lineage_payload = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = "Lineage Viewer - {0}".format(
            clean_filename(lineage_payload.get("model", {}).get("title", "Revit Model"))
        )
        self.Width = 1280
        self.Height = 840
        self.MinWidth = 900
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing Lineage Viewer WebView2..."

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
        self.status_text.Text = "Loading Lineage Viewer from:\n{0}".format(self.index_uri.AbsoluteUri)
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
                self.status_text.Text = "Lineage Viewer navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus),
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_lineage_payload()

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
            self.status_text.Text = "Could not send data to the Lineage Viewer web app:\n{0}".format(
                traceback.format_exc()
            )

    def call_lineage_api(self, method_name, payload):
        script_text = "window.ffeLineage && window.ffeLineage.{0}({1});".format(
            method_name,
            json_dumps(payload),
        )
        self.execute_script(script_text)

    def send_lineage_payload(self):
        if self.has_sent_lineage_payload:
            return
        self.has_sent_lineage_payload = True
        self.call_lineage_api("loadData", self.lineage_payload)

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
            self.send_lineage_payload()
            return

        if message_type == "closeWindow":
            self.Close()


# ____________________________________________________________________ MAIN
if doc is None:
    forms.alert(
        "Open a Revit model before running Lineage Viewer.",
        title="Lineage Viewer",
        exitscript=True,
    )

if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The Lineage Viewer web app was not found:\n{0}".format(PATH_INDEX),
        title="Lineage Viewer",
        exitscript=True,
    )

if not focus_existing_window():
    try:
        payload = build_lineage_payload(doc)
    except Exception as data_error:
        forms.alert(
            "Could not read linked items from the current project.\n\n{0}".format(safe_str(data_error)),
            title="Lineage Viewer",
            warn_icon=True,
            exitscript=True,
        )

    if len(payload.get("nodes", [])) <= 1:
        forms.alert(
            "No linked items were found in the active model.",
            title="Lineage Viewer",
            exitscript=True,
        )

    try:
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as webview_error:
        forms.alert(
            "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
            title="Lineage Viewer",
            exitscript=True,
        )

    window = LineageViewerWindow(WebView2, CoreWebView2CreationProperties, payload)
    WINDOW_REFS.append(window)
    window.Show()
