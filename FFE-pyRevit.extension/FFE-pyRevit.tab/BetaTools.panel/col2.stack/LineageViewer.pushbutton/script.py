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


EMPTY_GUID = "00000000-0000-0000-0000-000000000000"


def normalize_guid(value):
    text = safe_str(value).strip().strip("{}").lower()
    if not text or text == EMPTY_GUID:
        return ""
    return text


def call_optional_member(target, member_names):
    if target is None:
        return None

    for member_name in member_names:
        try:
            member = getattr(target, member_name)
            value = member() if callable(member) else member
            if value:
                return value
        except:
            pass
    return None


def model_path_to_user_visible(model_path):
    if not model_path:
        return ""

    try:
        return safe_str(ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)).strip()
    except:
        return ""


def normalize_identity_path(path_value):
    text = safe_str(path_value).strip()
    if not text:
        return ""
    try:
        text = os.path.normcase(os.path.normpath(text))
    except:
        pass
    return text.lower()


def element_id_key(element_id):
    if not element_id:
        return ""

    for attr_name in ["IntegerValue", "Value"]:
        try:
            value = getattr(element_id, attr_name)
            if value is not None:
                return safe_str(value)
        except:
            pass

    return safe_str(element_id)


def try_get_external_model_path(owner_doc, element_id):
    try:
        ext_ref = ExternalFileUtils.GetExternalFileReference(owner_doc, element_id)
        if ext_ref:
            return call_optional_member(ext_ref, ["GetAbsolutePath", "GetPath"])
    except:
        pass
    return None


def try_get_external_path(owner_doc, element_id):
    model_path = try_get_external_model_path(owner_doc, element_id)
    if model_path:
        return model_path_to_user_visible(model_path)
    return ""


def get_document_path(target_doc):
    if target_doc is None:
        return ""

    try:
        return safe_str(target_doc.PathName).strip()
    except:
        return ""


def get_document_central_path(target_doc):
    model_path = call_optional_member(target_doc, ["GetWorksharingCentralModelPath"])
    return model_path_to_user_visible(model_path)


def get_model_guid_from_model_path(model_path):
    guid_value = call_optional_member(
        model_path,
        ["GetModelGUID", "GetModelGuid", "ModelGUID", "ModelGuid"]
    )
    return normalize_guid(guid_value)


def get_model_guid_from_document(target_doc):
    guid_value = call_optional_member(
        target_doc,
        ["WorksharingCentralGUID", "WorksharingCentralGuid", "ModelGUID", "ModelGuid"]
    )
    guid_text = normalize_guid(guid_value)
    if guid_text:
        return guid_text, "Model GUID"

    cloud_model_path = call_optional_member(target_doc, ["GetCloudModelPath"])
    guid_text = get_model_guid_from_model_path(cloud_model_path)
    if guid_text:
        return guid_text, "Cloud Model GUID"

    return "", ""


def get_model_identity(target_doc=None, model_path=None, path_value="", fallback_label=""):
    model_guid = ""
    guid_source = ""

    path_text = safe_str(path_value).strip()
    if not path_text and model_path:
        path_text = model_path_to_user_visible(model_path)
    if not path_text and target_doc is not None:
        path_text = get_document_central_path(target_doc) or get_document_path(target_doc)

    if target_doc is not None:
        model_guid, guid_source = get_model_guid_from_document(target_doc)

    if not model_guid and model_path:
        model_guid = get_model_guid_from_model_path(model_path)
        if model_guid:
            guid_source = "Link ModelPath GUID"

    normalized_path = normalize_identity_path(path_text)
    if normalized_path:
        return {
            "modelGuid": model_guid,
            "identityKey": "path:{0}".format(normalized_path),
            "identitySource": "Model path",
            "path": path_text,
            "guidSource": guid_source,
        }

    if model_guid:
        return {
            "modelGuid": model_guid,
            "identityKey": "guid:{0}".format(model_guid),
            "identitySource": "Fallback Model GUID (missing model path)",
            "path": "",
            "guidSource": guid_source,
        }

    fallback_text = safe_str(fallback_label).strip()
    if fallback_text:
        return {
            "modelGuid": "",
            "identityKey": "name:{0}".format(fallback_text.lower()),
            "identitySource": "Fallback name (missing model path)",
            "path": "",
            "guidSource": "",
        }

    try:
        fallback_text = "DOCID:{0}".format(target_doc.GetHashCode())
    except:
        fallback_text = "DOC:{0}".format(id(target_doc)) if target_doc else "UNKNOWN"

    return {
        "modelGuid": "",
        "identityKey": fallback_text,
        "identitySource": "Fallback document id (missing model path)",
        "path": "",
        "guidSource": "",
    }


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
    def __init__(
        self,
        node_id,
        label,
        sublabel,
        kind,
        depth,
        parent_id=None,
        status="",
        model_guid="",
        identity_key="",
        identity_source="",
        identity_path=""
    ):
        self.id = node_id
        self.label = label
        self.sublabel = sublabel
        self.kind = kind
        self.depth = depth
        self.parent_id = parent_id
        self.status = status
        self.model_guid = model_guid
        self.identity_key = identity_key
        self.identity_source = identity_source
        self.identity_path = identity_path


class LineageGraph(object):
    def __init__(self):
        self.nodes = {}
        self.node_order = []
        self.edges = []
        self.node_counter = 0
        self.edge_counter = 0
        self.model_node_ids_by_identity = {}
        self.visited_model_keys = set()
        self.host_node_id = None
        self.host_identity_key = ""

    def next_id(self, prefix):
        self.node_counter += 1
        return "{0}_{1}".format(prefix, self.node_counter)

    def next_edge_id(self):
        self.edge_counter += 1
        return "edge_{0}".format(self.edge_counter)

    def iter_nodes(self):
        for node_id in self.node_order:
            yield self.nodes[node_id]

    def add_edge(
        self,
        from_id,
        to_id,
        kind="contains",
        style="normal",
        label="",
        status="",
        instance_id="",
        type_id=""
    ):
        self.edges.append({
            "id": self.next_edge_id(),
            "from": from_id,
            "to": to_id,
            "kind": safe_str(kind) or "contains",
            "style": safe_str(style) or "normal",
            "label": safe_str(label),
            "status": safe_str(status),
            "instanceId": safe_str(instance_id),
            "typeId": safe_str(type_id),
        })

    def add_node(
        self,
        label,
        sublabel,
        kind,
        depth,
        parent_id=None,
        status="",
        model_guid="",
        identity_key="",
        identity_source="",
        identity_path=""
    ):
        node_id = self.next_id(kind)
        self.nodes[node_id] = GraphNode(
            node_id,
            safe_str(label),
            safe_str(sublabel),
            safe_str(kind),
            int(depth),
            parent_id,
            safe_str(status),
            safe_str(model_guid),
            safe_str(identity_key),
            safe_str(identity_source),
            safe_str(identity_path),
        )
        self.node_order.append(node_id)
        if parent_id:
            self.add_edge(parent_id, node_id)
        return node_id

    def add_model_node(self, label, sublabel, kind, depth, identity_info, parent_id=None, status=""):
        identity_info = identity_info or {}
        identity_key = safe_str(identity_info.get("identityKey"))
        model_guid = safe_str(identity_info.get("modelGuid"))
        identity_source = safe_str(identity_info.get("identitySource"))
        identity_path = safe_str(identity_info.get("path"))

        if identity_key and identity_key in self.model_node_ids_by_identity:
            node_id = self.model_node_ids_by_identity[identity_key]
            node = self.nodes[node_id]
            try:
                if int(depth) < node.depth:
                    node.depth = int(depth)
            except:
                pass
            if kind == "host":
                node.kind = "host"
                node.status = safe_str(status) or node.status
            if not node.model_guid and model_guid:
                node.model_guid = model_guid
            if identity_source and not node.identity_source:
                node.identity_source = identity_source
            if identity_path and not node.identity_path:
                node.identity_path = identity_path
            return node_id, False

        node_id = self.add_node(
            label=label,
            sublabel=sublabel,
            kind=kind,
            depth=depth,
            parent_id=parent_id,
            status=status,
            model_guid=model_guid,
            identity_key=identity_key,
            identity_source=identity_source,
            identity_path=identity_path,
        )
        if identity_key:
            self.model_node_ids_by_identity[identity_key] = node_id
        return node_id, True


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


def collect_revit_link_instances(owner_doc):
    try:
        return FilteredElementCollector(owner_doc).OfClass(RevitLinkInstance).ToElements()
    except:
        return []


def collect_revit_link_types(owner_doc):
    try:
        return FilteredElementCollector(owner_doc).OfClass(RevitLinkType).ToElements()
    except:
        return []


def collect_link_instances_by_type(owner_doc):
    result = defaultdict(list)
    link_instances = collect_revit_link_instances(owner_doc)

    for link_instance in link_instances:
        try:
            type_key = element_id_key(link_instance.GetTypeId())
            if type_key:
                result[type_key].append(link_instance)
        except:
            pass

    return result


def get_first_loaded_link_document(link_instances):
    for link_instance in link_instances or []:
        try:
            link_doc = link_instance.GetLinkDocument()
            if link_doc is not None:
                return link_doc
        except:
            pass
    return None


def is_nested_revit_link_type(link_type):
    if link_type is None:
        return False

    for attr_name in ["IsNestedLink", "IsNested"]:
        try:
            value = getattr(link_type, attr_name)
            if callable(value):
                value = value()
            if bool(value):
                return True
        except:
            pass

    return False


def collect_revit_links(owner_doc, parent_node_id, depth, graph):
    link_types = collect_revit_link_types(owner_doc)
    instances_by_type = collect_link_instances_by_type(owner_doc)
    processed_type_ids = set()
    processed_link_count = [0]

    def process_link_group(link_type, type_id, link_instances):
        try:
            if link_type is None and not link_instances:
                return False

            if not type_id and link_type is not None:
                type_id = element_id_key(link_type.Id)
            if type_id:
                processed_type_ids.add(type_id)

            link_doc = get_first_loaded_link_document(link_instances)

            type_name = ""
            try:
                if link_type is not None:
                    type_name = safe_str(link_type.Name)
            except:
                pass
            if not type_name:
                try:
                    if link_instances:
                        type_name = safe_str(link_instances[0].Name)
                except:
                    pass
            if not type_name:
                type_name = "Revit Link"

            path_value = ""
            model_path = None
            if link_type is not None:
                model_path = try_get_external_model_path(owner_doc, link_type.Id)
                path_value = model_path_to_user_visible(model_path)
            if not path_value and link_doc is not None:
                try:
                    path_value = safe_str(link_doc.PathName)
                except:
                    pass

            status = safe_get_link_status(link_type) if link_type is not None else "Unknown"
            if link_doc is not None:
                status = "Loaded"
            model_label = ""
            try:
                if link_doc is not None:
                    model_label = safe_str(link_doc.Title)
            except:
                pass
            if not model_label:
                model_label = type_name
            if not model_label:
                model_label = "Revit Link"

            identity_info = get_model_identity(
                target_doc=link_doc,
                model_path=model_path,
                path_value=path_value,
                fallback_label=model_label,
            )
            node_kind = "host" if (
                graph.host_identity_key and
                identity_info.get("identityKey") == graph.host_identity_key
            ) else "revitlink"

            node_id, created = graph.add_model_node(
                label=model_label,
                sublabel=basename_or_value(path_value) if path_value else "No path available",
                kind=node_kind,
                depth=depth,
                parent_id=None,
                status=status,
                identity_info=identity_info,
            )

            if node_kind != "host":
                try:
                    node = graph.nodes[node_id]
                    if not node.parent_id:
                        node.parent_id = parent_node_id
                except:
                    pass

            edge_style = "normal"
            if (
                graph.host_identity_key and
                identity_info.get("identityKey") == graph.host_identity_key and
                parent_node_id != graph.host_node_id
            ):
                edge_style = "reciprocal"

            if link_instances:
                for link_instance in link_instances:
                    instance_name = ""
                    try:
                        instance_name = safe_str(link_instance.Name)
                    except:
                        pass
                    graph.add_edge(
                        from_id=parent_node_id,
                        to_id=node_id,
                        kind="revitlink",
                        style=edge_style,
                        label=instance_name or type_name,
                        status=status,
                        instance_id=element_id_key(link_instance.Id),
                        type_id=type_id,
                    )
            else:
                graph.add_edge(
                    from_id=parent_node_id,
                    to_id=node_id,
                    kind="revitlink",
                    style=edge_style,
                    label=type_name,
                    status=status,
                    instance_id="",
                    type_id=type_id,
                )

            model_key = identity_info.get("identityKey")
            if link_doc is not None and model_key:
                if model_key not in graph.visited_model_keys:
                    graph.visited_model_keys.add(model_key)
                    collect_doc_contents(link_doc, node_id, depth + 1, graph)
            return True
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass
        return False

    for link_type in link_types:
        try:
            type_id = element_id_key(link_type.Id)
            if is_nested_revit_link_type(link_type):
                continue
            if process_link_group(link_type, type_id, instances_by_type.get(type_id, [])):
                processed_link_count[0] += 1
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass

    for type_id, link_instances in instances_by_type.items():
        if type_id in processed_type_ids:
            continue

        link_type = None
        try:
            if link_instances:
                link_type = owner_doc.GetElement(link_instances[0].GetTypeId())
        except:
            link_type = None

        if link_type is not None and is_nested_revit_link_type(link_type) and processed_link_count[0] > 0:
            continue

        if process_link_group(link_type, type_id, link_instances):
            processed_link_count[0] += 1


def collect_doc_contents(owner_doc, parent_node_id, depth, graph):
    collect_revit_links(owner_doc, parent_node_id, depth, graph)
    collect_imports(owner_doc, parent_node_id, depth, graph)
    collect_images_and_pdfs(owner_doc, parent_node_id, depth, graph)
    collect_pointclouds(owner_doc, parent_node_id, depth, graph)


def build_lineage_payload(active_doc):
    graph = LineageGraph()
    host_path = safe_str(active_doc.PathName) if active_doc.PathName else ""
    host_identity = get_model_identity(
        target_doc=active_doc,
        path_value=host_path,
        fallback_label=active_doc.Title,
    )
    host_node_id, created = graph.add_model_node(
        label=active_doc.Title,
        sublabel=basename_or_value(active_doc.PathName) if active_doc.PathName else "Unsaved model",
        kind="host",
        depth=0,
        parent_id=None,
        status="Active Model",
        identity_info=host_identity,
    )

    graph.host_node_id = host_node_id
    graph.host_identity_key = host_identity.get("identityKey")
    if graph.host_identity_key:
        graph.visited_model_keys.add(graph.host_identity_key)
    collect_doc_contents(active_doc, host_node_id, 1, graph)

    counts = defaultdict(int)
    node_payloads = []
    for edge in graph.edges:
        if edge.get("kind") == "revitlink":
            counts["revitlink"] += 1

    for node in graph.iter_nodes():
        if node.kind not in ["host", "revitlink"]:
            counts[node.kind] += 1
        node_payloads.append({
            "id": node.id,
            "label": node.label,
            "sublabel": node.sublabel,
            "kind": node.kind,
            "depth": node.depth,
            "parentId": node.parent_id,
            "status": node.status,
            "modelGuid": node.model_guid,
            "identityKey": node.identity_key,
            "identitySource": node.identity_source,
            "identityPath": node.identity_path,
        })

    return {
        "model": {
            "title": safe_str(active_doc.Title),
            "path": safe_str(active_doc.PathName) if active_doc.PathName else "Unsaved model",
            "modelGuid": host_identity.get("modelGuid", ""),
            "identityKey": host_identity.get("identityKey", ""),
            "identitySource": host_identity.get("identitySource", ""),
            "identityPath": host_identity.get("path", ""),
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

        # try:
        #     WindowInteropHelper(self).Owner = __revit__.MainWindowHandle
        #     self.ShowInTaskbar = True
        # except:
        #     pass

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

    if len(payload.get("nodes", [])) <= 1 and not payload.get("edges", []):
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
